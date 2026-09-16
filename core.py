"""Core implementation for the Makro listing uploader.

This module intentionally uses only the Python standard library.  PySide6 is
kept in app.py so the workbook parser, database and HTTP client can be tested
without a GUI.
"""

from __future__ import annotations

import hashlib
import json
import os
import posixpath
import random
import re
import sqlite3
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Iterator, Sequence
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
DEFAULT_HEADERS = ["商品名称", "品牌", "价格(R)", "原价(R)", "优惠价(R)", "配送", "预计天数", "详情页地址"]


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def excel_column_name(index: int) -> str:
    """Return a 1-based Excel column name."""
    result = []
    while index:
        index, rem = divmod(index - 1, 26)
        result.append(chr(65 + rem))
    return "".join(reversed(result))


def parse_dimension(ref: str) -> tuple[int, int]:
    match = re.search(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref or "")
    if not match:
        match = re.search(r"([A-Z]+)(\d+)", ref or "")
        if not match:
            return 0, 0
        return int(match.group(2)), int(match.group(2))

    def col_number(value: str) -> int:
        total = 0
        for char in value:
            total = total * 26 + ord(char) - 64
        return total

    return int(match.group(2)), int(match.group(4))


def _xml_local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


class XlsxReader:
    """Streaming reader for the simple .xlsx files used by this workflow."""

    def __init__(self, path: str | os.PathLike[str]):
        self.path = str(path)
        self._sheet_paths: dict[str, str] | None = None
        self._dimensions: dict[str, tuple[int, int]] | None = None

    def _load_workbook_metadata(self) -> None:
        if self._sheet_paths is not None:
            return
        with zipfile.ZipFile(self.path) as archive:
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            rel_map = {
                rel.attrib["Id"]: rel.attrib["Target"]
                for rel in relationships
                if _xml_local(rel.tag) == "Relationship"
            }
            sheet_paths: dict[str, str] = {}
            dimensions: dict[str, tuple[int, int]] = {}
            for sheet in workbook.iter():
                if _xml_local(sheet.tag) != "sheet":
                    continue
                name = sheet.attrib["name"]
                relation_id = sheet.attrib[f"{{{NS_REL}}}id"]
                # Relationship targets may be absolute package paths
                # (``/xl/worksheets/sheet1.xml``) or paths relative to the
                # ``xl`` directory (``worksheets/sheet1.xml``).  Strip the
                # leading slash before checking the prefix, otherwise an
                # absolute target becomes the invalid ``xl/xl/...`` path.
                target = rel_map[relation_id].lstrip("/")
                if target.startswith("../"):
                    target = posixpath.normpath(posixpath.join("xl", target))
                elif not target.startswith("xl/"):
                    target = "xl/" + target
                sheet_paths[name] = target
                try:
                    # Do not read a multi-hundred-megabyte worksheet into
                    # memory just to discover its row count.  The dimension
                    # element is near the beginning of normal XLSX sheets.
                    with archive.open(target) as sheet_stream:
                        dimensions[name] = (0, 0)
                        for _, node in ET.iterparse(sheet_stream, events=("end",)):
                            if _xml_local(node.tag) == "dimension":
                                dimensions[name] = parse_dimension(node.attrib.get("ref", ""))
                                break
                except (KeyError, ET.ParseError):
                    dimensions[name] = (0, 0)
            self._sheet_paths = sheet_paths
            self._dimensions = dimensions

    def sheet_names(self) -> list[str]:
        self._load_workbook_metadata()
        return list(self._sheet_paths or {})

    def dimension(self, sheet_name: str) -> tuple[int, int]:
        self._load_workbook_metadata()
        return (self._dimensions or {}).get(sheet_name, (0, 0))

    def _sheet_path(self, sheet_name: str) -> str:
        self._load_workbook_metadata()
        try:
            return (self._sheet_paths or {})[sheet_name]
        except KeyError as exc:
            raise ValueError(f"工作表不存在: {sheet_name}") from exc

    def shared_strings(self) -> list[str]:
        with zipfile.ZipFile(self.path) as archive:
            try:
                stream = archive.open("xl/sharedStrings.xml")
            except KeyError:
                return []
            strings: list[str] = []
            for _, element in ET.iterparse(stream, events=("end",)):
                if _xml_local(element.tag) == "si":
                    strings.append("".join(node.text or "" for node in element.iter() if _xml_local(node.tag) == "t"))
                    element.clear()
            return strings

    def rows(self, sheet_name: str, start_row: int = 1, max_rows: int | None = None) -> Iterator[tuple[int, list[str]]]:
        strings = self.shared_strings()
        sheet_path = self._sheet_path(sheet_name)
        emitted = 0
        with zipfile.ZipFile(self.path) as archive, archive.open(sheet_path) as stream:
            for _, element in ET.iterparse(stream, events=("end",)):
                if _xml_local(element.tag) != "row":
                    continue
                row_number = int(element.attrib.get("r", emitted + 1))
                if row_number >= start_row:
                    cells: dict[int, str] = {}
                    fallback_col = 1
                    for cell in list(element):
                        if _xml_local(cell.tag) != "c":
                            continue
                        cell_ref = cell.attrib.get("r", "")
                        col_match = re.match(r"([A-Z]+)", cell_ref)
                        if col_match:
                            col = 0
                            for char in col_match.group(1):
                                col = col * 26 + ord(char) - 64
                        else:
                            col = fallback_col
                        fallback_col = col + 1
                        value_node = next((node for node in list(cell) if _xml_local(node.tag) == "v"), None)
                        inline_node = next((node for node in list(cell) if _xml_local(node.tag) == "is"), None)
                        if inline_node is not None:
                            value = "".join(node.text or "" for node in inline_node.iter() if _xml_local(node.tag) == "t")
                        elif value_node is None:
                            value = ""
                        else:
                            value = value_node.text or ""
                            if cell.attrib.get("t") == "s" and value:
                                try:
                                    value = strings[int(value)]
                                except (ValueError, IndexError):
                                    value = ""
                        cells[col] = value
                    last_col = max(cells, default=0)
                    values = [cells.get(col, "") for col in range(1, last_col + 1)]
                    yield row_number, values
                    emitted += 1
                    if max_rows is not None and emitted >= max_rows:
                        return
                element.clear()


def find_header(headers: Sequence[str], *terms: str, fallback: int | None = None) -> int:
    normalized = [normalize_text(header).lower() for header in headers]
    for index, header in enumerate(normalized):
        for term in terms:
            needle = term.lower()
            # A generic substring such as "url" incorrectly matches words
            # like "curler" in a product title.  Generic English tokens are
            # matched as complete labels; Chinese labels can still use a
            # contains match for common suffixes such as (R).
            if needle in {"url", "product url", "mrp", "selling", "sale", "days"}:
                matched = header == needle or header.replace("_", " ") == needle
            else:
                matched = needle in header
            if matched:
                return index
    if fallback is not None and fallback < len(headers):
        return fallback
    raise ValueError(f"找不到字段: {'/'.join(terms)}")


def has_header_row(values: Sequence[str]) -> bool:
    normalized = {normalize_text(value).lower() for value in values}
    known = {"商品名称", "品牌", "价格(r)", "原价(r)", "优惠价(r)", "配送", "预计天数", "详情页地址"}
    if normalized.intersection(known):
        return True
    return any(value in {"product url", "url", "brand", "mrp", "selling price"} for value in normalized)


def extract_product_id(url: str) -> str:
    parsed = urllib.parse.urlparse(normalize_text(url))
    values = urllib.parse.parse_qs(parsed.query).get("pid", [])
    if values and values[0].strip():
        return values[0].strip()
    match = re.search(r"[?&]pid=([^&#]+)", url, flags=re.I)
    return urllib.parse.unquote(match.group(1)).strip() if match else ""


def number_string(value: str, field_name: str) -> str:
    text = normalize_text(value)
    if not text:
        raise ValueError(f"{field_name}为空")
    try:
        number = float(text.replace(",", ""))
    except ValueError as exc:
        raise ValueError(f"{field_name}不是数字: {text}") from exc
    if number < 0:
        raise ValueError(f"{field_name}不能为负数")
    if number.is_integer():
        return str(int(number))
    return format(number, "g")


@dataclass(frozen=True)
class ListingRow:
    row_number: int
    values: list[str]
    product_id: str
    original_price: str
    selling_price: str
    shipping_days: str
    fingerprint: str


def make_listing_row(
    row_number: int,
    values: list[str],
    mapping: dict[str, int],
    shipping_days_override: str | None = None,
) -> ListingRow:
    padded = values + [""] * (max(mapping.values(), default=0) + 1 - len(values))
    url = padded[mapping["url"]]
    product_id = extract_product_id(url)
    if not product_id:
        raise ValueError("详情页地址中没有找到 pid")
    original_price = number_string(padded[mapping["original_price"]], "原价")
    discount = float(number_string(padded[mapping["selling_price"]], "优惠价")) - 1
    if discount < 0:
        raise ValueError("优惠价减1后不能为负数")
    selling_price = str(int(discount)) if discount.is_integer() else format(discount, "g")
    shipping_days = number_string(
        shipping_days_override if shipping_days_override is not None else padded[mapping["shipping_days"]],
        "Pick Pack SLA" if shipping_days_override is not None else "预计天数",
    )
    fingerprint = hashlib.sha256(json.dumps(values, ensure_ascii=False, separators=(",", ":")).encode("utf-8")).hexdigest()
    return ListingRow(row_number, values, product_id, original_price, selling_price, shipping_days, fingerprint)


def infer_mapping(headers: Sequence[str]) -> dict[str, int]:
    return {
        "url": find_header(headers, "详情页地址", "product url", "url", fallback=7),
        "original_price": find_header(headers, "原价", "mrp", fallback=3),
        "selling_price": find_header(headers, "优惠价", "selling", "sale", fallback=4),
        "shipping_days": find_header(headers, "预计天数", "shipping", "days", fallback=6),
    }


def split_sku_seed(text: str) -> tuple[str, int]:
    normalized = normalize_text(text).upper()
    match = re.fullmatch(r"([A-Z0-9_-]*?)(\d+)", normalized)
    if not match:
        raise ValueError("SKU 起始值必须以数字结尾，例如 AA1001")
    prefix, number_text = match.groups()
    if not prefix and not number_text:
        raise ValueError("SKU 起始值不能为空")
    return prefix, int(number_text)


class UploadDatabase:
    def __init__(self, path: str | os.PathLike[str] = "makro_uploader.sqlite3"):
        self.path = str(path)
        self.connection = sqlite3.connect(self.path, check_same_thread=False)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA journal_mode=WAL")
        self.connection.execute("PRAGMA synchronous=FULL")
        self.connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS runs (
                run_id TEXT PRIMARY KEY,
                source_file TEXT NOT NULL,
                sheet_name TEXT NOT NULL,
                headers_json TEXT NOT NULL,
                started_at TEXT NOT NULL,
                finished_at TEXT,
                status TEXT NOT NULL DEFAULT 'running'
            );
            CREATE TABLE IF NOT EXISTS uploads (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id TEXT NOT NULL,
                source_file TEXT NOT NULL,
                sheet_name TEXT NOT NULL,
                excel_row INTEGER NOT NULL,
                row_fingerprint TEXT NOT NULL,
                source_values_json TEXT NOT NULL,
                sku TEXT NOT NULL,
                product_id TEXT,
                status TEXT NOT NULL,
                message TEXT NOT NULL,
                http_status INTEGER,
                response_json TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(run_id) REFERENCES runs(run_id)
            );
            CREATE INDEX IF NOT EXISTS idx_uploads_fingerprint ON uploads(row_fingerprint, created_at);
            CREATE INDEX IF NOT EXISTS idx_uploads_run ON uploads(run_id, id);
            """
        )
        self.connection.commit()
        self._lock = threading.RLock()

    def get_setting(self, key: str, default: str = "") -> str:
        row = self.connection.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return str(row[0]) if row else default

    def set_setting(self, key: str, value: str) -> None:
        with self._lock, self.connection:
            self.connection.execute(
                "INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, str(value)),
            )

    def delete_setting(self, key: str) -> None:
        with self._lock, self.connection:
            self.connection.execute("DELETE FROM settings WHERE key=?", (key,))

    def create_run(self, source_file: str, sheet_name: str, headers: Sequence[str]) -> str:
        run_id = uuid.uuid4().hex
        with self._lock, self.connection:
            self.connection.execute(
                "INSERT INTO runs(run_id,source_file,sheet_name,headers_json,started_at) VALUES(?,?,?,?,?)",
                (run_id, source_file, sheet_name, json.dumps(list(headers), ensure_ascii=False), utc_now()),
            )
        return run_id

    def finish_run(self, run_id: str, status: str) -> None:
        with self._lock, self.connection:
            self.connection.execute("UPDATE runs SET finished_at=?, status=? WHERE run_id=?", (utc_now(), status, run_id))

    def allocate_sku(self, seed_text: str) -> str:
        prefix, seed = split_sku_seed(seed_text)
        with self._lock, self.connection:
            active_prefix = self.get_setting("active_prefix")
            last_number_text = self.get_setting("last_number")
            if not active_prefix:
                last_number = seed - 1
            elif active_prefix != prefix:
                # A changed prefix is interpreted as the last used value.  Thus
                # entering AACC1021 starts the new series at AACC1022.
                last_number = seed
            else:
                try:
                    last_number = int(last_number_text)
                except ValueError:
                    last_number = seed - 1
                # Allow a user to advance the same prefix manually.
                if seed > last_number + 1:
                    last_number = seed - 1
            next_number = last_number + 1
            self.connection.execute(
                "INSERT INTO settings(key,value) VALUES('active_prefix',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (prefix,),
            )
            self.connection.execute(
                "INSERT INTO settings(key,value) VALUES('last_number',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (str(next_number),),
            )
            return f"{prefix}{next_number}"

    def latest_sku_for_fingerprint(self, fingerprint: str) -> str | None:
        row = self.connection.execute(
            "SELECT sku FROM uploads WHERE row_fingerprint=? ORDER BY id DESC LIMIT 1", (fingerprint,)
        ).fetchone()
        return str(row[0]) if row else None

    def latest_upload_for_fingerprint(self, fingerprint: str) -> sqlite3.Row | None:
        return self.connection.execute(
            "SELECT * FROM uploads WHERE row_fingerprint=? ORDER BY id DESC LIMIT 1", (fingerprint,)
        ).fetchone()

    def add_upload(
        self,
        run_id: str,
        source_file: str,
        sheet_name: str,
        listing_row: ListingRow,
        sku: str,
        status: str,
        message: str,
        http_status: int | None = None,
        response: Any = None,
    ) -> None:
        with self._lock, self.connection:
            self.connection.execute(
                """
                INSERT INTO uploads(
                    run_id,source_file,sheet_name,excel_row,row_fingerprint,source_values_json,
                    sku,product_id,status,message,http_status,response_json,created_at
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    run_id,
                    source_file,
                    sheet_name,
                    listing_row.row_number,
                    listing_row.fingerprint,
                    json.dumps(listing_row.values, ensure_ascii=False),
                    sku,
                    listing_row.product_id,
                    status,
                    message,
                    http_status,
                    json.dumps(response, ensure_ascii=False) if response is not None else None,
                    utc_now(),
                ),
            )

    def run_rows(self, run_id: str) -> list[sqlite3.Row]:
        return list(self.connection.execute("SELECT * FROM uploads WHERE run_id=? ORDER BY id", (run_id,)))

    def all_runs(self) -> list[sqlite3.Row]:
        return list(self.connection.execute("SELECT * FROM runs ORDER BY started_at DESC"))

    def close(self) -> None:
        self.connection.close()


def build_payload(listing_row: ListingRow, sku: str, seller_id: str) -> dict[str, Any]:
    attrs = {
        "sku_id": [{"value": sku, "qualifier": ""}],
        "listing_status": [{"value": "ACTIVE", "qualifier": ""}],
        "mrp": [{"value": listing_row.original_price, "qualifier": "INR"}],
        "flipkart_selling_price": [{"value": listing_row.selling_price, "qualifier": "INR"}],
        "minimum_order_quantity": [{"value": "1", "qualifier": ""}],
        "max_order_quantity_allowed": [{"value": "99", "qualifier": ""}],
        "service_profile": [{"value": "NON_FBF", "qualifier": ""}],
        "shipping_days": [{"value": listing_row.shipping_days, "qualifier": "DAY"}],
        "forbid_shipping": [{"qualifier": "", "value": "none"}],
        "country_of_origin": [{"value": "CN", "qualifier": ""}],
        "manufacturer_details": [{"value": "N/A", "qualifier": ""}],
        "packer_details": [{"value": "N/A", "qualifier": ""}],
        "importer_details": [{"value": "N/A", "qualifier": ""}],
    }
    package = {
        "id": {"value": "packages-0"},
        "length": {"value": "1", "qualifier": "CM"},
        "breadth": {"value": "1", "qualifier": "CM"},
        "height": {"value": "1", "qualifier": "CM"},
        "weight": {"value": "1", "qualifier": "KG"},
        "sku_id": {"value": sku, "qualifier": ""},
    }
    return {
        "bulkRequests": [
            {
                "attributeValues": attrs,
                "context": {"ignore_warnings": False},
                "productId": listing_row.product_id,
                "skuId": sku,
                "packages": [package],
            }
        ],
        "sellerId": seller_id,
    }


@dataclass
class HttpResult:
    ok: bool
    message: str
    http_status: int | None
    response: Any
    sku_conflict: bool = False


class MakroClient:
    def __init__(self, seller_id: str, cookie: str, csrf_token: str = "", base_url: str = "https://seller.makro.co.za"):
        self.seller_id = normalize_text(seller_id)
        self.cookie = re.sub(r"[\r\n]+", "", normalize_text(cookie))
        if self.cookie.lower().startswith("cookie:"):
            self.cookie = self.cookie.split(":", 1)[1].strip()
        self.csrf_token = normalize_text(csrf_token)
        if self.csrf_token.lower().startswith("fk-csrf-token:"):
            self.csrf_token = self.csrf_token.split(":", 1)[1].strip()
        self.base_url = base_url.rstrip("/")

    def upload(self, listing_row: ListingRow, sku: str, timeout: float = 40) -> HttpResult:
        url = f"{self.base_url}/napi/listing/create-update-listings?{urllib.parse.urlencode({'sellerId': self.seller_id})}"
        request = urllib.request.Request(
            url,
            data=json.dumps(build_payload(listing_row, sku, self.seller_id), ensure_ascii=False).encode("utf-8"),
            method="POST",
            headers={
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "Accept-Language": "zh-CN,zh;q=0.9",
                "Content-Type": "application/json",
                "Cookie": self.cookie,
                "Origin": self.base_url,
                "Referer": f"{self.base_url}/index.html",
                "Sourceid": "ui.latch-on",
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/153 Safari/537.36",
                "X-Requested-With": "XMLHttpRequest",
                **({"fk-csrf-token": self.csrf_token} if self.csrf_token else {}),
            },
        )
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                raw = response.read().decode("utf-8", errors="replace")
                status = int(response.status)
        except urllib.error.HTTPError as exc:
            raw = exc.read().decode("utf-8", errors="replace")
            return self._interpret(raw, int(exc.code))
        except Exception as exc:  # network timeout, DNS, TLS, etc.
            return HttpResult(False, f"网络请求失败: {exc}", None, None)
        return self._interpret(raw, status)

    @staticmethod
    def _interpret(raw: str, status: int) -> HttpResult:
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            message = f"HTTP {status}: {raw[:500]}" if raw else f"HTTP {status}"
            return HttpResult(False, message, status, raw[:4000], sku_conflict=status == 409)
        result = data.get("result", {}) if isinstance(data, dict) else {}
        bulk = result.get("bulkResponse", []) if isinstance(result, dict) else []
        item = bulk[0] if bulk and isinstance(bulk[0], dict) else {}
        status_text = str(result.get("status", "")).lower() if isinstance(result, dict) else ""
        item_status = str(item.get("status", "")).lower()
        ok = status < 400 and (status_text == "success" or item_status in {"created", "updated", "success"})
        if ok:
            return HttpResult(True, item_status or "success", status, data)
        errors: list[str] = []
        for error in item.get("globalErrors", []) if isinstance(item, dict) else []:
            if isinstance(error, dict):
                errors.append(str(error.get("message") or error.get("errorMessage") or error))
            else:
                errors.append(str(error))
        for field, values in (item.get("attributeErrors", {}) or {}).items() if isinstance(item, dict) else []:
            for error in values if isinstance(values, list) else [values]:
                if isinstance(error, dict):
                    errors.append(f"{field}: {error.get('message') or error.get('errorMessage') or error}")
                else:
                    errors.append(f"{field}: {error}")
        if not errors:
            errors.append(f"HTTP {status}: {json.dumps(data, ensure_ascii=False)[:700]}")
        message = "；".join(errors)
        message_lower = message.lower()
        sku_conflict = status == 409 or bool(
            re.search(
                r"sku.{0,80}(already|exist|duplicate|unique|taken)|"
                r"(already|exist|duplicate|unique|taken).{0,80}sku",
                message_lower,
            )
        )
        return HttpResult(False, message, status, data, sku_conflict=sku_conflict)


def export_failed_xlsx(db: UploadDatabase, run_id: str, output_path: str | os.PathLike[str]) -> int:
    """Write failed rows for retry, retaining source columns and appending 错误."""
    run = db.connection.execute("SELECT * FROM runs WHERE run_id=?", (run_id,)).fetchone()
    if not run:
        raise ValueError("找不到上传批次")
    headers = json.loads(run["headers_json"])
    rows = [row for row in db.run_rows(run_id) if row["status"] == "failure"]
    if not rows:
        raise ValueError("本批次没有失败记录")
    output_path = str(output_path)
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    max_col = len(headers) + 1
    last_col = excel_column_name(max_col)
    sheet_fd, sheet_path = tempfile.mkstemp(prefix="makro-failed-", suffix=".xml")
    os.close(sheet_fd)
    try:
        with open(sheet_path, "w", encoding="utf-8", newline="") as sheet:
            sheet.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
            sheet.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
            sheet.write(f'<dimension ref="A1:{last_col}{len(rows) + 1}"/>')
            sheet.write("<sheetData>")
            _write_inline_row(sheet, 1, list(headers) + ["错误"])
            for index, row in enumerate(rows, start=2):
                values = json.loads(row["source_values_json"])
                _write_inline_row(sheet, index, values + [f"SKU {row['sku']}: {row['message']}"])
            sheet.write("</sheetData></worksheet>")

        content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>'''
        root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
        workbook = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="失败重试" sheetId="1" r:id="rId1"/></sheets></workbook>'''
        workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", content_types)
            archive.writestr("_rels/.rels", root_rels)
            archive.writestr("xl/workbook.xml", workbook)
            archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
            archive.write(sheet_path, "xl/worksheets/sheet1.xml")
    finally:
        try:
            os.unlink(sheet_path)
        except FileNotFoundError:
            pass
    return len(rows)


def _write_inline_row(stream: Any, row_number: int, values: Sequence[Any]) -> None:
    stream.write(f'<row r="{row_number}">')
    for index, value in enumerate(values, start=1):
        text = normalize_text(value)
        ref = f"{excel_column_name(index)}{row_number}"
        stream.write(f'<c r="{ref}" t="inlineStr"><is><t>{escape(text)}</t></is></c>')
    stream.write("</row>")


def wait_with_stop(seconds: float, stop_event: threading.Event) -> bool:
    return not stop_event.wait(max(0.0, seconds))
