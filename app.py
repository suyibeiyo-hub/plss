from __future__ import annotations

import json
import os
import random
import sys
import threading
from pathlib import Path

from PySide6.QtCore import QObject, QThread, QTimer, Qt, Signal, Slot
from PySide6.QtGui import QAction, QCloseEvent
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core import (
    MakroClient,
    UploadDatabase,
    XlsxReader,
    export_failed_xlsx,
    infer_mapping,
    DEFAULT_HEADERS,
    has_header_row,
    make_listing_row,
    split_sku_seed,
    wait_with_stop,
)


APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "makro_uploader.sqlite3"


class UploadWorker(QObject):
    row_result = Signal(dict)
    progress = Signal(int, int)
    status = Signal(str)
    finished = Signal(str, int, int)
    failed = Signal(str)

    def __init__(self, config: dict):
        super().__init__()
        self.config = config
        self.stop_event = threading.Event()
        self.db = UploadDatabase(DB_PATH)

    def stop(self) -> None:
        self.stop_event.set()

    @Slot()
    def run(self) -> None:
        run_id = ""
        success_count = 0
        failure_count = 0
        try:
            path = self.config["file_path"]
            reader = XlsxReader(path)
            first_row = next(reader.rows(self.config["sheet_name"], start_row=1, max_rows=1), (1, []))
            raw_headers = first_row[1]
            header_present = has_header_row(raw_headers)
            has_error_column = bool(raw_headers and raw_headers[-1].strip() == "错误")
            if header_present:
                headers = raw_headers[:-1] if has_error_column else raw_headers
                data_start_row = 2
            else:
                headers = DEFAULT_HEADERS[:max(8, len(raw_headers))]
                data_start_row = 1
            mapping = infer_mapping(headers)
            shipping_days_override = self.config.get("shipping_days_override")
            run_id = self.db.create_run(path, self.config["sheet_name"], headers)
            client = MakroClient(self.config["seller_id"], self.config["cookie"], self.config["csrf_token"])
            total = self.config["count"]
            self.status.emit(f"开始上传，共 {total} 条。")
            processed = 0
            for row_number, raw_values in reader.rows(self.config["sheet_name"], start_row=data_start_row):
                if processed >= total or self.stop_event.is_set():
                    break
                processed += 1
                values = raw_values[:-1] if (header_present and has_error_column) else raw_values
                status = "failure"
                message = ""
                sku = ""
                product_id = ""
                try:
                    listing_row = make_listing_row(row_number, values, mapping, shipping_days_override)
                    product_id = listing_row.product_id
                    sku = self.db.latest_sku_for_fingerprint(listing_row.fingerprint) or self.db.allocate_sku(self.config["sku_seed"])
                    self.status.emit(f"第 {processed}/{total} 条，SKU {sku}，等待请求")
                    delay = random.uniform(self.config["delay_min"], self.config["delay_max"])
                    if not wait_with_stop(delay, self.stop_event):
                        break
                    result = client.upload(listing_row, sku)
                    status = "success" if result.ok else "failure"
                    message = result.message
                    if result.ok:
                        success_count += 1
                    else:
                        failure_count += 1
                    self.db.add_upload(run_id, path, self.config["sheet_name"], listing_row, sku, status, message, result.http_status, result.response)
                    self.row_result.emit({"row": row_number, "sku": sku, "product_id": product_id, "status": status, "message": message})
                except Exception as exc:
                    failure_count += 1
                    message = str(exc)
                    values_for_log = values
                    # Invalid rows do not have a ListingRow yet, so create a
                    # lightweight record with the available raw values.
                    from core import ListingRow

                    fallback = ListingRow(row_number, values_for_log, product_id, "", "", "", "invalid-" + str(row_number))
                    sku = sku or self.db.allocate_sku(self.config["sku_seed"])
                    self.db.add_upload(run_id, path, self.config["sheet_name"], fallback, sku, "failure", message)
                    self.row_result.emit({"row": row_number, "sku": sku, "product_id": product_id, "status": "failure", "message": message})
                self.progress.emit(processed, total)
            final_status = "stopped" if self.stop_event.is_set() else "finished"
            self.db.finish_run(run_id, final_status)
            self.finished.emit(run_id, success_count, failure_count)
        except Exception as exc:
            if run_id:
                self.db.finish_run(run_id, "failed")
            self.failed.emit(str(exc))
        finally:
            self.db.close()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Makro 商品发布工具")
        self.resize(1040, 760)
        self.db = UploadDatabase(DB_PATH)
        self.reader: XlsxReader | None = None
        self.current_run_id = ""
        self.worker: UploadWorker | None = None
        self.thread: QThread | None = None
        self._build_ui()
        self._load_saved_settings()

    def _build_ui(self) -> None:
        central = QWidget()
        root = QVBoxLayout(central)
        root.setSpacing(10)

        connection_box = QGroupBox("连接设置")
        connection = QFormLayout(connection_box)
        self.cookie_edit = QPlainTextEdit()
        self.cookie_edit.setPlaceholderText("粘贴浏览器 Cookie 请求头的完整内容，例如 key=value; key2=value2")
        self.cookie_edit.setFixedHeight(72)
        connection.addRow("Cookie *", self.cookie_edit)
        self.csrf_edit = QLineEdit()
        self.csrf_edit.setPlaceholderText("粘贴浏览器请求头 fk-csrf-token 的值")
        self.csrf_edit.setEchoMode(QLineEdit.EchoMode.Password)
        connection.addRow("CSRF Token *", self.csrf_edit)
        self.seller_edit = QLineEdit()
        self.seller_edit.setPlaceholderText("填写当前登录账号的 Seller ID")
        connection.addRow("Seller ID *", self.seller_edit)
        root.addWidget(connection_box)

        file_box = QGroupBox("Excel 与 SKU")
        file_grid = QGridLayout(file_box)
        self.file_edit = QLineEdit()
        self.file_button = QPushButton("选择 Excel")
        self.file_button.clicked.connect(self.choose_file)
        file_grid.addWidget(QLabel("目录 Excel *"), 0, 0)
        file_grid.addWidget(self.file_edit, 0, 1)
        file_grid.addWidget(self.file_button, 0, 2)
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.refresh_estimate)
        file_grid.addWidget(QLabel("工作表"), 1, 0)
        file_grid.addWidget(self.sheet_combo, 1, 1)
        self.sku_seed_edit = QLineEdit("AA1001")
        self.sku_seed_edit.textChanged.connect(self.refresh_estimate)
        file_grid.addWidget(QLabel("SKU 起始值"), 2, 0)
        file_grid.addWidget(self.sku_seed_edit, 2, 1)
        self.count_spin = QSpinBox()
        self.count_spin.setRange(1, 2_000_000)
        self.count_spin.valueChanged.connect(self.refresh_estimate)
        file_grid.addWidget(QLabel("本轮上传数量"), 3, 0)
        file_grid.addWidget(self.count_spin, 3, 1)
        self.estimate_label = QLabel("请选择 Excel")
        self.estimate_label.setWordWrap(True)
        file_grid.addWidget(self.estimate_label, 4, 0, 1, 3)
        root.addWidget(file_box)

        options_box = QGroupBox("固定字段")
        options = QFormLayout(options_box)
        fixed = QLabel("状态 Active；最小/最大数量 1/99；仓配 NON_FBF；尺寸 1×1×1 cm；重量 1 kg；原产国 China；制造商/包装商/进口商 N/A")
        fixed.setWordWrap(True)
        options.addRow("固定值", fixed)
        shipping_row = QHBoxLayout()
        self.shipping_mode = QComboBox()
        self.shipping_mode.addItem("GUI 统一值", "gui")
        self.shipping_mode.addItem("使用 Excel G 列", "excel")
        self.shipping_days_spin = QSpinBox()
        self.shipping_days_spin.setRange(1, 365)
        self.shipping_days_spin.setValue(32)
        self.shipping_mode.currentIndexChanged.connect(
            lambda _index: self.shipping_days_spin.setEnabled(self.shipping_mode.currentData() == "gui")
        )
        shipping_row.addWidget(self.shipping_mode)
        shipping_row.addWidget(self.shipping_days_spin)
        shipping_row.addWidget(QLabel("DAY"))
        shipping_row.addStretch()
        options.addRow("Pick Pack SLA", shipping_row)
        root.addWidget(options_box)

        action_row = QHBoxLayout()
        self.start_button = QPushButton("开始上传")
        self.start_button.setMinimumHeight(36)
        self.start_button.clicked.connect(self.start_upload)
        self.stop_button = QPushButton("停止")
        self.stop_button.setEnabled(False)
        self.stop_button.clicked.connect(self.stop_upload)
        self.export_button = QPushButton("导出本批次失败 Excel")
        self.export_button.setEnabled(False)
        self.export_button.clicked.connect(self.export_failures)
        action_row.addWidget(self.start_button); action_row.addWidget(self.stop_button); action_row.addWidget(self.export_button); action_row.addStretch()
        root.addLayout(action_row)
        self.progress_bar = QProgressBar(); self.progress_bar.setValue(0)
        root.addWidget(self.progress_bar)
        self.status_label = QLabel("就绪")
        root.addWidget(self.status_label)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Excel 行", "SKU", "商品 ID", "结果", "日志"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        root.addWidget(self.table, 1)
        self.setCentralWidget(central)

        menu = self.menuBar().addMenu("设置")
        clear_action = QAction("清空 Cookie 输入", self)
        clear_action.triggered.connect(self.cookie_edit.clear)
        menu.addAction(clear_action)

    def _load_saved_settings(self) -> None:
        self.sku_seed_edit.setText(self.db.get_setting("last_sku_seed", "AA1001"))
        self.seller_edit.setText(self.db.get_setting("seller_id", ""))
        # Older versions saved CSRF tokens.  They are session credentials and
        # should remain only in memory from now on.
        self.db.delete_setting("csrf_token")

    def choose_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择目录 Excel", str(Path.home()), "Excel 文件 (*.xlsx)")
        if not path:
            return
        try:
            reader = XlsxReader(path)
            sheets = reader.sheet_names()
            self.reader = reader
            self.file_edit.setText(path)
            self.sheet_combo.blockSignals(True)
            self.sheet_combo.clear(); self.sheet_combo.addItems(sheets)
            self.sheet_combo.blockSignals(False)
            self.refresh_estimate()
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", str(exc))

    def refresh_estimate(self) -> None:
        path = self.file_edit.text().strip()
        sheet = self.sheet_combo.currentText()
        if not path or not sheet:
            self.estimate_label.setText("请选择 Excel")
            return
        try:
            self.reader = self.reader or XlsxReader(path)
            start, end = self.reader.dimension(sheet)
            first_row = next(self.reader.rows(sheet, start_row=1, max_rows=1), (1, []))[1]
            rows = max(0, end - max(start, 1) + (0 if has_header_row(first_row) else 1))
            if rows:
                if self.count_spin.value() <= 1:
                    self.count_spin.setValue(min(rows, 100))
                count = min(self.count_spin.value(), rows)
                next_sku = self.preview_next_sku(self.sku_seed_edit.text())
                last_sku = self.preview_sku_after(self.sku_seed_edit.text(), count)
                self.estimate_label.setText(f"预计可上传 {rows:,} 条；本轮 {count:,} 条；SKU 范围约 {next_sku} 至 {last_sku}")
            else:
                self.estimate_label.setText("该工作表没有可上传数据")
        except Exception as exc:
            self.estimate_label.setText(f"无法估算：{exc}")

    def preview_next_sku(self, seed_text: str) -> str:
        prefix, seed = split_sku_seed(seed_text)
        active = self.db.get_setting("active_prefix")
        last_text = self.db.get_setting("last_number")
        try: last = int(last_text)
        except ValueError: last = seed - 1
        if not active: number = seed
        elif active != prefix: number = seed + 1
        else: number = max(last + 1, seed)
        return f"{prefix}{number}"

    def preview_sku_after(self, seed_text: str, count: int) -> str:
        prefix, seed = split_sku_seed(seed_text)
        active = self.db.get_setting("active_prefix")
        last_text = self.db.get_setting("last_number")
        try: last = int(last_text)
        except ValueError: last = seed - 1
        if not active: first = seed
        elif active != prefix: first = seed + 1
        else: first = max(last + 1, seed)
        return f"{prefix}{first + max(0, count - 1)}"

    def start_upload(self) -> None:
        if self.worker is not None:
            return
        path = self.file_edit.text().strip()
        sheet = self.sheet_combo.currentText().strip()
        cookie = self.cookie_edit.toPlainText().strip()
        seller_id = self.seller_edit.text().strip()
        seed = self.sku_seed_edit.text().strip()
        if not path or not os.path.exists(path): return self.warn("请选择存在的 Excel 文件")
        if not cookie: return self.warn("请粘贴 Cookie")
        if not self.csrf_edit.text().strip(): return self.warn("请填写 CSRF Token")
        if not seller_id: return self.warn("请填写 Seller ID")
        try: split_sku_seed(seed)
        except ValueError as exc: return self.warn(str(exc))
        if not sheet: return self.warn("请选择工作表")
        try:
            _, last_row = (self.reader or XlsxReader(path)).dimension(sheet)
            available = max(1, last_row - 1)
            count = min(self.count_spin.value(), available)
        except Exception:
            count = self.count_spin.value()
        self.db.set_setting("last_sku_seed", seed); self.db.set_setting("seller_id", seller_id)
        shipping_override = None if self.shipping_mode.currentData() == "excel" else str(self.shipping_days_spin.value())
        config = {"file_path": path, "sheet_name": sheet, "cookie": cookie, "csrf_token": self.csrf_edit.text().strip(), "seller_id": seller_id, "sku_seed": seed, "count": count, "delay_min": 5, "delay_max": 8, "shipping_days_override": shipping_override}
        self.table.setRowCount(0); self.progress_bar.setValue(0); self.start_button.setEnabled(False); self.stop_button.setEnabled(True); self.export_button.setEnabled(False)
        self.thread = QThread(self); self.worker = UploadWorker(config); self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run); self.worker.row_result.connect(self.add_row); self.worker.progress.connect(lambda current, total: self.progress_bar.setValue(int(current * 100 / max(total, 1))))
        self.worker.status.connect(self.status_label.setText); self.worker.finished.connect(self.upload_finished); self.worker.failed.connect(self.upload_failed)
        self.thread.finished.connect(self.worker.deleteLater); self.thread.finished.connect(self.thread.deleteLater); self.thread.start()

    def stop_upload(self) -> None:
        if self.worker: self.worker.stop(); self.status_label.setText("正在停止，当前 HTTP 请求结束后停止")

    @Slot(dict)
    def add_row(self, data: dict) -> None:
        row = self.table.rowCount(); self.table.insertRow(row)
        values = [str(data.get(key, "")) for key in ("row", "sku", "product_id", "status", "message")]
        for col, value in enumerate(values): self.table.setItem(row, col, QTableWidgetItem(value))
        self.table.scrollToBottom()

    @Slot(str, int, int)
    def upload_finished(self, run_id: str, success: int, failure: int) -> None:
        self.current_run_id = run_id; self.worker = None; self.start_button.setEnabled(True); self.stop_button.setEnabled(False); self.export_button.setEnabled(failure > 0); self.status_label.setText(f"完成：成功 {success} 条，失败 {failure} 条。")
        if self.thread: self.thread.quit(); self.thread = None
        self.refresh_estimate()

    @Slot(str)
    def upload_failed(self, message: str) -> None:
        self.worker = None; self.start_button.setEnabled(True); self.stop_button.setEnabled(False); self.status_label.setText("上传批次启动失败")
        if self.thread: self.thread.quit(); self.thread = None
        QMessageBox.critical(self, "上传失败", message)

    def export_failures(self) -> None:
        if not self.current_run_id: return self.warn("当前没有可导出的批次")
        path, _ = QFileDialog.getSaveFileName(self, "保存失败 Excel", str(Path.home() / "makro_失败重试.xlsx"), "Excel 文件 (*.xlsx)")
        if not path: return
        try:
            count = export_failed_xlsx(self.db, self.current_run_id, path)
            QMessageBox.information(self, "导出完成", f"已导出 {count} 条失败记录，可直接选择这个 Excel 重试。\n{path}")
        except Exception as exc: QMessageBox.critical(self, "导出失败", str(exc))

    def warn(self, message: str): QMessageBox.warning(self, "请检查", message)

    def closeEvent(self, event: QCloseEvent) -> None:
        if self.worker is not None:
            QMessageBox.warning(self, "正在上传", "请先点击停止，等待当前请求结束后再退出。")
            event.ignore(); return
        self.db.close(); event.accept()


def main() -> int:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow(); window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
