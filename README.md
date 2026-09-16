# Makro 商品发布工具

这是一个本地 PySide6 GUI。它按 `目录.xlsx` 的商品详情页地址提取 `pid`，然后按 HAR 中的 `create-update-listings` 请求格式逐条发布商品。

## 使用

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

1. 在浏览器开发者工具中复制 Makro 站点请求使用的 Cookie，粘贴到 Cookie 框。
2. 填写当前登录账号的 Seller ID。
3. 把浏览器请求头 `fk-csrf-token` 的值填入 CSRF Token。
4. 选择 Excel 和工作表。默认按表头匹配，找不到表头时回退到 D 列原价、E 列优惠价、G 列预计天数、H 列详情页地址。
5. SKU 起始值首次使用时，例如 `AA1001` 会从 `AA1001` 开始。数据库已有同一前缀记录时会继续递增。更换前缀时，例如输入 `AACC1021`，下一条从 `AACC1022` 开始。
6. 点击开始上传。程序会在每条 HTTP 请求前随机等待 5–8 秒，停止按钮会在当前请求结束后停止。
7. 失败记录会写入 `makro_uploader.sqlite3`。上传结束后点击“导出本批次失败 Excel”，导出的文件只保留失败行，并在最后追加“错误”列，可直接再次选择该文件重试。

## 说明

- 数据库、Cookie 和上传日志都保存在本机，不会上传到第三方服务。
- Excel 解析采用流式读取，适合当前超过百万行的目录文件，不会一次性把整本表载入内存。
- SKU 会在发送前去除首尾空白，并在本地数据库中记录已分配值。失败重试会按商品行指纹复用原 SKU，避免重复生成 SKU。
- 本工具只实现 HAR 中已确认的商品发布请求，不模拟浏览器页面点击。
