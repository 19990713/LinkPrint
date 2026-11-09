# LinkPrint

热文件夹批量打印：把 PDF / Word / 表格 / PPT / 图片丢进「热文件夹」，网页里逐份确认页数、
尺寸、方向，设好份数/双面/色彩/纸张，一次性发出打印，并按页计费记账。

- 预检（页数、实际尺寸、方向、混合尺寸提示）由 PyMuPDF 读 PDF 得出
- Office 文件靠 `Convert2PDF.ps1`（WPS/Office COM）转 PDF，转换结果按 `名字_大小_时间.pdf` 缓存
- 图片（jpg/png/bmp/tif/webp/gif）由 PyMuPDF 直接转 PDF，多页 TIFF 保留多页
- 送印走 SumatraPDF 命令行；纸张用 `paperkind=`（DMPAPER 编号）避开驱动纸名匹配问题
- 打印记录写入 `打印记录.csv`

## 运行

需要 Windows、Python 3.12+、`pip install pymupdf pywin32`。

**SumatraPDF.exe 不在本仓库**（16MB 第三方程序）：从 https://www.sumatrapdfreader.org/download-free-pdf-viewer
下载 portable 版，把 `SumatraPDF.exe` 放到本目录，与 `print_system.py` 同级。

```
python print_system.py        # 或双击 启动打印系统.bat
python 自检.py                 # 纸型识别 / 参数串 / 计费 / 真实 PDF 预检 自检
```

启动后自动打开 `http://127.0.0.1:8777`（端口被占会自动往后找）。

## config.json

首次启动自动生成；改完重启生效。

| 字段 | 说明 |
| --- | --- |
| `port` | 起始端口，默认 8777 |
| `hotfolder` | 热文件夹，相对路径按本目录算 |
| `printers` | 按**关键字**匹配本机打印机名（如 `D95`、`C7580`），改显卡/改机器只需改这里 |
| `papers` | 界面可选的纸张，默认 `A4`/`A3` |
| `sumatra` | SumatraPDF 路径 |
| `convert2pdf` | `Convert2PDF.ps1` 的路径，找不到它就只支持 PDF 和图片 |

单价在 `print_system.py` 的 `A4_PRICE`（A4 每页：黑白 / 彩色，A3 翻倍）。
