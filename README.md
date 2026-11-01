# 文件转换助手

本地下载即可使用的**向导式**转换软件。双击启动后按屏幕上的步骤点即可，全程在本机完成，文件不会上传。

---

## 怎么用（五步）

解压后：

- **Windows**：双击 `start.bat`
- **Mac / Linux**：在终端执行 `./start.sh`

启动后窗口顶部会一直显示步骤条：

```
① 选功能  →  ② 加文件  →  ③ 排序  →  ④ 输出设置  →  ⑤ 开始
```

### 第 1 步 · 选功能

点选一张功能卡片（点亮后再点右下角「下一步」）：

| 卡片 | 做什么 |
|------|--------|
| **图片合成 PDF** | 多张图片按顺序合成一份 PDF |
| **PDF 拆成图片** | 把 PDF 每一页导出成 PNG / JPG / WEBP |
| **Word/PPT 转 PDF** | 文档转成 PDF（需安装 LibreOffice） |
| **PDF 转 Word** | PDF → 可编辑的 .docx |
| **PDF 转 PPT** | PDF → PPT（每页一张图，适合放映） |
| **合并多个 PDF** | 多个 PDF 按顺序合成一份 |

选好功能后，软件会自动限制可添加的文件类型，并给出合适的默认「是否合并」。

### 第 2 步 · 加文件

- 点 **「＋ 添加文件…」** 可多选
- 或 **「＋ 添加文件夹…」** 自动收集匹配格式
- 可用「删除选中 / 清空列表」整理

### 第 3 步 · 排序

合并或合成时，**列表从上到下就是最终顺序**。

- **按文件名排序（推荐）**：自然排序（`img2` 会排在 `img10` 前面）
- 或选中一行后 **上移 / 下移 / 移到最上 / 移到最下**

### 第 4 步 · 输出设置

- 是否合并（部分功能固定合并）
- 图片格式 / DPI（仅相关功能显示）
- **保存到哪个文件夹**
- **结果文件名**（扩展名自动补上）

页面下方有「即将执行」摘要，方便核对。

### 第 5 步 · 开始

核对摘要后点蓝色大按钮 **「▶ 开始转换」**。

- 下方有进度条和当前处理提示
- 成功后弹出对话框，可点 **「打开输出文件夹」**

---

## 支持格式

| 方向 | 格式 | 说明 |
|------|------|------|
| 图片 → PDF | png / jpg / jpeg / webp / gif / bmp / tiff / ico / heic / heif / jfif 等 | 多图按列表顺序合并 |
| PDF → 图片 | pdf → png / jpg / webp | 每页一张；可先合并多 PDF |
| Office → PDF | doc / docx / ppt / pptx（及 odt / odp） | 需 **LibreOffice** |
| PDF → Word | pdf → docx | 尽量保留文字表格 |
| PDF → PPT | pdf → pptx | 每页一张图，文字不可直接编辑 |
| 多 PDF 合并 | pdf → pdf | 按列表顺序 |
| SVG → PDF | .svg（可选） | 需额外 `pip install cairosvg` |

每条路径都支持：**批量导入 → 按名/手动排序 → 合并 → 转换**。

---

## 系统要求

1. **Python 3.10+**（首次运行 `start.bat` / `start.sh` 会自动装依赖）
2. **LibreOffice**（仅「Word/PPT 转 PDF」需要）
   - Windows：https://www.libreoffice.org/download/
   - macOS：`brew install --cask libreoffice`
   - Linux：`sudo apt install libreoffice`
3. Linux 若提示缺 tkinter：`sudo apt install python3-tk`

---

## 命令行（进阶，可选）

```bash
cd file-converter
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# 多图合并为 PDF
python convert.py a.png b.jpg c.webp --to pdf --merge -o ./out --name album.pdf

# 多 PDF 合并
python convert.py a.pdf b.pdf --to pdf --merge -o ./out --name all.pdf

# PDF 每页导出为 PNG
python convert.py scan.pdf --to png -o ./pages/

# 文件夹内图片按文件名排序后合并
python convert.py ./photos/ --to pdf --sort --merge -o ./out --name photos.pdf

# Office → PDF / PDF → Word / PPT
python convert.py report.docx --to pdf
python convert.py file.pdf --to docx
python convert.py file.pdf --to pptx --dpi 200
```

常用参数：`--to`、`--merge` / `--no-merge`、`--sort`、`-o`、`--name`、`--dpi`。

---

## 目录结构

```
file-converter/
├── gui.py              # 向导式图形界面（推荐）
├── convert.py          # 转换核心 + 命令行
├── requirements.txt
├── start.bat / start.sh
└── README.md
```

---

## 常见问题

**Q：窗口打不开 / No module named tkinter？**  
A：Windows / macOS 官方 Python 一般自带。Linux：`sudo apt install python3-tk`。

**Q：提示未找到 LibreOffice？**  
A：安装后重启软件；Windows 可确认  
`C:\Program Files\LibreOffice\program\soffice.exe`。

**Q：HEIC 打不开？**  
A：确保依赖已安装（`start.bat` 会装 `pillow-heif`）。

**Q：PDF 转 Word 版式乱了 / PPT 不能改字？**  
A：属能力限制。Word 为尽力还原；PPT 为整页截图，适合放映存档。

---

## 许可与隐私

转换全程本地完成。请遵守所转换文档本身的版权与保密要求。

## 在 Cursor 里继续开发

1. 解压本文件夹到任意短路径目录（例如 `~/projects/file-converter`）
2. Cursor：**文件 → 打开文件夹**，选该目录
3. 用 Agent（Ctrl+I）继续改功能

## 打包成 Windows EXE

1. 在本机打开项目文件夹
2. 双击 `打包成EXE.bat`（需已装 Python）
3. 完成后打开 `dist\FileConverter\FileConverter.exe`
4. 可将整个 `dist\FileConverter` 文件夹拷给别人用（无需再装 Python）
5. Word/PPT→PDF 仍需本机安装 LibreOffice

卸载：删除项目文件夹和 `dist` 即可。
