# 文档转换工具（Word→PDF、Markdown→Word、PDF→图片）

一个基于 Python 的文档转换工具，支持 Word 文档转 PDF（GUI/命令行），并新增 Markdown 转 Word（.docx）脚本。

## 功能特点

- ✅ **单文件转换（Word→PDF）**: 支持将单个 Word 文档转换为 PDF
- ✅ **批量转换（Word→PDF）**: 支持同时转换多个 Word 文档
- ✅ **图形界面（Word→PDF）**: 友好的 GUI 界面，操作简单直观
- ✅ **命令行支持（Word→PDF）**: 支持命令行模式，便于脚本调用
- ✅ **文件格式支持（Word→PDF）**: 支持 .doc、.docx、.rtf 格式
- ✅ **Markdown→Word（docx）**: 使用 Pandoc 高质量转换，自动生成目录
- ✅ **自动目录创建**: 自动创建输出目录
- ✅ **详细日志**: 完整的转换日志记录
- ✅ **错误处理**: 完善的错误处理和异常捕获
- ✅ **进度显示**: 实时显示转换进度

## 安装依赖

```bash
pip install -r requirements.txt
```

> 首次使用 Markdown→Word 功能时，工具会尝试自动下载 Pandoc；网络或安装环境限制可能导致该步骤失败。

## 使用方法

### 图形界面模式（Word→PDF，推荐）

直接运行脚本，会打开图形界面：

```bash
python word_pdf.py
```

#### 单文件转换
1. 点击"浏览"按钮选择要转换的Word文件
2. 选择输出PDF文件的保存位置（可选，默认与输入文件同目录）
3. 点击"开始转换"按钮

#### 批量转换
1. 点击"选择文件"按钮选择多个Word文件
2. 选择输出目录
3. 点击"开始批量转换"按钮

### 命令行模式（Word→PDF）

```bash
# 转换单个文件（自动生成输出文件名）
python word_pdf.py "input.docx"

# 转换单个文件（指定输出文件名）
python word_pdf.py "input.docx" "output.pdf"
```

### 命令行模式（Markdown→Word）

```bash
# 将 README.md 转为 README.docx（同目录输出）
python md_to_word.py README.md

# 指定输出文件路径
python md_to_word.py docs/guide.md out/guide.docx
```

#### Markdown→Word 转换说明
- 支持 .md / .markdown 扩展名
- 自动生成目录（最多 3 级标题）
- 支持常见 Markdown 扩展：emoji、裸链接、紧凑列表等

## 系统要求

- Windows 操作系统
- Microsoft Word 已安装（仅 Word→PDF 功能需要）
- Python 3.8+
- 依赖：pywin32（Word→PDF），pypandoc（Markdown→Word）

## 注意事项

1. Word→PDF 功能需要已安装 Microsoft Word
2. 转换过程中请勿关闭 Word 应用程序
3. 大文件转换可能需要较长时间，请耐心等待
4. Word→PDF 日志保存在 `word_to_pdf.log`
5. Markdown→Word 日志保存在 `md_to_word.log`

## 错误排查

如果遇到转换失败，请检查：

1.（Word→PDF）Word 应用程序是否正常运行
2. 输入文件是否存在且格式正确
3. 输出目录是否有写入权限
4. 查看日志文件获取详细错误信息
5.（Markdown→Word）首次运行如提示缺少 Pandoc，请联网后重试

## 更新日志

### v2.1
- 新增 `md_to_word.py`，支持 Markdown→Word（docx）
- 自动下载 Pandoc 以简化安装

### v2.0
- 添加图形界面支持
- 支持批量转换
- 改进错误处理和日志记录
- 添加进度显示
- 支持命令行模式
- 自动目录创建
- 文件格式验证

### v1.0
- 基础单文件转换功能

## PDF 转图片

- [`pdf_to_images.py`](./pdf_to_images.py)：PDF 转图片脚本。
- [`pdf_to_images_gui.py`](./pdf_to_images_gui.py)：图形界面入口。

具体选项与依赖请参照脚本。

## 脚本、依赖与适用平台

| 脚本 | 功能 | 关键依赖 |
| --- | --- | --- |
| [word_pdf.py](./word_pdf.py) | Word→PDF，GUI/命令行 | Windows、Microsoft Word、pywin32 |
| [md_to_word.py](./md_to_word.py) | Markdown→DOCX | pypandoc 与 Pandoc |
| [pdf_to_images.py](./pdf_to_images.py) | PDF 每页导出图片 | PyMuPDF |
| [pdf_to_images_gui.py](./pdf_to_images_gui.py) | PDF 导图界面 | PyMuPDF 与 Tk 图形环境 |

统一 [requirements.txt](./requirements.txt) 包含 pywin32，因此整份依赖列表主要适用于 Windows。若仅在其他平台使用 Markdown 或 PDF 导图，可按所用脚本配置依赖；仓库没有给出完整跨平台测试矩阵。

## 推荐安装流程

以 Windows PowerShell 为例：

```powershell
git clone --branch main https://github.com/heisd/word_to_pdf.git
cd word_to_pdf
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Word→PDF 通过 `Dispatch("Word.Application")` 调用本机 Word，并不是纯 Python 解析转换。安装 pywin32 不能替代安装 Microsoft Word。

## PDF 导图：完整命令参数

```text
python pdf_to_images.py INPUT_PDF [OUTPUT_DIR] [png|jpg] [ZOOM] [PAGE_RANGE]
```

| 参数 | 默认行为 |
| --- | --- |
| INPUT_PDF | 必填，输入 PDF 文件 |
| OUTPUT_DIR | 输入文件旁的“文件名_images”目录 |
| 图片格式 | png |
| ZOOM | 2.0，数值更大时像素尺寸与资源开销也更大 |
| PAGE_RANGE | 全部页；也支持 1-5 或单页 3 |

示例：

```bash
python pdf_to_images.py "input.pdf"
python pdf_to_images.py "input.pdf" "out" png 2.0 1-5
python pdf_to_images.py "input.pdf" "out" jpg 2.0 3
```

输出按原始页码命名，例如 `input_p1.png`；位数会随页范围调整。导图函数还定义了 JPEG 质量和透明背景相关参数，但当前命令行入口没有将它们全部暴露为选项。

图形界面入口：

```bash
python pdf_to_images_gui.py
```

界面可选择输入、输出目录、PNG/JPG、缩放和页范围。PDF 导图是把页面渲染为图片，不是 OCR 或文字提取。

## Markdown 输出说明

`md_to_word.py` 默认将输入同名文件扩展名替换为 `.docx`，也可以指定输出位置。脚本会创建缺失的输出目录，并传入三级目录等 Pandoc 参数。

```bash
python md_to_word.py "notes.md" "out/notes.docx"
```

复杂表格、特殊字体与自定义样式仍需在 Word 中检查。Pandoc 可用性、Markdown 转换与 Word→PDF 是三套不同前提，应分别排查。

## 日志与故障定位

| 日志 | 对应功能 |
| --- | --- |
| word_to_pdf.log | Word→PDF |
| md_to_word.log | Markdown→DOCX |
| pdf_to_images.log | PDF→图片 |

日志位于进程当前工作目录。建议从仓库根目录运行，路径含空格时在命令行使用引号。

- **Word COM 无法初始化**：先确认本机 Word 可正常打开、所用环境为 Windows，再检查 pywin32。
- **Pandoc 下载失败**：检查网络和本机 Pandoc 可用性；脚本会尝试下载，不保证所有环境自动安装成功。
- **找不到 fitz**：安装的是 PyMuPDF，导入名才是 fitz。
- **导图占用过高**：先减小缩放倍数或缩小页范围。
- **输出不符合预期**：先用单页/单文件确认格式、目录与页面，再进行大批量转换。

重复转换时应检查目标位置，避免把同名输出误当成另一次结果。
