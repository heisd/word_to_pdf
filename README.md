# 文档转换工具（Word⇄PDF、Markdown→Word）

一个功能强大的文档转换工具，支持 Word 文档转 PDF（GUI/命令行），并新增 Markdown 转 Word（.docx）脚本。

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

> 首次使用 Markdown→Word 功能时，工具会自动下载并安装 Pandoc（无需管理员权限）。

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
