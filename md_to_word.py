import os
import sys
import logging
from pathlib import Path

try:
    import pypandoc
except ImportError as err:  # Defer hard fail to runtime with clear message
    pypandoc = None  # type: ignore


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('md_to_word.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def ensure_pandoc_available() -> None:
    """Ensure pandoc is available. Try to download a private copy if missing."""
    if pypandoc is None:
        raise RuntimeError(
            "缺少依赖: pypandoc 未安装。请先运行 `pip install -r requirements.txt`。"
        )
    try:
        _ = pypandoc.get_pandoc_version()
        return
    except OSError:
        logger.info("未检测到 Pandoc，正在尝试自动下载...")
        try:
            pypandoc.download_pandoc()
            logger.info("Pandoc 下载完成。")
        except Exception as e:
            raise RuntimeError(
                f"无法自动下载 Pandoc，请手动安装后重试。错误: {e}"
            ) from e


def convert_markdown_to_docx(input_md: str, output_docx: str) -> None:
    """
    Convert a Markdown file to a Word (.docx) file using pandoc.
    """
    input_path = Path(input_md)
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_md}")
    if input_path.suffix.lower() not in {'.md', '.markdown'}:
        raise ValueError("仅支持 .md / .markdown 文件")

    output_path = Path(output_docx)
    if not output_path.suffix:
        output_path = output_path.with_suffix('.docx')
    if not output_path.parent.exists():
        output_path.parent.mkdir(parents=True, exist_ok=True)

    ensure_pandoc_available()

    # Extra args: enable common extensions and smart punctuation
    extra_args = [
        "--from=markdown+emoji+autolink_bare_uris+lists_without_preceding_blankline",
        "--to=docx",
        "--standalone",
        "--toc",  # generate table of contents if headings present
        "--toc-depth=3",
        "--markdown-headings=setext",
        "--wrap=auto",
        "--quiet",
    ]

    # Perform conversion
    pypandoc.convert_file(
        source_file=str(input_path),
        to='docx',
        outputfile=str(output_path),
        extra_args=extra_args,
    )

    logger.info("转换成功: %s -> %s", input_path.name, output_path)


def main() -> int:
    if len(sys.argv) < 2:
        print("用法: python md_to_word.py <input.md> [output.docx]")
        return 2

    input_md = sys.argv[1]
    output_docx = (
        sys.argv[2]
        if len(sys.argv) >= 3
        else str(Path(input_md).with_suffix('.docx'))
    )

    try:
        convert_markdown_to_docx(input_md, output_docx)
        print("转换成功！")
        return 0
    except Exception as e:
        logger.error("转换失败: %s", e)
        print(f"转换失败: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())


