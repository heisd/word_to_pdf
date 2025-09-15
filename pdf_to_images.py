import os
import sys
import logging
from pathlib import Path
from typing import Optional, Tuple

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None  # type: ignore


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pdf_to_images.log', encoding='utf-8'),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)


def parse_page_range(page_range: Optional[str], page_count: int) -> Tuple[int, int]:
    """Parse page range like '1-5' (1-based inclusive). Returns 0-based [start, end).
    Empty or None means all pages.
    """
    if not page_range:
        return 0, page_count
    try:
        if '-' in page_range:
            start_s, end_s = page_range.split('-', 1)
            start = int(start_s) if start_s.strip() else 1
            end = int(end_s) if end_s.strip() else page_count
        else:
            start = int(page_range)
            end = start
        start = max(1, start)
        end = min(page_count, end)
        if end < start:
            raise ValueError("页码范围无效：结束页小于起始页")
        return start - 1, end  # convert to 0-based [start, end)
    except Exception as exc:
        raise ValueError(f"页码范围格式错误: '{page_range}'，示例：'1-5' 或 '3'") from exc


def convert_pdf_to_images(
    input_pdf: str,
    output_dir: Optional[str] = None,
    image_format: str = 'png',
    zoom: float = 2.0,
    page_range: Optional[str] = None,
    jpg_quality: int = 92,
    no_alpha: bool = True,
) -> Path:
    """Convert PDF pages to images.

    - image_format: 'png' or 'jpg'
    - zoom: 1.0=72dpi 基础缩放，2.0≈144dpi，3.0≈216dpi
    - page_range: 'start-end' or 'n' (1-based). None for all
    - no_alpha: True to remove alpha channel (recommended for PNG)

    Returns the output directory as Path.
    """
    if fitz is None:
        raise RuntimeError("缺少依赖：未安装 PyMuPDF。请先运行 `pip install -r requirements.txt`。")

    input_path = Path(input_pdf)
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_pdf}")
    if input_path.suffix.lower() != '.pdf':
        raise ValueError("仅支持 .pdf 文件")

    out_dir = Path(output_dir) if output_dir else input_path.parent / f"{input_path.stem}_images"
    out_dir.mkdir(parents=True, exist_ok=True)

    image_format = image_format.lower()
    if image_format not in {"png", "jpg", "jpeg"}:
        raise ValueError("image_format 仅支持 'png' 或 'jpg'")
    if image_format == 'jpeg':
        image_format = 'jpg'

    with fitz.open(str(input_path)) as doc:
        start, end = parse_page_range(page_range, doc.page_count)
        matrix = fitz.Matrix(zoom, zoom)

        digits = len(str(end))
        for i in range(start, end):
            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=matrix, alpha=not no_alpha)

            page_index = i + 1
            out_name = f"{input_path.stem}_p{page_index:0{digits}d}.{image_format}"
            out_file = out_dir / out_name

            if image_format == 'png':
                pix.save(str(out_file))
            else:
                # Save as JPEG with quality
                pix.save(str(out_file), output='jpg', jpg_quality=jpg_quality)

            logger.info("导出: %s", out_file)

    logger.info("转换完成，输出目录: %s", out_dir)
    return out_dir


def main() -> int:
    if len(sys.argv) < 2:
        print(
            "用法: python pdf_to_images.py <input.pdf> [output_dir] [png|jpg] [zoom] [page_range]\n"
            "示例: python pdf_to_images.py file.pdf out png 2.0 1-5"
        )
        return 2

    input_pdf = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) >= 3 else None
    image_format = sys.argv[3] if len(sys.argv) >= 4 else 'png'
    zoom = float(sys.argv[4]) if len(sys.argv) >= 5 else 2.0
    page_range = sys.argv[5] if len(sys.argv) >= 6 else None

    try:
        convert_pdf_to_images(
            input_pdf=input_pdf,
            output_dir=output_dir,
            image_format=image_format,
            zoom=zoom,
            page_range=page_range,
        )
        print("转换成功！")
        return 0
    except Exception as e:
        logger.error("转换失败: %s", e)
        print(f"转换失败: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())


