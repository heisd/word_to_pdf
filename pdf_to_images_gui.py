import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import logging

from pdf_to_images import convert_pdf_to_images


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pdf_to_images.log', encoding='utf-8'),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)


class PDFToImagesGUI:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("PDF转图片工具")
        self.root.geometry("600x420")

        style = ttk.Style()
        style.theme_use('clam')

        self.setup_ui()

    def setup_ui(self) -> None:
        main = ttk.Frame(self.root, padding="10")
        main.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))

        # 标题
        title = ttk.Label(main, text="PDF 转 图片 (每页导出)", font=("Arial", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, pady=(0, 16))

        # 输入 PDF
        ttk.Label(main, text="选择PDF文件:").grid(row=1, column=0, sticky=tk.W)
        self.input_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.input_var, width=48).grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=4)
        ttk.Button(main, text="浏览", command=self.pick_pdf).grid(row=2, column=2, padx=(6, 0))

        # 输出目录
        ttk.Label(main, text="输出目录:").grid(row=3, column=0, sticky=tk.W)
        self.output_dir_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.output_dir_var, width=48).grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=4)
        ttk.Button(main, text="浏览", command=self.pick_outdir).grid(row=4, column=2, padx=(6, 0))

        # 参数：格式、缩放、页码
        options = ttk.LabelFrame(main, text="参数", padding="10")
        options.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(8, 8))

        ttk.Label(options, text="图片格式:").grid(row=0, column=0, sticky=tk.W)
        self.format_var = tk.StringVar(value="png")
        ttk.Combobox(options, textvariable=self.format_var, values=["png", "jpg"], width=8, state="readonly").grid(row=0, column=1, padx=(6, 16))

        ttk.Label(options, text="缩放(1.0~4.0):").grid(row=0, column=2, sticky=tk.W)
        self.zoom_var = tk.DoubleVar(value=2.0)
        ttk.Spinbox(options, from_=1.0, to=4.0, increment=0.5, textvariable=self.zoom_var, width=6).grid(row=0, column=3, padx=(6, 16))

        ttk.Label(options, text="页码范围(如1-5):").grid(row=0, column=4, sticky=tk.W)
        self.range_var = tk.StringVar()
        ttk.Entry(options, textvariable=self.range_var, width=10).grid(row=0, column=5, padx=(6, 0))

        # 操作区
        ttk.Button(main, text="开始转换", command=self.convert).grid(row=6, column=0, pady=(8, 8))

        # 进度与状态
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(main, textvariable=self.status_var).grid(row=7, column=0, columnspan=3, sticky=tk.W)

        self.log = tk.Text(main, height=10, width=70)
        scroll = ttk.Scrollbar(main, orient=tk.VERTICAL, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        scroll.grid(row=8, column=2, sticky=(tk.N, tk.S))

        main.columnconfigure(0, weight=1)
        main.rowconfigure(8, weight=1)

    def pick_pdf(self) -> None:
        path = filedialog.askopenfilename(title="选择PDF文件", filetypes=[("PDF", "*.pdf"), ("所有文件", "*.*")])
        if path:
            self.input_var.set(path)
            # 默认输出目录
            p = Path(path)
            default_out = p.parent / f"{p.stem}_images"
            self.output_dir_var.set(str(default_out))

    def pick_outdir(self) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir_var.set(path)

    def append_log(self, text: str) -> None:
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.root.update()

    def convert(self) -> None:
        input_pdf = self.input_var.get().strip()
        out_dir = self.output_dir_var.get().strip()
        img_fmt = self.format_var.get().strip()
        zoom = float(self.zoom_var.get())
        rng = self.range_var.get().strip() or None

        if not input_pdf:
            messagebox.showerror("错误", "请选择PDF文件")
            return

        def task():
            try:
                self.status_var.set("正在转换...")
                out = convert_pdf_to_images(
                    input_pdf=input_pdf,
                    output_dir=out_dir or None,
                    image_format=img_fmt,
                    zoom=zoom,
                    page_range=rng,
                )
                self.append_log(f"转换完成: {out}")
                self.status_var.set("转换完成")
                messagebox.showinfo("完成", f"图片已导出到: {out}")
            except Exception as e:
                self.append_log(f"出错: {e}")
                self.status_var.set("转换失败")
                messagebox.showerror("错误", str(e))

        threading.Thread(target=task, daemon=True).start()

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = PDFToImagesGUI()
    app.run()


if __name__ == "__main__":
    main()


