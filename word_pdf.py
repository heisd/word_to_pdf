import os
import sys
import logging
from pathlib import Path
from typing import List, Optional
from win32com.client import Dispatch, constants
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('word_to_pdf.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class WordToPDFConverter:
    """Word转PDF转换器类"""
    
    def __init__(self):
        self.word_app = None
        self.supported_formats = ['.doc', '.docx', '.rtf']
    
    def _init_word_app(self):
        """初始化Word应用程序"""
        try:
            self.word_app = Dispatch("Word.Application")
            self.word_app.Visible = False
            logger.info("Word应用程序启动成功")
            return True
        except Exception as e:
            logger.error(f"启动Word应用程序失败: {str(e)}")
            return False
    
    def _cleanup_word_app(self):
        """清理Word应用程序"""
        if self.word_app:
            try:
                self.word_app.Quit()
                self.word_app = None
                logger.info("Word应用程序已关闭")
            except Exception as e:
                logger.error(f"关闭Word应用程序时出错: {str(e)}")
    
    def validate_file(self, file_path: str) -> bool:
        """验证文件是否存在且格式支持"""
        if not os.path.exists(file_path):
            logger.error(f"文件不存在: {file_path}")
            return False
        
        file_ext = Path(file_path).suffix.lower()
        if file_ext not in self.supported_formats:
            logger.error(f"不支持的文件格式: {file_ext}")
            return False
        
        return True
    
    def convert_single_file(self, input_file: str, output_file: str) -> bool:
        """转换单个文件"""
        if not self.validate_file(input_file):
            return False
        
        # 确保输出目录存在
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            logger.info(f"创建输出目录: {output_dir}")
        
        try:
            if not self.word_app:
                if not self._init_word_app():
                    return False
            
            # 打开文档
            doc = self.word_app.Documents.Open(input_file)
            logger.info(f"正在转换: {os.path.basename(input_file)}")
            
            # 转换为PDF
            doc.SaveAs(output_file, FileFormat=17)  # 17代表PDF格式
            doc.Close()
            
            logger.info(f"转换成功: {output_file}")
            return True
            
        except Exception as e:
            logger.error(f"转换文件失败 {input_file}: {str(e)}")
            return False
    
    def convert_batch(self, input_files: List[str], output_dir: str) -> dict:
        """批量转换文件"""
        results = {
            'success': [],
            'failed': [],
            'total': len(input_files)
        }
        
        if not self._init_word_app():
            logger.error("无法启动Word应用程序")
            return results
        
        try:
            for i, input_file in enumerate(input_files, 1):
                logger.info(f"处理文件 {i}/{len(input_files)}: {os.path.basename(input_file)}")
                
                # 生成输出文件名
                input_name = Path(input_file).stem
                output_file = os.path.join(output_dir, f"{input_name}.pdf")
                
                if self.convert_single_file(input_file, output_file):
                    results['success'].append(output_file)
                else:
                    results['failed'].append(input_file)
        
        finally:
            self._cleanup_word_app()
        
        return results
    
    def convert_single(self, input_file: str, output_file: Optional[str] = None) -> bool:
        """转换单个文件（简化接口）"""
        if not output_file:
            # 自动生成输出文件名
            input_path = Path(input_file)
            output_file = str(input_path.parent / f"{input_path.stem}.pdf")
        
        if not self._init_word_app():
            return False
        
        try:
            return self.convert_single_file(input_file, output_file)
        finally:
            self._cleanup_word_app()

class WordToPDFGUI:
    """图形界面类"""
    
    def __init__(self):
        self.converter = WordToPDFConverter()
        self.root = tk.Tk()
        self.root.title("Word转PDF工具")
        self.root.geometry("600x500")
        
        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        
        self.setup_ui()
    
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 标题
        title_label = ttk.Label(main_frame, text="Word转PDF转换工具", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 单文件转换区域
        single_frame = ttk.LabelFrame(main_frame, text="单文件转换", padding="10")
        single_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 输入文件选择
        ttk.Label(single_frame, text="选择Word文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.input_file_var = tk.StringVar()
        input_entry = ttk.Entry(single_frame, textvariable=self.input_file_var, width=50)
        input_entry.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(single_frame, text="浏览", 
                  command=self.browse_input_file).grid(row=1, column=2, padx=(5, 0))
        
        # 输出文件选择
        ttk.Label(single_frame, text="输出PDF文件:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.output_file_var = tk.StringVar()
        output_entry = ttk.Entry(single_frame, textvariable=self.output_file_var, width=50)
        output_entry.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(single_frame, text="浏览", 
                  command=self.browse_output_file).grid(row=3, column=2, padx=(5, 0))
        
        # 单文件转换按钮
        ttk.Button(single_frame, text="开始转换", 
                  command=self.convert_single).grid(row=4, column=0, pady=10)
        
        # 批量转换区域
        batch_frame = ttk.LabelFrame(main_frame, text="批量转换", padding="10")
        batch_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 批量文件选择
        ttk.Label(batch_frame, text="选择多个Word文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Button(batch_frame, text="选择文件", 
                  command=self.browse_input_files).grid(row=0, column=1, padx=(10, 0))
        
        # 输出目录选择
        ttk.Label(batch_frame, text="输出目录:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_dir_var = tk.StringVar()
        output_dir_entry = ttk.Entry(batch_frame, textvariable=self.output_dir_var, width=50)
        output_dir_entry.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(batch_frame, text="浏览", 
                  command=self.browse_output_dir).grid(row=2, column=2, padx=(5, 0))
        
        # 批量转换按钮
        ttk.Button(batch_frame, text="开始批量转换", 
                  command=self.convert_batch).grid(row=3, column=0, pady=10)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, 
                                          maximum=100, length=400)
        self.progress_bar.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # 状态标签
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.grid(row=4, column=0, columnspan=3, pady=5)
        
        # 日志文本框
        log_frame = ttk.LabelFrame(main_frame, text="转换日志", padding="5")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        self.log_text = tk.Text(log_frame, height=8, width=70)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 配置网格权重
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        single_frame.columnconfigure(0, weight=1)
        batch_frame.columnconfigure(0, weight=1)
    
    def browse_input_file(self):
        """浏览输入文件"""
        file_path = filedialog.askopenfilename(
            title="选择Word文件",
            filetypes=[("Word文档", "*.doc *.docx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.input_file_var.set(file_path)
            # 自动设置输出文件路径
            input_path = Path(file_path)
            output_path = input_path.parent / f"{input_path.stem}.pdf"
            self.output_file_var.set(str(output_path))
    
    def browse_output_file(self):
        """浏览输出文件"""
        file_path = filedialog.asksaveasfilename(
            title="保存PDF文件",
            defaultextension=".pdf",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file_path:
            self.output_file_var.set(file_path)
    
    def browse_input_files(self):
        """浏览多个输入文件"""
        file_paths = filedialog.askopenfilenames(
            title="选择多个Word文件",
            filetypes=[("Word文档", "*.doc *.docx"), ("所有文件", "*.*")]
        )
        if file_paths:
            self.input_files = list(file_paths)
            self.status_var.set(f"已选择 {len(file_paths)} 个文件")
    
    def browse_output_dir(self):
        """浏览输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_dir_var.set(dir_path)
    
    def log_message(self, message):
        """在日志文本框中显示消息"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def convert_single(self):
        """转换单个文件"""
        input_file = self.input_file_var.get()
        output_file = self.output_file_var.get()
        
        if not input_file:
            messagebox.showerror("错误", "请选择输入文件")
            return
        
        if not output_file:
            messagebox.showerror("错误", "请选择输出文件")
            return
        
        def convert_thread():
            self.status_var.set("正在转换...")
            self.progress_var.set(0)
            
            try:
                success = self.converter.convert_single(input_file, output_file)
                if success:
                    self.status_var.set("转换完成")
                    self.progress_var.set(100)
                    self.log_message(f"转换成功: {output_file}")
                    messagebox.showinfo("成功", f"文件已转换完成！\n保存位置: {output_file}")
                else:
                    self.status_var.set("转换失败")
                    self.log_message(f"转换失败: {input_file}")
                    messagebox.showerror("错误", "转换失败，请检查日志")
            except Exception as e:
                self.status_var.set("转换出错")
                self.log_message(f"转换出错: {str(e)}")
                messagebox.showerror("错误", f"转换出错: {str(e)}")
        
        threading.Thread(target=convert_thread, daemon=True).start()
    
    def convert_batch(self):
        """批量转换文件"""
        if not hasattr(self, 'input_files') or not self.input_files:
            messagebox.showerror("错误", "请选择要转换的文件")
            return
        
        output_dir = self.output_dir_var.get()
        if not output_dir:
            messagebox.showerror("错误", "请选择输出目录")
            return
        
        def convert_thread():
            self.status_var.set("正在批量转换...")
            self.progress_var.set(0)
            
            try:
                results = self.converter.convert_batch(self.input_files, output_dir)
                
                self.progress_var.set(100)
                self.status_var.set("批量转换完成")
                
                success_count = len(results['success'])
                failed_count = len(results['failed'])
                
                self.log_message(f"批量转换完成: 成功 {success_count} 个，失败 {failed_count} 个")
                
                if results['failed']:
                    self.log_message("失败的文件:")
                    for failed_file in results['failed']:
                        self.log_message(f"  - {failed_file}")
                
                messagebox.showinfo("完成", 
                    f"批量转换完成！\n成功: {success_count} 个\n失败: {failed_count} 个")
                
            except Exception as e:
                self.status_var.set("批量转换出错")
                self.log_message(f"批量转换出错: {str(e)}")
                messagebox.showerror("错误", f"批量转换出错: {str(e)}")
        
        threading.Thread(target=convert_thread, daemon=True).start()
    
    def run(self):
        """运行GUI"""
        self.root.mainloop()

def main():
    """主函数"""
    if len(sys.argv) > 1:
        # 命令行模式
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        
        converter = WordToPDFConverter()
        success = converter.convert_single(input_file, output_file)
        
        if success:
            print("转换成功！")
            sys.exit(0)
        else:
            print("转换失败！")
            sys.exit(1)
    else:
        # GUI模式
        app = WordToPDFGUI()
        app.run()

if __name__ == "__main__":
    main()