import pymupdf as fitz
import numpy as np
from tqdm import tqdm
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
import os
import threading
import queue

class RedirectText:
    """用于重定向控制台输出到GUI文本框"""
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.queue = queue.Queue()
        self.update_timer = None
        self.start_timer()

    def write(self, string):
        self.queue.put(string)

    def flush(self):
        pass
    
    def update_text_widget(self):
        while not self.queue.empty():
            text = self.queue.get_nowait()
            self.text_widget.configure(state="normal")
            self.text_widget.insert(tk.END, text)
            self.text_widget.see(tk.END)  # 自动滚动到底部
            self.text_widget.configure(state="disabled")
        self.start_timer()
    
    def start_timer(self):
        self.update_timer = self.text_widget.after(100, self.update_text_widget)
    
    def stop_timer(self):
        if self.update_timer:
            self.text_widget.after_cancel(self.update_timer)

class PDFColorSplitter(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("彩色打印分离工具")
        self.geometry("600x500")
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.create_widgets()
        self.center_window()
        
        # 状态变量
        self.input_pdf_path = tk.StringVar()
        self.is_double_sided = tk.BooleanVar(value=True)
        self.is_processing = False
        
    def center_window(self):
        """将窗口居中显示"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"+{x}+{y}")
        
    def create_widgets(self):
        """创建所有界面组件"""
        # 标题标签
        title_label = ttk.Label(self, text="彩色打印分离工具", font=("SimHei", 16, "bold"))
        title_label.pack(pady=15)
        
        # 文件选择框架
        file_frame = ttk.Frame(self)
        file_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Label(file_frame, text="PDF文件:").pack(side=tk.LEFT, padx=5)
        self.file_entry = ttk.Entry(file_frame)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        ttk.Button(file_frame, text="选择文件", command=self.select_file).pack(side=tk.LEFT, padx=5)
        
        # 选项框架
        options_frame = ttk.Frame(self)
        options_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.double_sided_check = ttk.Checkbutton(options_frame, text="双面打印模式")
        self.double_sided_check.pack(side=tk.LEFT, padx=5)
        
        # 操作按钮
        actions_frame = ttk.Frame(self)
        actions_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.convert_button = ttk.Button(actions_frame, text="开始转换", command=self.start_conversion)
        self.convert_button.pack(side=tk.LEFT, padx=5)
        
        # 进度框架
        progress_frame = ttk.Frame(self)
        progress_frame.pack(fill=tk.X, padx=20, pady=5)
        
        ttk.Label(progress_frame, text="处理进度:").pack(side=tk.LEFT, padx=5)
        self.progress_bar = ttk.Progressbar(progress_frame, length=200, mode="indeterminate")
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 输出文本框
        output_frame = ttk.LabelFrame(self, text="处理日志")
        output_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        self.output_text = scrolledtext.ScrolledText(output_frame, width=70, height=12)
        self.output_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.output_text.configure(state="disabled")
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 设置重定向
        self.redirect = RedirectText(self.output_text)
        
    def select_file(self):
        """选择PDF文件"""
        file_path = filedialog.askopenfilename(
            title="选择要分离的PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.input_pdf_path.set(file_path)
            self.status_var.set(f"已选择文件: {os.path.basename(file_path)}")
            
    def start_conversion(self):
        """开始转换过程"""
        if self.is_processing:
            messagebox.showwarning("处理中", "正在处理文件，请稍候...")
            return
            
        input_path = self.file_entry.get()
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("错误", "请先选择有效的PDF文件")
            return
            
        # 准备开始处理
        self.is_processing = True
        self.convert_button.configure(state="disabled")
        self.progress_bar.start()
        self.status_var.set("正在处理PDF文件...")
        
        # 清空输出区域
        self.output_text.configure(state="normal")
        self.output_text.delete(1.0, tk.END)
        self.output_text.configure(state="disabled")
        
        # 在后台线程中处理以避免UI冻结
        threading.Thread(target=self.process_pdf, daemon=True).start()
        
    def process_pdf(self):
        """在后台处理PDF文件"""
        import sys
        # 保存原始的stdout
        original_stdout = sys.stdout
        sys.stdout = self.redirect
        
        try:
            # 提取文件信息
            input_dir = os.path.dirname(self.input_pdf_path.get())
            input_filename = os.path.splitext(os.path.basename(self.input_pdf_path.get()))[0]
            
            # 构建输出路径
            output_color_pdf_path = os.path.join(input_dir, f"{input_filename}_color.pdf")
            output_bw_pdf_path = os.path.join(input_dir, f"{input_filename}_bw.pdf")
            
            # 调用分离函数
            has_color, color_path, bw_path, color_pages, bw_pages = self.split_pdf(
                self.input_pdf_path.get(), 
                output_color_pdf_path, 
                output_bw_pdf_path, 
                self.is_double_sided.get()
            )
            
            if has_color:
                print(f"\n彩色文件已分离: {color_path}")
                print(f"黑白文件已分离: {bw_path}")
                
                # 显示页面分布表格
                total_pages = len(color_pages) + len(bw_pages)
                self.display_page_distribution(color_pages, bw_pages, total_pages, self.is_double_sided.get())
                
                self.after(0, lambda: self.status_var.set("处理完成"))
                self.after(0, lambda: messagebox.showinfo("处理完成", 
                                                        f"PDF文件已成功分离为彩色和黑白两个文件\n"
                                                        f"彩色文件: {os.path.basename(color_path)}\n"
                                                        f"黑白文件: {os.path.basename(bw_path)}"))
            else:
                print(f"\n文件 {self.input_pdf_path.get()} 是纯黑白文档，无需分离。")
                self.after(0, lambda: self.status_var.set("文件是纯黑白文档"))
                self.after(0, lambda: messagebox.showinfo("处理结果", "该文件是纯黑白文档，无需分离"))
        
        except Exception as e:
            print(f"处理过程中出现错误: {str(e)}")
            self.after(0, lambda: self.status_var.set("处理失败"))
            self.after(0, lambda: messagebox.showerror("错误", f"处理PDF时发生错误:\n{str(e)}"))
        
        finally:
            # 恢复标准输出
            sys.stdout = original_stdout
            
            # 更新UI状态
            self.after(0, lambda: self.progress_bar.stop())
            self.after(0, lambda: self.convert_button.configure(state="normal"))
            self.is_processing = False
    
    def is_color_image(self, image, saturation_threshold=0.35, color_fraction_threshold=0.001):
        image = image.convert('RGB')
        pixels = np.array(image) / 255.0  # 归一化像素值到[0,1]范围
     
        # 将RGB转换为HSV
        max_rgb = np.max(pixels, axis=2)
        min_rgb = np.min(pixels, axis=2)
        delta = max_rgb - min_rgb
     
        # 饱和度
        saturation = delta / (max_rgb + 1e-7)  # 防止除以零
     
        # 判断饱和度大于阈值的彩色像素
        color_pixels = saturation > saturation_threshold
        color_fraction = np.mean(color_pixels)
     
        return color_fraction > color_fraction_threshold
     
    def is_color_page(self, page):
        """
        Check if a page is a color page.
        """
        # Render page to a pixmap
        pix = page.get_pixmap()
        # Convert pixmap to an image
        img = pix.tobytes("png")
     
        # Create an image object using PIL
        from PIL import Image
        from io import BytesIO
        image = Image.open(BytesIO(img))
     
        return self.is_color_image(image)
    
    def display_page_distribution(self, color_pages, bw_pages, total_pages, is_double_sided):
        """
        显示页面分布的表格，按照原始页码顺序显示页面会被分到彩色和黑白哪个文档中
        """
        print("\n页面分布情况：")
        print("彩色打印\t黑白打印")
        print("-" * 30)
        
        for page_num in range(total_pages):
            color_page = "\t\t" 
            bw_page = "\t\t"
            
            # 确定当前页码应该显示在哪一列
            if page_num in color_pages:
                color_page = f"{page_num+1}\t\t"  # +1因为页码从1开始显示
            if page_num in bw_pages:
                bw_page = f"{page_num+1}\t\t"  # +1因为页码从1开始显示
                
            print(f"{color_page}{bw_page}")
            
            # 如果启用双面打印，每两页添加一条分隔线
            if is_double_sided and page_num % 2 == 1 and page_num < total_pages - 1:
                print("-" * 30)
     
    def split_pdf(self, input_pdf_path, output_color_pdf_path, output_bw_pdf_path, is_double_sized_printing):
        # Open the input PDF
        doc = fitz.open(input_pdf_path)
     
        # Create new PDFs for color and black & white pages
        color_doc = fitz.open()
        bw_doc = fitz.open()
     
        # Save color and bw pages number
        color_pages = []
        bw_pages = []
     
        # Iterate over each page in the input PDF
        print("正在分析PDF页面...")
        for page_num in tqdm(range(len(doc))):
            page = doc.load_page(page_num)
     
            # Check if the page is a color page
            if self.is_color_page(page):
                color_pages.append(page_num)
     
        # Handle double sized printing
        if is_double_sized_printing:
            # 改进双面打印逻辑：确保配对页也被标记为彩色
            additional_color_pages = []
            for page_num in color_pages:
                # 如果是偶数页，其配对页是下一页
                if page_num % 2 == 0:
                    paired_page = page_num + 1
                    if paired_page < len(doc) and paired_page not in color_pages:
                        additional_color_pages.append(paired_page)
                # 如果是奇数页，其配对页是上一页
                else:
                    paired_page = page_num - 1
                    if paired_page >= 0 and paired_page not in color_pages:
                        additional_color_pages.append(paired_page)
            
            # 添加所有配对页到彩色页列表
            color_pages.extend(additional_color_pages)
            # 确保列表有序且无重复
            color_pages = sorted(set(color_pages))
     
        # Insert BW Pages
        for page_num in range(len(doc)):
            if page_num not in color_pages:
                bw_pages.append(page_num)
     
        # 判断是否存在彩色页面
        if not color_pages:
            # 如果没有彩色页面，整个文档都是黑白的
            doc.close()
            return False, None, None, None, None
        
        # Insert PDF pages
        print("正在生成彩色和黑白PDF文件...")
        for page_num in sorted(color_pages):
            color_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
     
        for page_num in sorted(bw_pages):
            bw_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
     
        # Save the new PDFs
        color_doc.save(output_color_pdf_path)
        bw_doc.save(output_bw_pdf_path)
     
        # Close all documents
        doc.close()
        color_doc.close()
        bw_doc.close()
        
        return True, output_color_pdf_path, output_bw_pdf_path, color_pages, bw_pages
    
    def on_closing(self):
        """窗口关闭时的处理"""
        if self.is_processing:
            if not messagebox.askokcancel("正在处理", "任务正在处理中，确定要退出吗？"):
                return
        if hasattr(self, 'redirect'):
            self.redirect.stop_timer()
        self.destroy()


if __name__ == "__main__":
    app = PDFColorSplitter()
    app.mainloop()
