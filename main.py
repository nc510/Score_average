import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import traceback
import datetime

class ScoreAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("成绩统计分析系统")
        self.root.geometry("900x600")
        self.root.resizable(True, True)
        
        self.set_window_icon()
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
    
    def set_window_icon(self):
        try:
            from PIL import Image, ImageTk
            icon_path = os.path.join(os.path.dirname(__file__), "ico.png")
            if os.path.exists(icon_path):
                icon_image = Image.open(icon_path)
                photo = ImageTk.PhotoImage(icon_image)
                self.root.iconphoto(True, photo)
        except Exception as e:
            self.log_message(f"设置窗口图标失败: {e}", level="ERROR")
        
        self.file_path = ""
        self.df = None
        self.stats_df = None
        self.sheet_name = "题型分数"
        
        self.log_file = "score_analyzer.log"
        self.log_message("程序启动")
        
        self.create_widgets()
    
    def log_message(self, message, level="INFO"):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] [{level}] {message}\n"
        print(log_entry.strip())
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(log_entry)
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=10)
        
        self.file_label = ttk.Label(top_frame, text="未选择文件", width=50)
        self.file_label.pack(side=tk.LEFT, padx=10)
        
        self.browse_btn = ttk.Button(top_frame, text="选择Excel文件", command=self.browse_file)
        self.browse_btn.pack(side=tk.LEFT, padx=5)
        
        self.sheet_label = ttk.Label(top_frame, text="工作表:")
        self.sheet_label.pack(side=tk.LEFT, padx=5)
        
        self.sheet_combo = ttk.Combobox(top_frame, width=15, state="readonly")
        self.sheet_combo.pack(side=tk.LEFT, padx=2)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        
        self.analyze_btn = ttk.Button(top_frame, text="开始分析", command=self.analyze_data, state=tk.DISABLED)
        self.analyze_btn.pack(side=tk.LEFT, padx=5)
        
        self.export_btn = ttk.Button(top_frame, text="导出统计表", command=self.export_stats, state=tk.DISABLED)
        self.export_btn.pack(side=tk.LEFT, padx=5)
        
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        self.data_frame = ttk.Frame(notebook)
        notebook.add(self.data_frame, text="原始数据")
        
        self.stats_frame = ttk.Frame(notebook)
        notebook.add(self.stats_frame, text="统计结果")
        
        self.data_tree = ttk.Treeview(self.data_frame)
        self.data_scroll_y = ttk.Scrollbar(self.data_frame, orient=tk.VERTICAL, command=self.data_tree.yview)
        self.data_scroll_x = ttk.Scrollbar(self.data_frame, orient=tk.HORIZONTAL, command=self.data_tree.xview)
        self.data_tree.configure(yscrollcommand=self.data_scroll_y.set, xscrollcommand=self.data_scroll_x.set)
        
        self.data_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.data_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.data_tree.pack(fill=tk.BOTH, expand=True)
        
        self.stats_tree = ttk.Treeview(self.stats_frame)
        self.stats_scroll_y = ttk.Scrollbar(self.stats_frame, orient=tk.VERTICAL, command=self.stats_tree.yview)
        self.stats_scroll_x = ttk.Scrollbar(self.stats_frame, orient=tk.HORIZONTAL, command=self.stats_tree.xview)
        self.stats_tree.configure(yscrollcommand=self.stats_scroll_y.set, xscrollcommand=self.stats_scroll_x.set)
        
        self.stats_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.stats_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.stats_tree.pack(fill=tk.BOTH, expand=True)
        
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=10)
        
        self.status_label = ttk.Label(status_frame, text="状态: 就绪", relief=tk.SUNKEN)
        self.status_label.pack(fill=tk.X)
        
        self.create_menu()
    
    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于", command=self.show_about)
        help_menu.add_command(label="添加微信", command=self.show_wechat)
        help_menu.add_command(label="支持打赏", command=self.show_donation)
        menubar.add_cascade(label="帮助", menu=help_menu)
        
        self.root.config(menu=menubar)
    
    def show_about(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("关于")
        about_window.geometry("400x300")
        about_window.resizable(False, False)
        about_window.transient(self.root)
        about_window.grab_set()
        
        frame = ttk.Frame(about_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="成绩统计分析系统", font=("微软雅黑", 16, "bold")).pack(pady=10)
        ttk.Label(frame, text="版本 1.0", font=("微软雅黑", 12)).pack(pady=5)
        
        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        ttk.Label(frame, text="作者信息", font=("微软雅黑", 12, "bold")).pack(pady=5)
        ttk.Label(frame, text="作者：倾尽温柔").pack(anchor=tk.W)
        ttk.Label(frame, text="QQ：121666880").pack(anchor=tk.W)
        ttk.Label(frame, text="电话：18671151019").pack(anchor=tk.W)
        ttk.Label(frame, text="学校：咸宁市理工中等职业学校").pack(anchor=tk.W)
        
        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        ttk.Label(frame, text="功能说明：", font=("微软雅黑", 12, "bold")).pack(pady=5, anchor=tk.W)
        ttk.Label(frame, text="• 读取Excel文件中的题型分数数据").pack(anchor=tk.W)
        ttk.Label(frame, text="• 统计各题型的总分、平均分、最高分、最低分").pack(anchor=tk.W)
        ttk.Label(frame, text="• 计算及格人数、及格率、优秀人数、优秀率").pack(anchor=tk.W)
        ttk.Label(frame, text="• 导出统计结果到Excel文件").pack(anchor=tk.W)
        
        ttk.Button(frame, text="确定", command=about_window.destroy).pack(pady=15)
    
    def show_wechat(self):
        wechat_window = tk.Toplevel(self.root)
        wechat_window.title("添加微信好友")
        wechat_window.geometry("350x450")
        wechat_window.resizable(False, False)
        wechat_window.transient(self.root)
        wechat_window.grab_set()
        
        frame = ttk.Frame(wechat_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="添加微信好友", font=("微软雅黑", 14, "bold")).pack(pady=10)
        ttk.Label(frame, text="（欢迎交流技术问题）", font=("微软雅黑", 10, "italic")).pack(pady=5)
        
        qr_frame = ttk.Frame(frame, relief=tk.SUNKEN, padding=10)
        qr_frame.pack(pady=10)
        
        try:
            from PIL import Image, ImageTk
            
            qr_path = os.path.join(os.path.dirname(__file__), "加好友.png")
            if os.path.exists(qr_path):
                img = Image.open(qr_path)
                img = img.resize((200, 200), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                qr_label = ttk.Label(qr_frame, image=photo)
                qr_label.image = photo
                qr_label.pack()
                
                wechat_window.photo = photo
            else:
                ttk.Label(qr_frame, text=f"图片文件不存在: {qr_path}", foreground="red").pack()
        except Exception as e:
            self.log_message(f"微信二维码加载失败: {e}", level="ERROR")
            ttk.Label(qr_frame, text="二维码加载失败", foreground="red").pack()
        
        ttk.Label(frame, text="微信号：qingjinwenrou").pack(pady=5)
        ttk.Label(frame, text="昵称：倾尽温柔").pack(pady=2)
        
        ttk.Button(frame, text="关闭", command=wechat_window.destroy).pack(pady=15)
    
    def show_donation(self):
        donation_window = tk.Toplevel(self.root)
        donation_window.title("支持打赏")
        donation_window.geometry("350x450")
        donation_window.resizable(False, False)
        donation_window.transient(self.root)
        donation_window.grab_set()
        
        frame = ttk.Frame(donation_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="支持开发者", font=("微软雅黑", 14, "bold")).pack(pady=10)
        ttk.Label(frame, text="（您的支持是我前进的动力）", font=("微软雅黑", 10, "italic")).pack(pady=5)
        
        qr_frame = ttk.Frame(frame, relief=tk.SUNKEN, padding=10)
        qr_frame.pack(pady=10)
        
        try:
            from PIL import Image, ImageTk
            
            qr_path = os.path.join(os.path.dirname(__file__), "打赏.png")
            if os.path.exists(qr_path):
                img = Image.open(qr_path)
                img = img.resize((200, 200), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                qr_label = ttk.Label(qr_frame, image=photo)
                qr_label.image = photo
                qr_label.pack()
                
                donation_window.photo = photo
            else:
                ttk.Label(qr_frame, text=f"图片文件不存在: {qr_path}", foreground="red").pack()
        except Exception as e:
            self.log_message(f"打赏二维码加载失败: {e}", level="ERROR")
            ttk.Label(qr_frame, text="二维码加载失败", foreground="red").pack()
        
        ttk.Label(frame, text="微信扫码打赏", font=("微软雅黑", 12)).pack(pady=5)
        ttk.Label(frame, text="感谢您的支持！", font=("微软雅黑", 10, "italic")).pack(pady=2)
        
        ttk.Button(frame, text="关闭", command=donation_window.destroy).pack(pady=15)
    
    def browse_file(self):
        self.log_message("开始选择Excel文件")
        self.file_path = filedialog.askopenfilename(
            title="选择成绩Excel文件",
            filetypes=[("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")]
        )
        
        if self.file_path:
            self.log_message(f"文件选择成功: {self.file_path}")
            self.log_message(f"文件是否存在: {os.path.exists(self.file_path)}")
            self.log_message(f"文件大小: {os.path.getsize(self.file_path) if os.path.exists(self.file_path) else '未知'} 字节")
            
            file_ext = os.path.splitext(self.file_path)[1].lower()
            self.log_message(f"文件扩展名: {file_ext}")
            
            if file_ext not in ['.xlsx', '.xls']:
                self.log_message(f"警告: 文件扩展名不是标准Excel格式 ({file_ext})", "WARNING")
            
            try:
                xls = pd.ExcelFile(self.file_path)
                sheet_names = xls.sheet_names
                self.log_message(f"工作表列表: {sheet_names}")
                self.sheet_combo["values"] = sheet_names
                if self.sheet_name in sheet_names:
                    self.sheet_combo.set(self.sheet_name)
                else:
                    self.sheet_combo.set(sheet_names[0])
                    self.sheet_name = sheet_names[0]
            except Exception as e:
                self.log_message(f"读取工作表列表失败: {str(e)}", "ERROR")
                self.sheet_combo.set("")
            
            self.file_label.config(text=os.path.basename(self.file_path))
            self.analyze_btn.config(state=tk.NORMAL)
            self.status_label.config(text="状态: 已选择文件")
        else:
            self.log_message("用户取消选择文件")
            self.file_label.config(text="未选择文件")
            self.analyze_btn.config(state=tk.DISABLED)
            self.sheet_combo.set("")
    
    def on_sheet_selected(self, event):
        self.sheet_name = self.sheet_combo.get()
        self.log_message(f"切换到工作表: {self.sheet_name}")
    
    def analyze_data(self):
        try:
            self.log_message("开始分析数据")
            self.log_message(f"当前文件路径: {self.file_path}")
            
            if not self.file_path:
                self.log_message("错误: 文件路径为空", "ERROR")
                messagebox.showerror("错误", "请先选择文件")
                return
            
            if not os.path.exists(self.file_path):
                self.log_message(f"错误: 文件不存在 - {self.file_path}", "ERROR")
                messagebox.showerror("错误", f"文件不存在:\n{self.file_path}")
                return
            
            self.status_label.config(text="状态: 正在读取数据...")
            self.root.update()
            
            self.log_message(f"开始读取Excel文件，工作表: {self.sheet_name}")
            try:
                self.df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)
                self.log_message(f"工作表 '{self.sheet_name}' 读取成功")
            except Exception as excel_error:
                self.log_message(f"Excel读取失败: {str(excel_error)}", "ERROR")
                self.log_message(f"完整错误信息:\n{traceback.format_exc()}", "ERROR")
                raise
            
            self.log_message(f"数据行数: {len(self.df)}")
            self.log_message(f"数据列数: {len(self.df.columns)}")
            self.log_message(f"列名: {list(self.df.columns)}")
            
            numeric_cols = self.df.select_dtypes(include=['number']).columns.tolist()
            self.log_message(f"数值列（用于计算）: {numeric_cols}")
            
            if not numeric_cols:
                self.log_message("警告: 未找到数值列，无法进行统计计算", "WARNING")
            
            self.clear_tree(self.data_tree)
            self.data_tree["columns"] = list(self.df.columns)
            self.data_tree["show"] = "headings"
            
            for col in self.df.columns:
                self.data_tree.heading(col, text=col)
                self.data_tree.column(col, width=100)
            
            for _, row in self.df.iterrows():
                self.data_tree.insert("", tk.END, values=list(row))
            
            self.log_message("原始数据已显示到表格")
            
            self.status_label.config(text="状态: 正在计算统计数据...")
            self.root.update()
            
            self.log_message("开始计算统计数据...")
            self.calculate_stats()
            self.log_message("统计数据计算完成")
            
            self.export_btn.config(state=tk.NORMAL)
            self.status_label.config(text="状态: 分析完成")
            self.log_message("数据分析完成")
            
        except Exception as e:
            self.log_message(f"分析数据时发生错误: {str(e)}", "ERROR")
            self.log_message(f"完整错误堆栈:\n{traceback.format_exc()}", "ERROR")
            messagebox.showerror("错误", f"读取文件失败: {str(e)}\n\n详细信息请查看日志文件")
            self.status_label.config(text="状态: 读取失败")
    
    def calculate_stats(self):
        PASS_LINE = 60
        EXCELLENT_LINE = 80
        
        id_cols = ['考号', '姓名', '学号']
        numeric_cols = self.df.select_dtypes(include=['number']).columns.tolist()
        
        question_cols = [col for col in numeric_cols if col not in id_cols and col != '成绩']
        has_total = '成绩' in numeric_cols
        total_cols = ['成绩'] + question_cols if has_total else question_cols
        
        stats_data = []
        headers = ["题型", "总分", "平均分", "最高分", "最低分", "人数", 
                   "及格人数", "及格率", "优秀人数", "优秀率"]
        stats_data.append(headers)
        
        for col in total_cols:
            total = self.df[col].sum()
            avg = self.df[col].mean()
            max_val = self.df[col].max()
            min_val = self.df[col].min()
            count = self.df[col].count()
            
            pass_count = len(self.df[self.df[col] >= max_val * PASS_LINE / 100])
            pass_rate = pass_count / count * 100 if count > 0 else 0
            
            excellent_count = len(self.df[self.df[col] >= max_val * EXCELLENT_LINE / 100])
            excellent_rate = excellent_count / count * 100 if count > 0 else 0
            
            stats_data.append([col, round(total, 2), round(avg, 2), round(max_val, 2), 
                              round(min_val, 2), int(count), int(pass_count), 
                              f"{round(pass_rate, 1)}%", int(excellent_count), 
                              f"{round(excellent_rate, 1)}%"])
        
        if has_total:
            avg_total = self.df['成绩'].mean()
            stats_data.append(["成绩平均分", "", round(avg_total, 2), "", "", "", "", "", "", ""])
        
        self.stats_df = pd.DataFrame(stats_data[1:], columns=stats_data[0])
        
        self.clear_tree(self.stats_tree)
        self.stats_tree["columns"] = headers
        self.stats_tree["show"] = "headings"
        
        col_widths = [90, 70, 70, 70, 70, 50, 70, 60, 70, 60]
        for i, col in enumerate(headers):
            self.stats_tree.heading(col, text=col)
            self.stats_tree.column(col, width=col_widths[i])
        
        for row in stats_data[1:]:
            self.stats_tree.insert("", tk.END, values=row)
    
    def clear_tree(self, tree):
        for item in tree.get_children():
            tree.delete(item)
    
    def export_stats(self):
        if self.stats_df is None:
            messagebox.showwarning("警告", "请先进行分析")
            return
        
        default_dir = os.path.expanduser("~/Desktop")
        if not os.path.exists(default_dir):
            default_dir = os.path.dirname(self.file_path) if self.file_path else "."
        
        output_path = filedialog.asksaveasfilename(
            title="保存统计表",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialdir=default_dir,
            initialfile="题型分数统计表.xlsx"
        )
        
        if output_path:
            try:
                self.stats_df.to_excel(output_path, index=False)
                messagebox.showinfo("成功", f"统计表已保存到:\n{output_path}")
                self.status_label.config(text=f"状态: 已导出到 {os.path.basename(output_path)}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ScoreAnalyzer(root)
    root.mainloop()