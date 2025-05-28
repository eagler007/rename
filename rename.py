import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import ctypes
import platform
from mutagen.easyid3 import EasyID3
from mutagen.mp3 import MP3
from mutagen.mp4 import MP4
from mutagen.id3 import TPE1, TIT2
from mutagen.flac import FLAC

class FileRenamerApp:
    def __init__(self, root):
        self.root = root
        self.setup_dpi_awareness()
        
        # 获取系统信息
        self.screen_width = root.winfo_screenwidth()
        self.screen_height = root.winfo_screenheight()
        self.os_type = platform.system()
        
        # 初始化UI缩放因子
        self.ui_scale = self.calculate_ui_scale()
        
        # 设置窗口初始大小
        self.setup_initial_window()
        
        # 主界面布局
        self.setup_ui()
        
    def setup_dpi_awareness(self):
        """设置DPI感知"""
        if platform.system() == "Windows":
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(2)
            except:
                try:
                    ctypes.windll.user32.SetProcessDPIAware()
                except:
                    pass
    
    def calculate_ui_scale(self):
        """计算UI缩放因子"""
        base_dpi = 96
        base_width = 1920
        
        if self.os_type == "Windows":
            try:
                hdc = ctypes.windll.user32.GetDC(0)
                dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
                ctypes.windll.user32.ReleaseDC(0, hdc)
                return dpi / base_dpi
            except:
                pass
        
        resolution_scale = self.screen_width / base_width
        return max(0.8, min(2.0, resolution_scale * 1.1))
    
    def setup_initial_window(self):
        """设置初始窗口大小"""
        if self.screen_width >= 3840:
            width, height = int(1000 * self.ui_scale), int(600 * self.ui_scale)
        elif self.screen_width >= 2560:
            width, height = int(900 * self.ui_scale), int(550 * self.ui_scale)
        else:
            width, height = int(850 * self.ui_scale), int(500 * self.ui_scale)
        
        self.root.geometry(f"{width}x{height}")
        self.root.minsize(int(800 * self.ui_scale), int(450 * self.ui_scale))
        self.root.title("高级文件重命名工具")
    
    def scaled(self, value):
        """根据缩放因子调整尺寸"""
        return int(value * self.ui_scale)
    
    def setup_ui(self):
        """初始化用户界面"""
        # 主框架
        self.main_frame = ttk.Frame(self.root, padding=self.scaled(10))
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧功能区域
        left_frame = ttk.Frame(self.main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 右侧说明区域
        right_frame = ttk.Frame(self.main_frame, width=self.scaled(300))
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(self.scaled(10), 0))
        
        # 配置左侧网格权重
        left_frame.grid_rowconfigure(3, weight=1)
        left_frame.grid_columnconfigure(0, weight=1)
        
        # ================= 字体样式配置 =================
        style = ttk.Style()
        
        # 基础字体配置
        base_font = ('Microsoft YaHei', 12, 'bold')
        
        # 全局默认样式
        style.configure('.', font=base_font)
        
        # 特别控件样式
        style.configure('TButton', 
                       font=('Microsoft YaHei', 12, 'bold'),
                       padding=6)
        
        style.configure('TLabel', font=base_font)
        style.configure('TEntry', font=base_font)
        
        # 标题特别样式
        style.configure('Title.TLabel', 
                       font=('Microsoft YaHei', 14, 'bold'))
        
        # 当前行计数器
        current_row = 0
        
        # 文件夹选择
        self.setup_folder_selector(left_frame, current_row)
        current_row += 1
        
        # 功能按钮
        self.setup_function_buttons(left_frame, current_row)
        current_row += 1
        
        # 预览区域
        self.setup_preview_area(left_frame, current_row)
        current_row += 1
        
        # 操作按钮
        self.setup_action_buttons(left_frame, current_row)
        
        # 添加软件使用说明
        self.setup_usage_instructions(right_frame)
        
        # 确保窗口内容适应
        self.root.update_idletasks()
    
    def setup_folder_selector(self, parent, row):
        """文件夹选择组件"""
        frame = ttk.LabelFrame(parent, text="选择文件夹")
        frame.grid(row=row, column=0, sticky="ew", pady=(0, self.scaled(5)))
        
        self.folder_path = tk.StringVar()
        entry = ttk.Entry(frame, textvariable=self.folder_path)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=self.scaled(5))
        
        ttk.Button(frame, text="浏览...", command=self.select_folder).pack(side=tk.RIGHT)
    
    def setup_function_buttons(self, parent, row):
        """功能按钮"""
        frame = ttk.LabelFrame(parent, text="功能选项")
        frame.grid(row=row, column=0, sticky="ew", pady=(0, self.scaled(5)))
        
        buttons = [
            ("提取章节信息", self.extract_chapter_info),
            ("统一编号格式", self.unify_number_format),
            ("检查缺失集数", self.check_missing_episodes),
            ("同步音频标题", self.sync_audio_titles)
        ]
        
        for i, (text, command) in enumerate(buttons):
            ttk.Button(frame, text=text, command=command).grid(
                row=i//2, column=i%2, padx=self.scaled(5), pady=self.scaled(5), sticky="ew"
            )
        
        # 配置按钮均匀分布
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)
    
    def setup_preview_area(self, parent, row):
        """预览区域"""
        frame = ttk.LabelFrame(parent, text="文件预览 (原文件名 → 新文件名)")
        frame.grid(row=row, column=0, sticky="nsew", pady=(0, self.scaled(5)))
        
        # 使用Treeview组件实现双列显示
        self.preview_tree = ttk.Treeview(frame, columns=('original', 'new'), show='headings', height=10)
        self.preview_tree.pack(fill=tk.BOTH, expand=True)
        
        # 设置列标题
        self.preview_tree.heading('original', text='原文件名')
        self.preview_tree.heading('new', text='新文件名')
        
        # 设置列宽
        self.preview_tree.column('original', width=self.scaled(200), anchor='w')
        self.preview_tree.column('new', width=self.scaled(200), anchor='w')
        
        # 添加滚动条
        #scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.preview_tree.yview)
        #scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        #self.preview_tree.configure(yscrollcommand=scrollbar.set)
    
    def setup_action_buttons(self, parent, row):
        """操作按钮"""
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=0, sticky="ew", pady=(self.scaled(5), 0))
        
        buttons = [
            ("预览更改", self.preview_changes),
            ("执行重命名", self.execute_rename),
            ("退出", self.root.quit)
        ]
        
        for text, command in buttons:
            ttk.Button(frame, text=text, command=command).pack(side=tk.LEFT, padx=self.scaled(5), expand=True)
    
    def setup_usage_instructions(self, parent):
        """添加软件使用说明"""
        #动态拉伸
        frame = ttk.LabelFrame(parent, text="使用说明", padding=self.scaled(10))
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建带滚动条的文本框
        text_frame = ttk.Frame(frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.instructions = tk.Text(    # 改为实例变量
            text_frame,
            wrap=tk.WORD,
            yscrollcommand=scrollbar.set,
            padx=self.scaled(5),
            pady=self.scaled(5),
            font=('Microsoft YaHei', 10)
        )
        self.instructions.pack(fill=tk.Y, expand=True)
        scrollbar.config(command=self.instructions.yview)
  
        # 添加使用说明内容
        instructions_content = """
【功能说明】

1. 提取章节信息：
   - 只保留"第XXX章/节/集"部分
   - 删除文件名中其他所有内容
   - 保留文件后缀名不变

2. 统一编号格式：
   - 将编号统一为4位数
   - 例如：第1集 → 第0001集
   - 第100集 → 第0100集

3. 检查缺失集数：
   - 自动检测文件编号连续性
   - 弹出窗口显示缺失的集数
   - 支持章/节/集多种格式

4. 同步音频标题：
   - 将MP3/M4A/FLAC文件名同步到音频元数据
   - 仅修改标题，不影响其他元数据
   - 需要安装mutagen库

【操作步骤】

1. 选择目标文件夹
2. 点击相应功能按钮
3. 预览更改效果
4. 确认无误后执行重命名

【注意事项】

• 操作前建议备份重要文件
• 重命名操作不可逆
• 音频功能需要mutagen库
• 高分辨率屏幕自动适配
"""
        self.instructions.insert(tk.END, instructions_content)
        # 设置大类标题为粗体并标红
        self.instructions.tag_configure(
            "title", 
            foreground="red", 
            font=('Microsoft YaHei', 10, 'bold'),
            spacing1=10,  # 段前间距
            spacing2=5    # 段间间距
        )
        
        # 查找并标记所有大类标题
        for title in ["【功能说明】", "【操作步骤】", "【注意事项】"]:
            start = "1.0"
            while True:
                pos = self.instructions.search(title, start, stopindex=tk.END)
                if not pos:
                    break
                end = f"{pos}+{len(title)}c"
                self.instructions.tag_add("title", pos, end)     
                start = end
        # 配置列表项样式
        self.instructions.tag_configure(
            "list", 
            lmargin1=20,  # 一级缩进
            lmargin2=30,  # 二级缩进
            foreground="#333333"
        )
        
        # 标记所有列表项
        for line in instructions_content.split('\n'):
            if line.strip().startswith(('•', '-', '1.', '2.', '3.', '4.')):
                start = "1.0"
                while True:
                    pos = self.instructions.search(line.strip(), start, stopindex=tk.END)
                    if not pos:
                        break
                    end = f"{pos}+{len(line)}c"
                    self.instructions.tag_add("list", pos, end)
                    start = end
        
        self.instructions.config(state='disabled')
    
    def select_folder(self):
        """选择文件夹"""
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)
    
    def get_files(self):
        """获取文件夹中的文件列表"""
        folder = self.folder_path.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("警告", "请先选择有效文件夹")
            return None
        
        return [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    
    def clear_preview(self):
        """清空预览区域"""
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
    
    def preview_changes(self):
        """预览更改"""
        files = self.get_files()
        if files is None:
            return
        
        self.clear_preview()
        
        for filename in sorted(files):
            new_name = self.process_filename(filename)
            if new_name != filename:
                self.preview_tree.insert('', 'end', values=(filename, new_name))
    
    def process_filename(self, filename):
        """处理文件名"""
        # 提取基本名称和扩展名
        basename, ext = os.path.splitext(filename)
        
        # 尝试匹配章/节/集
        match = re.search(r'(第[零一二三四五六七八九十百千万\d]+[章节集])', basename)
        if match:
            return match.group(1) + ext
        
        return filename
    
    def extract_chapter_info(self):
        """提取章节信息"""
        self.preview_changes()
    
    def unify_number_format(self):
        """统一编号格式"""
        files = self.get_files()
        if files is None:
            return
        
        self.clear_preview()
        
        for filename in sorted(files):
            basename, ext = os.path.splitext(filename)
            
            # 匹配数字编号
            match = re.search(r'(第)(\d+)([章节集])', basename)
            if match:
                prefix = match.group(1)
                number = match.group(2).zfill(4)  # 补齐4位
                suffix = match.group(3)
                new_name = f"{prefix}{number}{suffix}{ext}"
                self.preview_tree.insert('', 'end', values=(filename, new_name))
            else:
                # 如果不是数字编号，保持原样
                self.preview_tree.insert('', 'end', values=(filename, filename))
    
    def check_missing_episodes(self):
        """检查缺失集数"""
        files = self.get_files()
        if files is None:
            return
        
        numbers = set()
        
        for filename in files:
            basename = os.path.splitext(filename)[0]
            
            # 匹配数字编号
            match = re.search(r'第(\d+)[章节集]', basename)
            if match:
                numbers.add(int(match.group(1)))
        
        if not numbers:
            messagebox.showinfo("检查结果", "没有找到带编号的文件")
            return
        
        min_num = min(numbers)
        max_num = max(numbers)
        full_set = set(range(min_num, max_num + 1))
        missing = sorted(full_set - numbers)
        
        if not missing:
            messagebox.showinfo("检查结果", f"编号从{min_num}到{max_num}，没有缺失")
        else:
            # 创建弹出窗口显示缺失章节
            dialog = tk.Toplevel(self.root)
            dialog.title("缺失集数列表")
            dialog.geometry(f"{self.scaled(400)}x{self.scaled(300)}")
            
            # 主框架
            main_frame = ttk.Frame(dialog)
            main_frame.pack(fill=tk.BOTH, expand=True, padx=self.scaled(10), pady=self.scaled(10))
            
            # 标题
            ttk.Label(
                main_frame, 
                text=f"共缺失 {len(missing)} 集 (编号{min_num}-{max_num})", 
                font=('Microsoft YaHei', 10, 'bold')
            ).pack(pady=(0, self.scaled(10)))
            
            # 文本框显示缺失章节
            text_frame = ttk.Frame(main_frame)
            text_frame.pack(fill=tk.BOTH, expand=True)
            
            scrollbar = ttk.Scrollbar(text_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            text = tk.Text(
                text_frame,
                wrap=tk.WORD,
                yscrollcommand=scrollbar.set,
                padx=self.scaled(5),
                pady=self.scaled(5),
                font=('Microsoft YaHei', 10)
            )
            text.pack(fill=tk.BOTH, expand=True)
            scrollbar.config(command=text.yview)
            
            # 添加缺失章节
            text.insert(tk.END, "\n".join([f"第{n}集" for n in missing]))
            text.config(state='disabled')
            
            # 关闭按钮
            ttk.Button(
                main_frame,
                text="关闭",
                command=dialog.destroy
            ).pack(pady=(self.scaled(10), 0))
    
    def sync_audio_titles(self):
        """同步音频标题（优化版）"""
        folder = self.folder_path.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("警告", "请先选择有效文件夹")
            return

        try:
            # 支持的音频文件扩展名
            supported_ext = {'.mp3', '.m4a', '.flac'}  # 添加了FLAC支持
            audio_files = [
                f for f in os.listdir(folder) 
                if os.path.splitext(f)[1].lower() in supported_ext
                and os.path.isfile(os.path.join(folder, f))
            ]
            
            if not audio_files:
                messagebox.showwarning("警告", "没有找到支持的音频文件(支持MP3/M4A/FLAC)")
                return
            
            # 更详细的确认对话框
            confirm = messagebox.askyesno(
                "确认操作",
                f"即将修改 {len(audio_files)} 个音频文件的元数据:\n"
                "• 将文件名(不含扩展名)设为标题\n"
                "• MP3文件同时设置艺术家标签\n\n"
                "此操作不可逆，建议先备份文件\n\n"
                "确定要继续吗?"
            )
            
            if not confirm:
                return
            
            success_count = 0
            failed_files = []
            
            for filename in audio_files:
                try:
                    filepath = os.path.join(folder, filename)
                    basename = os.path.splitext(filename)[0]
                    ext = os.path.splitext(filename)[1].lower()
                    
                    if ext == '.mp3':
                        # MP3文件处理（使用ID3v2.4标准）
                        audio = MP3(filepath)
                        
                        # 确保标签存在
                        if audio.tags is None:
                            audio.add_tags(ID3=ID3)
                        
                        # 删除旧标题和艺术家标签（如果存在）
                        for tag in ['TIT2', 'TPE1']:
                            if tag in audio.tags:
                                del audio.tags[tag]
                        
                        # 添加新标签（UTF-8编码）
                        audio.tags.add(TIT2(encoding=3, text=basename))  # 标题
                        audio.tags.add(TPE1(encoding=3, text=basename))  # 艺术家
                        
                        # 强制保存为ID3v2.4
                        audio.tags.save(filepath, v2_version=4)
                        
                    elif ext == '.m4a':
                        # M4A文件处理
                        audio = MP4(filepath)
                        
                        # 确保标签存在
                        if audio.tags is None:
                            audio.add_tags()
                        
                        # 设置标题标签
                        audio.tags["\xa9nam"] = [basename]  # 标题（注意这里是列表）
                        audio.tags["\xa9ART"] = [basename]  # 艺术家（可选）
                        audio.save()
                    
                    elif ext == '.flac':
                        # FLAC文件处理
                        audio = FLAC(filepath)
                        
                        # 确保标签存在
                        if not audio.tags:
                            audio.add_tags()
                        
                        # 设置标题标签
                        audio.tags["TITLE"] = basename
                        audio.save()
                    
                    success_count += 1
                    
                except Exception as e:
                    # 获取更详细的错误信息
                    error_detail = f"{type(e).__name__}: {str(e)}"
                    failed_files.append(f"{filename} ({error_detail})")
            
            # 更友好的结果展示
            result_msg = [
                f"操作完成:",
                f"✓ 成功处理: {success_count} 个文件",
                f"✗ 处理失败: {len(failed_files)} 个文件"
            ]
            
            if failed_files:
                result_msg.append("\n失败详情:")
                result_msg.extend(failed_files)
            
            # 使用ScrolledText显示可能的长结果
            self.show_result_dialog("同步结果", "\n".join(result_msg))
            
        except ImportError as e:
            messagebox.showwarning(
                "缺少依赖",
                "音频元数据功能需要安装mutagen库\n\n"
                "请运行以下命令安装:\n"
                "pip install mutagen\n\n"
                f"错误详情: {str(e)}"
            )
        except Exception as e:
            messagebox.showerror(
                "意外错误",
                f"发生未预期的错误:\n\n{type(e).__name__}: {str(e)}"
            )

    def show_result_dialog(self, title, message):
        """显示带滚动条的结果对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry(f"{self.scaled(600)}x{self.scaled(400)}")
        
        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=self.scaled(10), pady=self.scaled(10))
        
        text_frame = ttk.Frame(frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text = tk.Text(
            text_frame,
            wrap=tk.WORD,
            yscrollcommand=scrollbar.set,
            padx=self.scaled(5),
            pady=self.scaled(5),
            font=('Microsoft YaHei', 10)
        )
        text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=text.yview)
        
        text.insert(tk.END, message)
        text.config(state='disabled')
        
        ttk.Button(
            frame,
            text="关闭",
            command=dialog.destroy
        ).pack(pady=(self.scaled(10), 0))

    
    def execute_rename(self):
        """执行重命名"""
        # 获取预览中的所有项目
        items = self.preview_tree.get_children()
        if not items:
            messagebox.showwarning("警告", "没有可执行的重命名操作")
            return
        
        folder = self.folder_path.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("警告", "请先选择有效文件夹")
            return
        
        # 确认对话框
        confirm = messagebox.askyesno(
            "确认重命名",
            f"即将重命名 {len(items)} 个文件\n"
            "此操作不可逆，建议先备份文件\n\n"
            "确定要继续吗?"
        )
        
        if not confirm:
            return
        
        try:
            success_count = 0
            failed_files = []
            
            for item in items:
                original, new_name = self.preview_tree.item(item, 'values')
                
                try:
                    src = os.path.join(folder, original)
                    dst = os.path.join(folder, new_name)
                    
                    # 检查目标文件是否已存在
                    if os.path.exists(dst):
                        failed_files.append(f"{original} → {new_name} (目标文件已存在)")
                        continue
                    
                    os.rename(src, dst)
                    success_count += 1
                except Exception as e:
                    failed_files.append(f"{original} → {new_name} (错误: {str(e)})")
            
            # 显示结果
            result_msg = f"成功重命名 {success_count} 个文件"
            if failed_files:
                result_msg += f"\n\n处理失败的文件:\n" + "\n".join(failed_files)
            
            messagebox.showinfo("重命名结果", result_msg)
            
            # 刷新预览
            self.preview_changes()
        
        except Exception as e:
            messagebox.showerror("错误", f"重命名过程中发生错误: {str(e)}")
if __name__ == "__main__":
    root = tk.Tk()
    app = FileRenamerApp(root)
    root.mainloop()