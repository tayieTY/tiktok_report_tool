import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import threading
import sys
import os
import time
import random
import traceback

import user_report
import video_report

# 必须在顶部导入这两个模块，否则会报 NameError
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    # 为了防止因为没安装库直接闪退，这里先给个空值，后面主程序会拦截报错
    DND_FILES = None
    TkinterDnD = None

# 引入 ImageDraw 用来画那个正方形
from PIL import Image, ImageTk, ImageDraw

# =========================================================================
# 全局配置
# =========================================================================
ASCII_LOGO = "抖音自动化举报工具箱 (v4.0 Pro)"

ANNOUNCEMENT_TEXT = """📢 系统公告 (v4.0)：

1. 【开源声明】本软件为免费开源项目，仅供个人学习与技术研究，严禁用于任何商业用途。

2. 【获取源码】
   • GitHub地址：https://github.com/tayieTY/tiktok_report_tool 
   • 作者邮箱：3031439300@qq.com or tayie3257@gmail.com
   (如无法访问GitHub，请通过邮件联系作者获取源码)

3. 【使用须知】
   • 提前下载Google Chrome浏览器，登录个人抖音账号网页版。
   • 在模板excel里按分类依次录入举报的账号（视频）、举报类型、举报理由。
   • 支持拖拽文件到 Excel 路径输入框。
   • 点击打开屏幕设置，进入系统修改分辨率和缩放后（屏幕分辨率 1920x1080，缩放 100% ）才可以正常使用。
   • 双击运行“批量举报助手”，在路径导入指定excel文件，并点击“开始执行任务”即可实现批量自动举报。
   • 软件运行期间请勿移动鼠标和触碰键盘。
   • 若电脑性能不足，请慎重选择极速模式。

4. 【问题与不足】
   • 举报用户时，“举报原因”只有“内容违规”，“举报类型”只有“不实信息”“色情低俗”“政治敏感”三种选择。
   • 举报视频时，“举报原因”只有“不实信息”“色情低俗”“政治敏感“三种选择，但是三个大类下的子类“举报类型”是完善的。

⚠️ 免责声明：请合理合法使用本工具，用户使用本工具产生的一切后果由用户自行承担。
⚠️ 提示：如果窗口底部被任务栏遮挡，请尝试隐藏任务栏或手动拉伸。
"""


# =========================================================================
# 辅助类与函数
# =========================================================================

class TaskStoppedError(Exception):
    pass


def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class Logger(object):
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        if self.text_widget:
            self.text_widget.after(0, self._append_text, string)

    def _append_text(self, string):
        try:
            self.text_widget.config(state="normal")
            self.text_widget.insert("end", string)
            self.text_widget.see("end")
            self.text_widget.config(state="disabled")
        except:
            pass

    def flush(self):
        pass


# =========================================================================
# 主程序类
# =========================================================================

class ReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("抖音自动化举报工具箱 v4.0 Pro")
        self.root.geometry("1100x980")
        self.root.resizable(False, True)

        # 1. 初始化资源 (画正方形图标)
        self.setup_custom_assets()

        # 2. 样式初始化
        self.setup_styles()
        self.last_path_val = ""
        self.last_v_sheet = ""
        self.last_u_sheet = ""

        try:
            icon_path = get_resource_path(os.path.join("images", "1.ico"))
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass

        self.original_sleep = time.sleep
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.pause_event.set()
        self.is_running = False

        # 3. 顶部导航
        self.nav_frame = ttk.Frame(root, style="Nav.TFrame")
        self.nav_frame.pack(side="top", fill="x")

        self.btn_nav_work = ttk.Button(self.nav_frame, text="🖥️ 工作控制台", command=self.show_work_area,
                                       style="NavSelected.TButton")
        self.btn_nav_work.pack(side="left", fill="y", ipadx=20, ipady=15)

        self.btn_nav_notice = ttk.Button(self.nav_frame, text="📜 系统公告", command=self.show_notice_area,
                                         style="NavNormal.TButton")
        self.btn_nav_notice.pack(side="left", fill="y", ipadx=20, ipady=15)

        # 4. 内容容器
        self.content_container = ttk.Frame(root, style="Main.TFrame")
        self.content_container.pack(side="bottom", fill="both", expand=True)

        self.frame_work = ttk.Frame(self.content_container, style="Main.TFrame")
        self.setup_work_ui()

        self.frame_notice = ttk.Frame(self.content_container, style="Main.TFrame")
        self.setup_notice_ui()

        self.show_work_area()
        print(">>> 系统初始化完成，等待指令...")

    # ==========================================
    # 🎨 资源生成 (手绘正方形图标)
    # ==========================================
    def setup_custom_assets(self):
        """生成自定义的复选框图标（正方形）"""
        # 1. 未选中状态：灰色边框空心正方形
        img_off = Image.new('RGBA', (24, 24), (255, 255, 255, 0))  # 透明背景
        draw_off = ImageDraw.Draw(img_off)
        # 画一个圆角矩形框 (或者纯方框)
        draw_off.rectangle([2, 2, 21, 21], outline="#999999", width=2)

        # 2. 选中状态：蓝色边框 + 蓝色实心中心
        img_on = Image.new('RGBA', (24, 24), (255, 255, 255, 0))
        draw_on = ImageDraw.Draw(img_on)
        draw_on.rectangle([2, 2, 21, 21], outline="#1890FF", width=2)  # 蓝框
        draw_on.rectangle([6, 6, 17, 17], fill="#1890FF")  # 实心芯

        self.icon_chk_off = ImageTk.PhotoImage(img_off)
        self.icon_chk_on = ImageTk.PhotoImage(img_on)

    # ==========================================
    # 🎨 样式定义
    # ==========================================
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

        # 调色板
        BG_COLOR = "#F0F2F5"
        CARD_BG = "#FFFFFF"
        PRIMARY = "#1890FF"
        TEXT_MAIN = "#333333"

        # 字体设定
        FONT_NORMAL = ("Microsoft YaHei UI", 12)
        FONT_BOLD = ("Microsoft YaHei UI", 12, "bold")
        FONT_TITLE = ("Microsoft YaHei UI", 14, "bold")
        FONT_LARGE = ("Microsoft YaHei UI", 16, "bold")

        # 样式配置
        style.configure("Main.TFrame", background=BG_COLOR)
        style.configure("Card.TFrame", background=CARD_BG, relief="flat", borderwidth=0)
        style.configure("Card.TLabelframe", background=CARD_BG, relief="flat", borderwidth=1)
        style.configure("Card.TLabelframe.Label", font=FONT_TITLE, foreground=PRIMARY, background=CARD_BG)

        style.configure("Nav.TFrame", background="#FFFFFF")
        style.configure("NavSelected.TButton", font=FONT_TITLE, background="#E6F7FF", foreground=PRIMARY, borderwidth=0)
        style.configure("NavNormal.TButton", font=("Microsoft YaHei UI", 14), background="#FFFFFF",
                        foreground="#666666", borderwidth=0)

        style.configure("TLabel", background=CARD_BG, foreground=TEXT_MAIN, font=FONT_NORMAL)
        style.configure("Gray.TLabel", background=BG_COLOR, foreground=TEXT_MAIN, font=FONT_NORMAL)

        style.configure("TButton", font=FONT_NORMAL, borderwidth=1, background="#FFFFFF")
        style.map("TButton", background=[('active', '#E6F7FF')], foreground=[('active', PRIMARY)])

        style.configure("Big.TRadiobutton", font=FONT_BOLD, background=CARD_BG, foreground=TEXT_MAIN, indicatorwidth=18)

        # 按钮样式
        style.configure("Start.TButton", font=FONT_LARGE, background="#F6FFED", foreground="#389E0D")
        style.map("Start.TButton", background=[('active', '#52C41A')], foreground=[('active', 'white')])

        style.configure("Pause.TButton", font=FONT_LARGE, background="#FFF7E6", foreground="#D46B08")
        style.map("Pause.TButton", background=[('active', '#FA8C16')], foreground=[('active', 'white')])

        style.configure("Stop.TButton", font=FONT_LARGE, background="#FFF1F0", foreground="#CF1322")
        style.map("Stop.TButton", background=[('active', '#FF4D4F')], foreground=[('active', 'white')])

    # ==========================================
    # 界面布局
    # ==========================================
    def setup_work_ui(self):
        main_layout = ttk.Frame(self.frame_work, style="Main.TFrame")
        main_layout.pack(fill="both", expand=True, padx=20, pady=20)

        # === 上半部分：左右分栏 ===
        top_section = ttk.Frame(main_layout, style="Main.TFrame")
        top_section.pack(side="top", fill="x", anchor="n")

        # [左侧] 操作区
        left_col = ttk.Frame(top_section, style="Main.TFrame")
        left_col.pack(side="left", fill="both", expand=True, padx=(0, 20))

        self.create_card_config(left_col)
        self.create_card_mode(left_col)
        self.create_card_control(left_col)

        # [右侧] 图片区
        right_col = ttk.Frame(top_section, style="Main.TFrame")
        right_col.pack(side="right", fill="y")
        self.create_card_images(right_col)

        # === 下半部分：日志区 ===
        log_card = ttk.LabelFrame(main_layout, text=" 📟 实时运行日志 ", style="Card.TLabelframe", padding=15)
        log_card.pack(side="bottom", fill="both", expand=True, pady=(20, 0))

        self.log_text = scrolledtext.ScrolledText(log_card, state="disabled",
                                                  bg="#1E1E1E", fg="#52C41A",
                                                  font=("Consolas", 11), borderwidth=0,
                                                  insertbackground="white")
        self.log_text.pack(fill="both", expand=True)

        sys.stdout = Logger(self.log_text)
        sys.stderr = Logger(self.log_text)

    # ----------------------------------------
    # 子组件拆分
    # ----------------------------------------
    def create_card_config(self, parent):
        card = ttk.LabelFrame(parent, text=" 🛠️ 基础配置 ", style="Card.TLabelframe", padding=20)
        card.pack(fill="x", pady=(0, 15))

        row1 = ttk.Frame(card, style="Card.TFrame")
        row1.pack(fill="x", pady=(0, 15))

        ttk.Label(row1, text="Excel 路径:").pack(side="left")

        self.entry_path = ttk.Entry(row1, font=("Microsoft YaHei UI", 12))
        self.entry_path.insert(0, r"D:\抖音举报.xlsx")
        self.entry_path.pack(side="left", fill="x", expand=True, padx=15, ipady=3)

        if DND_FILES:
            self.entry_path.drop_target_register(DND_FILES)
            self.entry_path.dnd_bind('<<Drop>>', self.drop_file)

        self.last_path_val = self.entry_path.get()
        self.entry_path.bind("<FocusOut>",
                             lambda e: self.check_entry_change("Excel路径", self.entry_path, "last_path_val"))

        ttk.Button(row1, text="📂 浏览文件", width=10, command=self.browse_file).pack(side="right")

        row2 = ttk.Frame(card, style="Card.TFrame")
        row2.pack(fill="x")
        ttk.Button(row2, text="🖥️ 打开屏幕设置", command=self.open_display_settings).pack(side="left")
        ttk.Label(row2, text="⚠️ 必须设置：缩放 100% | 分辨率 1920x1080", foreground="#FF4D4F").pack(side="left",
                                                                                                    padx=15)

    def create_card_mode(self, parent):
        card = ttk.LabelFrame(parent, text=" 🎯 模式选择 ", style="Card.TLabelframe", padding=20)
        card.pack(fill="x", pady=(0, 15))

        self.mode_var = tk.StringVar(value="mixed")
        self.mode_var.trace_add("write", self.on_mode_change)

        grid_frame = ttk.Frame(card, style="Card.TFrame")
        grid_frame.pack(fill="x")
        PAD_Y_ROW = 8

        ttk.Radiobutton(grid_frame, text="仅视频举报", variable=self.mode_var, value="video",
                        style="Big.TRadiobutton").grid(row=0, column=0, sticky="w", pady=PAD_Y_ROW)
        f_video = ttk.Frame(grid_frame, style="Card.TFrame")
        f_video.grid(row=0, column=1, sticky="e", pady=PAD_Y_ROW)
        ttk.Label(f_video, text="Sheet名:", foreground="#999").pack(side="left", padx=(0, 5))
        self.entry_video_sheet = ttk.Entry(f_video, width=20, font=("Microsoft YaHei UI", 11))
        self.entry_video_sheet.insert(0, "举报指定视频")
        self.entry_video_sheet.pack(side="left", ipady=2)
        self.last_v_sheet = self.entry_video_sheet.get()
        self.entry_video_sheet.bind("<FocusOut>", lambda e: self.check_entry_change("视频Sheet", self.entry_video_sheet,
                                                                                    "last_v_sheet"))

        ttk.Radiobutton(grid_frame, text="仅用户举报", variable=self.mode_var, value="user",
                        style="Big.TRadiobutton").grid(row=1, column=0, sticky="w", pady=PAD_Y_ROW)
        f_user = ttk.Frame(grid_frame, style="Card.TFrame")
        f_user.grid(row=1, column=1, sticky="e", pady=PAD_Y_ROW)
        ttk.Label(f_user, text="Sheet名:", foreground="#999").pack(side="left", padx=(0, 5))
        self.entry_user_sheet = ttk.Entry(f_user, width=20, font=("Microsoft YaHei UI", 11))
        self.entry_user_sheet.insert(0, "举报指定用户")
        self.entry_user_sheet.pack(side="left", ipady=2)
        self.last_u_sheet = self.entry_user_sheet.get()
        self.entry_user_sheet.bind("<FocusOut>", lambda e: self.check_entry_change("用户Sheet", self.entry_user_sheet,
                                                                                   "last_u_sheet"))

        ttk.Radiobutton(grid_frame, text="混合模式 (推荐)", variable=self.mode_var, value="mixed",
                        style="Big.TRadiobutton").grid(row=2, column=0, sticky="w", pady=PAD_Y_ROW)

        grid_frame.columnconfigure(0, weight=1)
        grid_frame.columnconfigure(1, weight=0)

        ttk.Separator(card, orient="horizontal").pack(fill="x", pady=15)

        # === 【修改重点】改为使用自定义正方形图标的 tk.Checkbutton ===
        self.fast_mode_var = tk.BooleanVar(value=False)

        # 使用标准 tk.Checkbutton 来完全控制外观 (indicatoron=False 移除系统默认样式)
        self.chk_fast = tk.Checkbutton(card,
                                       text=" ⚡ 开启极速模式 ",
                                       variable=self.fast_mode_var,
                                       command=self.on_fast_change,
                                       bg="#FFFFFF",  # 背景白
                                       activebackground="#FFFFFF",
                                       font=("Microsoft YaHei UI", 12, "bold"),
                                       fg="#1890FF",  # 字体蓝
                                       selectcolor="#FFFFFF",  # 选中时背景
                                       bd=0,  # 无边框
                                       image=self.icon_chk_off,  # 未选中图片
                                       selectimage=self.icon_chk_on,  # 选中图片
                                       compound="left",  # 图片在文字左边
                                       indicatoron=False)  # 关键：关闭系统自带的打钩/打叉图标

        self.chk_fast.pack(anchor="w", padx=2)

    def create_card_control(self, parent):
        card = ttk.LabelFrame(parent, text=" 🕹️ 任务控制 ", style="Card.TLabelframe", padding=20)
        card.pack(fill="x")

        btn_box = ttk.Frame(card, style="Card.TFrame")
        btn_box.pack(fill="x")

        btn_box.columnconfigure(0, weight=1)
        btn_box.columnconfigure(1, weight=1)
        btn_box.columnconfigure(2, weight=1)

        self.btn_start = ttk.Button(btn_box, text="🚀 开始任务", style="Start.TButton", command=self.start_thread)
        self.btn_start.grid(row=0, column=0, padx=(0, 10), sticky="ew", ipady=10)

        self.btn_pause = ttk.Button(btn_box, text="⏸ 暂停", style="Pause.TButton", command=self.toggle_pause,
                                    state="disabled")
        self.btn_pause.grid(row=0, column=1, padx=10, sticky="ew", ipady=10)

        self.btn_stop = ttk.Button(btn_box, text="🛑 停止", style="Stop.TButton", command=self.stop_task,
                                   state="disabled")
        self.btn_stop.grid(row=0, column=2, padx=(10, 0), sticky="ew", ipady=10)

    def create_card_images(self, parent):
        card1 = ttk.LabelFrame(parent, text="   抖音TikTok", style="Card.TLabelframe", padding=10)
        card1.pack(fill="x", pady=(0, 15))

        self.mascot_fixed_frame = tk.Frame(card1, width=260, height=300, bg="#FFFFFF")
        self.mascot_fixed_frame.pack_propagate(False)
        self.mascot_fixed_frame.pack()

        self.mascot_label = ttk.Label(self.mascot_fixed_frame, background="#FFFFFF")
        self.mascot_label.place(relx=0.5, rely=0.5, anchor="center")

        card2 = ttk.LabelFrame(parent, text="    看板娘监工" , style="Card.TLabelframe", padding=10)
        card2.pack(fill="x")

        self.sticker_fixed_frame = tk.Frame(card2, width=260, height=160, bg="#FFFFFF")
        self.sticker_fixed_frame.pack_propagate(False)
        self.sticker_fixed_frame.pack()

        self.sticker_label = ttk.Label(self.sticker_fixed_frame, background="#FFFFFF")
        self.sticker_label.place(relx=0.5, rely=0.5, anchor="center")

        self.load_random_mascot(silent=True)
        self.load_random_sticker(silent=True)

    def setup_notice_ui(self):
        main_layout = ttk.Frame(self.frame_notice, style="Main.TFrame")
        main_layout.pack(fill="both", expand=True, padx=40, pady=40)

        lbl_title = ttk.Label(main_layout, text="📌 官方公告板", font=("Microsoft YaHei UI", 24, "bold"),
                              background="#F0F2F5")
        lbl_title.pack(anchor="w", pady=(0, 20))

        card = ttk.Frame(main_layout, style="Card.TFrame")
        card.pack(fill="both", expand=True)

        txt_notice = tk.Text(card, font=("Microsoft YaHei UI", 12), bg="#FFFFFF",
                             borderwidth=0, highlightthickness=0, padx=30, pady=30)
        txt_notice.insert("end", ANNOUNCEMENT_TEXT)
        txt_notice.config(state="disabled")
        txt_notice.pack(fill="both", expand=True)

    # ==========================================
    # 核心功能逻辑
    # ==========================================
    def open_display_settings(self):
        try:
            os.system("start ms-settings:display")
            print("[系统] 已打开系统显示设置。请确保：1. 缩放 100%  2. 分辨率 1920x1080")
        except:
            print("[错误] 无法打开设置")

    def check_entry_change(self, name, entry_widget, attr_name):
        current_val = entry_widget.get()
        old_val = getattr(self, attr_name)
        if current_val != old_val:
            print(f"[配置变更] {name} 已更新为: {current_val}")
            setattr(self, attr_name, current_val)

    def on_mode_change(self, *args):
        m = self.mode_var.get()
        map_text = {"video": "仅视频举报", "user": "仅用户举报", "mixed": "混合模式"}
        print(f"[模式切换] 当前选择: {map_text.get(m, m)}")

    def on_fast_change(self):
        state = "开启" if self.fast_mode_var.get() else "关闭"
        print(f"[极速模式] 已{state}")

    def show_work_area(self):
        self.frame_notice.pack_forget()
        self.frame_work.pack(fill="both", expand=True)
        self.btn_nav_work.configure(style="NavSelected.TButton")
        self.btn_nav_notice.configure(style="NavNormal.TButton")

    def show_notice_area(self):
        self.frame_work.pack_forget()
        self.frame_notice.pack(fill="both", expand=True)
        self.btn_nav_notice.configure(style="NavSelected.TButton")
        self.btn_nav_work.configure(style="NavNormal.TButton")

    def smart_sleep(self, seconds):
        if self.stop_event.is_set(): raise TaskStoppedError("任务已停止")
        while not self.pause_event.is_set():
            if self.stop_event.is_set(): raise TaskStoppedError("任务已停止")
            self.root.update()
            self.original_sleep(0.1)

        final_time = seconds
        if self.fast_mode_var.get():
            if seconds > 15:
                final_time = seconds - 10
            elif 2 < seconds < 3:
                final_time = seconds - 1
            elif 3 < seconds < 4:
                final_time = seconds - 2
            elif 4 < seconds < 5:
                final_time = seconds - 3
            if final_time != seconds: print(f"⚡ [极速] 优化等待: {seconds}s -> {final_time}s")

        rest = final_time
        while rest > 0:
            if self.stop_event.is_set(): raise TaskStoppedError("任务已停止")
            step = min(rest, 0.5)
            self.original_sleep(step)
            rest -= step

    def start_thread(self):
        if self.is_running: return
        self.stop_event.clear()
        self.pause_event.set()
        self.is_running = True
        self.btn_start.config(state="disabled", text="运行中...")
        self.btn_pause.config(state="normal", text="⏸ 暂停")
        self.btn_stop.config(state="normal")

        try:
            video_report.time.sleep = self.smart_sleep
            user_report.time.sleep = self.smart_sleep
        except NameError:
            pass

        t = threading.Thread(target=self.run_task)
        t.daemon = True
        t.start()

    def stop_task(self):
        if self.is_running and messagebox.askyesno("确认", "强制停止任务？"):
            self.stop_event.set()
            self.pause_event.set()
            print("\n🛑 正在停止...")

    def toggle_pause(self):
        if not self.is_running: return
        if self.pause_event.is_set():
            self.pause_event.clear()
            self.btn_pause.config(text="▶ 继续")
            print("\n⏸ 任务已暂停")
        else:
            self.pause_event.set()
            self.btn_pause.config(text="⏸ 暂停")
            print("\n▶ 任务继续")

    def run_task(self):
        path = self.entry_path.get().strip().replace('"', '').replace("'", "")
        mode = self.mode_var.get()
        v_sheet = self.entry_video_sheet.get().strip()
        u_sheet = self.entry_user_sheet.get().strip()

        print("\n" + "=" * 40)
        print(f"🚀 任务启动 | 模式: {mode}")
        print("=" * 40)

        try:
            try:
                import video_report
                import user_report
            except ImportError:
                print("❌ 未找到业务模块 (video_report.py 或 user_report.py)")
                print("   仅作为UI演示模式运行")
                time.sleep(2)
                return

            if not os.path.exists(path):
                messagebox.showerror("错误", f"文件不存在:\n{path}")
                return

            if mode == 'video':
                video_report.main(excel_path=path, sheet_name=v_sheet, auto_exit=True)
            elif mode == 'user':
                user_report.main(excel_path=path, sheet_name=u_sheet, auto_exit=True)
            elif mode == 'mixed':
                print(">> 阶段一：视频举报")
                video_report.main(excel_path=path, sheet_name=v_sheet, auto_exit=True)
                print("\n" + "-" * 20)
                print(">> 阶段二：用户举报")
                time.sleep(2)
                user_report.main(excel_path=path, sheet_name=u_sheet, auto_exit=True)

            print("\n" + "=" * 30)
            print("✅ 任务完成！")
            messagebox.showinfo("完成", "所有任务已处理完毕！")

        except TaskStoppedError:
            print("\n🛑 任务已被强制停止。")
        except Exception as e:
            print(f"\n❌ 发生错误: {e}")
            print(traceback.format_exc())
        finally:
            self.is_running = False
            self.root.after(0, self.reset_ui)

    def reset_ui(self):
        self.btn_start.config(state="normal", text="🚀 开始任务")
        self.btn_pause.config(state="disabled", text="⏸ 暂停")
        self.btn_stop.config(state="disabled")
        self.pause_event.set()

    def resize_image_to_fit(self, pil_image, max_w, max_h):
        w, h = pil_image.size
        ratio = min(max_w / w, max_h / h)
        return pil_image.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)

    def load_random_mascot(self, silent=False):
        try:
            img_list = ["dy_logo.jpg"]
            chosen = random.choice(img_list)
            path = get_resource_path(os.path.join("images", chosen))
            if os.path.exists(path):
                img = self.resize_image_to_fit(Image.open(path), 250, 280)
                self.photo_mascot = ImageTk.PhotoImage(img)
                self.mascot_label.config(image=self.photo_mascot)
                self.mascot_label.bind("<Button-1>", lambda e: self.load_random_mascot())
        except:
            pass

    def load_random_sticker(self, silent=False):
        try:
            img_list = ["m1.jpg", "m2.jpg", "m3.jpg", "m4.jpg", "m5.jpg", "m6.jpg", "m7.jpg"]
            chosen = random.choice(img_list)
            path = get_resource_path(os.path.join("images", chosen))
            if os.path.exists(path):
                img = self.resize_image_to_fit(Image.open(path), 250, 140)
                self.photo_sticker = ImageTk.PhotoImage(img)
                self.sticker_label.config(image=self.photo_sticker)
                self.sticker_label.bind("<Button-1>", lambda e: self.load_random_sticker())
        except:
            pass

    def drop_file(self, event):
        path = event.data.strip('{}')
        self.entry_path.delete(0, tk.END)
        self.entry_path.insert(0, path)
        print(f"[文件加载] 拖拽加载: {path}")
        self.last_path_val = path

    def browse_file(self):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if f:
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, f)
            print(f"[文件加载] 选择文件: {f}")
            self.last_path_val = f


if __name__ == "__main__":
    try:
        try:
            from tkinterdnd2 import TkinterDnD
        except ImportError:
            raise ImportError("缺少库: tkinterdnd2 (pip install tkinterdnd2)")

        root = TkinterDnD.Tk()
        try:
            from ctypes import windll

            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass

        app = ReportApp(root)
        root.mainloop()

    except Exception as e:
        import tkinter.messagebox

        if 'root' not in locals() or not root:
            temp_root = tk.Tk()
            temp_root.withdraw()

        err_msg = f"程序启动失败！\n\n{str(e)}\n\n{traceback.format_exc()}"
        tkinter.messagebox.showerror("发生错误", err_msg)