"""图形界面（只负责界面与用户输入，不含任何业务逻辑）。

业务执行统一交给 pyautogui.core 层，通过 TaskController 控制暂停/停止。
"""
from __future__ import annotations

import os
import random
import sys
import threading
import traceback

# 兼容“双击运行”与“python -m pyautogui.ui_main”两种方式
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from PIL import Image, ImageDraw, ImageTk

from pyautogui import config
from pyautogui.core import user_report, video_report
from pyautogui.core.controller import TaskController, TaskStopped

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None

APP_TITLE = "抖音自动化举报工具箱 v1.1"

ANNOUNCEMENT_TEXT = """📢 系统公告 (v1.1)：

1. 【开源声明】本软件为免费开源项目，仅供个人学习与技术研究，严禁用于任何商业用途。

2. 【v1.1 重构说明】
   • UI 与业务逻辑已解耦：界面只负责展示，业务在 pyautogui/core 中独立运行；
   • 所有可调参数（匹配阈值、重试次数、等待时间等）集中到 pyautogui/config.py；
   • 暂停/停止/极速模式改为统一 TaskController，不再替换 time.sleep；
   • 更现代的 v2 重构版（Playwright + PySide6）见 GitHub 仓库 tiktok-report-tool-v2。

3. 【使用须知】
   • 提前下载 Google Chrome 浏览器，登录个人抖音账号网页版；
   • 在模板 excel 里按分类依次录入举报的账号（视频）、举报类型、举报理由；
   • 支持拖拽文件到 Excel 路径输入框；
   • 屏幕分辨率必须为 1920x1080、缩放 100%（图像识别依赖固定屏幕环境）；
   • 软件运行期间请勿移动鼠标和触碰键盘；
   • 若电脑性能不足，请慎重选择极速模式。

⚠️ 免责声明：请合理合法使用本工具，用户使用本工具产生的一切后果由用户自行承担。
"""


class LogWriter:
    """把业务日志安全地送到界面控件（跨线程通过 after 调度）。"""

    def __init__(self, text_widget: tk.Text):
        self.text_widget = text_widget

    def write(self, string: str) -> None:
        self.text_widget.after(0, self._append, string)

    def _append(self, string: str) -> None:
        try:
            self.text_widget.config(state="normal")
            self.text_widget.insert("end", string)
            self.text_widget.see("end")
            self.text_widget.config(state="disabled")
        except Exception:
            pass


class ReportApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1100x980")
        self.root.resizable(False, True)

        self.is_running = False
        self.controller = TaskController()
        self.last_path_val = ""
        self.last_v_sheet = ""
        self.last_u_sheet = ""

        self.setup_custom_assets()
        self.setup_styles()
        self.setup_nav()
        self.setup_work_area()
        self.setup_notice_area()
        self.show_work_area()

        self.log_writer = LogWriter(self.log_text)
        self.log_writer.write(">>> 系统初始化完成，等待指令...\n")

    # ==========================================
    # 资源与样式（外观无关业务）
    # ==========================================
    def setup_custom_assets(self) -> None:
        img_off = Image.new("RGBA", (24, 24), (255, 255, 255, 0))
        ImageDraw.Draw(img_off).rectangle([2, 2, 21, 21], outline="#999999", width=2)
        img_on = Image.new("RGBA", (24, 24), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img_on)
        draw.rectangle([2, 2, 21, 21], outline="#1890FF", width=2)
        draw.rectangle([6, 6, 17, 17], fill="#1890FF")
        self.icon_chk_off = ImageTk.PhotoImage(img_off)
        self.icon_chk_on = ImageTk.PhotoImage(img_on)

    def setup_styles(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        bg, card, primary = "#F0F2F5", "#FFFFFF", "#1890FF"
        font = ("Microsoft YaHei UI", 12)
        bold = ("Microsoft YaHei UI", 12, "bold")
        title = ("Microsoft YaHei UI", 14, "bold")
        large = ("Microsoft YaHei UI", 16, "bold")
        style.configure("Main.TFrame", background=bg)
        style.configure("Card.TFrame", background=card, relief="flat", borderwidth=0)
        style.configure("Card.TLabelframe", background=card, relief="flat", borderwidth=1)
        style.configure("Card.TLabelframe.Label", font=title, foreground=primary, background=card)
        style.configure("Nav.TFrame", background="#FFFFFF")
        style.configure("NavSelected.TButton", font=title, background="#E6F7FF", foreground=primary, borderwidth=0)
        style.configure("NavNormal.TButton", font=("Microsoft YaHei UI", 14), background="#FFFFFF",
                        foreground="#666666", borderwidth=0)
        style.configure("TLabel", background=card, foreground="#333333", font=font)
        style.configure("Gray.TLabel", background=bg, foreground="#333333", font=font)
        style.configure("TButton", font=font, borderwidth=1, background="#FFFFFF")
        style.map("TButton", background=[("active", "#E6F7FF")], foreground=[("active", primary)])
        style.configure("Big.TRadiobutton", font=bold, background=card, foreground="#333333", indicatorwidth=18)
        style.configure("Start.TButton", font=large, background="#F6FFED", foreground="#389E0D")
        style.map("Start.TButton", background=[("active", "#52C41A")], foreground=[("active", "white")])
        style.configure("Pause.TButton", font=large, background="#FFF7E6", foreground="#D46B08")
        style.map("Pause.TButton", background=[("active", "#FA8C16")], foreground=[("active", "white")])
        style.configure("Stop.TButton", font=large, background="#FFF1F0", foreground="#CF1322")
        style.map("Stop.TButton", background=[("active", "#FF4D4F")], foreground=[("active", "white")])

    # ==========================================
    # 导航
    # ==========================================
    def setup_nav(self) -> None:
        nav = ttk.Frame(self.root, style="Nav.TFrame")
        nav.pack(side="top", fill="x")
        self.btn_nav_work = ttk.Button(nav, text="🖥️ 工作控制台", command=self.show_work_area, style="NavSelected.TButton")
        self.btn_nav_work.pack(side="left", fill="y", ipadx=20, ipady=15)
        self.btn_nav_notice = ttk.Button(nav, text="📜 系统公告", command=self.show_notice_area, style="NavNormal.TButton")
        self.btn_nav_notice.pack(side="left", fill="y", ipadx=20, ipady=15)

        container = ttk.Frame(self.root, style="Main.TFrame")
        container.pack(side="bottom", fill="both", expand=True)
        self.frame_work = ttk.Frame(container, style="Main.TFrame")
        self.frame_notice = ttk.Frame(container, style="Main.TFrame")
        self.setup_work_ui()
        self.setup_notice_ui()

    def show_work_area(self) -> None:
        self.frame_notice.pack_forget()
        self.frame_work.pack(fill="both", expand=True)
        self.btn_nav_work.configure(style="NavSelected.TButton")
        self.btn_nav_notice.configure(style="NavNormal.TButton")

    def show_notice_area(self) -> None:
        self.frame_work.pack_forget()
        self.frame_notice.pack(fill="both", expand=True)
        self.btn_nav_notice.configure(style="NavSelected.TButton")
        self.btn_nav_work.configure(style="NavNormal.TButton")

    # ==========================================
    # 工作区
    # ==========================================
    def setup_work_ui(self) -> None:
        layout = ttk.Frame(self.frame_work, style="Main.TFrame")
        layout.pack(fill="both", expand=True, padx=20, pady=20)

        top = ttk.Frame(layout, style="Main.TFrame")
        top.pack(side="top", fill="x", anchor="n")

        left = ttk.Frame(top, style="Main.TFrame")
        left.pack(side="left", fill="both", expand=True, padx=(0, 20))
        self.create_card_config(left)
        self.create_card_mode(left)
        self.create_card_control(left)

        right = ttk.Frame(top, style="Main.TFrame")
        right.pack(side="right", fill="y")
        self.create_card_images(right)

        log_card = ttk.LabelFrame(layout, text=" 📟 实时运行日志 ", style="Card.TLabelframe", padding=15)
        log_card.pack(side="bottom", fill="both", expand=True, pady=(20, 0))
        self.log_text = scrolledtext.ScrolledText(log_card, state="disabled", bg="#1E1E1E", fg="#52C41A",
                                                  font=("Consolas", 11), borderwidth=0, insertbackground="white")
        self.log_text.pack(fill="both", expand=True)

    def create_card_config(self, parent) -> None:
        card = ttk.LabelFrame(parent, text=" 🛠️ 基础配置 ", style="Card.TLabelframe", padding=20)
        card.pack(fill="x", pady=(0, 15))
        row1 = ttk.Frame(card, style="Card.TFrame")
        row1.pack(fill="x", pady=(0, 15))
        ttk.Label(row1, text="Excel 路径:").pack(side="left")
        self.entry_path = ttk.Entry(row1, font=("Microsoft YaHei UI", 12))
        self.entry_path.insert(0, config.DEFAULT_EXCEL_PATH)
        self.entry_path.pack(side="left", fill="x", expand=True, padx=15, ipady=3)
        if DND_FILES:
            self.entry_path.drop_target_register(DND_FILES)
            self.entry_path.dnd_bind("<<Drop>>", self.drop_file)
        self.last_path_val = self.entry_path.get()
        self.entry_path.bind("<FocusOut>", lambda e: self.check_entry_change("Excel路径", self.entry_path, "last_path_val"))
        ttk.Button(row1, text="📂 浏览文件", width=10, command=self.browse_file).pack(side="right")

        row2 = ttk.Frame(card, style="Card.TFrame")
        row2.pack(fill="x")
        ttk.Button(row2, text="🖥️ 打开屏幕设置", command=self.open_display_settings).pack(side="left")
        ttk.Label(row2, text=f"⚠️ 必须设置：缩放 {config.SCREEN_SCALE} | 分辨率 {config.SCREEN_WIDTH}x{config.SCREEN_HEIGHT}",
                  foreground="#FF4D4F").pack(side="left", padx=15)

    def create_card_mode(self, parent) -> None:
        card = ttk.LabelFrame(parent, text=" 🎯 模式选择 ", style="Card.TLabelframe", padding=20)
        card.pack(fill="x", pady=(0, 15))
        self.mode_var = tk.StringVar(value="mixed")
        grid = ttk.Frame(card, style="Card.TFrame")
        grid.pack(fill="x")

        ttk.Radiobutton(grid, text="仅视频举报", variable=self.mode_var, value="video",
                        style="Big.TRadiobutton").grid(row=0, column=0, sticky="w", pady=8)
        f_video = ttk.Frame(grid, style="Card.TFrame")
        f_video.grid(row=0, column=1, sticky="e", pady=8)
        ttk.Label(f_video, text="Sheet名:", foreground="#999").pack(side="left", padx=(0, 5))
        self.entry_video_sheet = ttk.Entry(f_video, width=20, font=("Microsoft YaHei UI", 11))
        self.entry_video_sheet.insert(0, config.VIDEO_SHEET_NAME)
        self.entry_video_sheet.pack(side="left", ipady=2)
        self.last_v_sheet = self.entry_video_sheet.get()
        self.entry_video_sheet.bind("<FocusOut>",
                                    lambda e: self.check_entry_change("视频Sheet", self.entry_video_sheet, "last_v_sheet"))

        ttk.Radiobutton(grid, text="仅用户举报", variable=self.mode_var, value="user",
                        style="Big.TRadiobutton").grid(row=1, column=0, sticky="w", pady=8)
        f_user = ttk.Frame(grid, style="Card.TFrame")
        f_user.grid(row=1, column=1, sticky="e", pady=8)
        ttk.Label(f_user, text="Sheet名:", foreground="#999").pack(side="left", padx=(0, 5))
        self.entry_user_sheet = ttk.Entry(f_user, width=20, font=("Microsoft YaHei UI", 11))
        self.entry_user_sheet.insert(0, config.USER_SHEET_NAME)
        self.entry_user_sheet.pack(side="left", ipady=2)
        self.last_u_sheet = self.entry_user_sheet.get()
        self.entry_user_sheet.bind("<FocusOut>",
                                    lambda e: self.check_entry_change("用户Sheet", self.entry_user_sheet, "last_u_sheet"))

        ttk.Radiobutton(grid, text="混合模式 (推荐)", variable=self.mode_var, value="mixed",
                        style="Big.TRadiobutton").grid(row=2, column=0, sticky="w", pady=8)
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=0)

        ttk.Separator(card, orient="horizontal").pack(fill="x", pady=15)
        self.fast_mode_var = tk.BooleanVar(value=False)
        self.chk_fast = tk.Checkbutton(card, text=" ⚡ 开启极速模式 ", variable=self.fast_mode_var, command=self.on_fast_change,
                                       bg="#FFFFFF", activebackground="#FFFFFF", font=("Microsoft YaHei UI", 12, "bold"),
                                       fg="#1890FF", selectcolor="#FFFFFF", bd=0, image=self.icon_chk_off,
                                       selectimage=self.icon_chk_on, compound="left", indicatoron=False)
        self.chk_fast.pack(anchor="w", padx=2)

    def create_card_control(self, parent) -> None:
        card = ttk.LabelFrame(parent, text=" 🕹️ 任务控制 ", style="Card.TLabelframe", padding=20)
        card.pack(fill="x")
        box = ttk.Frame(card, style="Card.TFrame")
        box.pack(fill="x")
        for col in range(3):
            box.columnconfigure(col, weight=1)
        self.btn_start = ttk.Button(box, text="🚀 开始任务", style="Start.TButton", command=self.start_task)
        self.btn_start.grid(row=0, column=0, padx=(0, 10), sticky="ew", ipady=10)
        self.btn_pause = ttk.Button(box, text="⏸ 暂停", style="Pause.TButton", command=self.toggle_pause, state="disabled")
        self.btn_pause.grid(row=0, column=1, padx=10, sticky="ew", ipady=10)
        self.btn_stop = ttk.Button(box, text="🛑 停止", style="Stop.TButton", command=self.stop_task, state="disabled")
        self.btn_stop.grid(row=0, column=2, padx=(10, 0), sticky="ew", ipady=10)

    def create_card_images(self, parent) -> None:
        card1 = ttk.LabelFrame(parent, text="   抖音TikTok", style="Card.TLabelframe", padding=10)
        card1.pack(fill="x", pady=(0, 15))
        frame1 = tk.Frame(card1, width=260, height=300, bg="#FFFFFF")
        frame1.pack_propagate(False)
        frame1.pack()
        self.mascot_label = ttk.Label(frame1, background="#FFFFFF")
        self.mascot_label.place(relx=0.5, rely=0.5, anchor="center")

        card2 = ttk.LabelFrame(parent, text="    看板娘监工 ", style="Card.TLabelframe", padding=10)
        card2.pack(fill="x")
        frame2 = tk.Frame(card2, width=260, height=160, bg="#FFFFFF")
        frame2.pack_propagate(False)
        frame2.pack()
        self.sticker_label = ttk.Label(frame2, background="#FFFFFF")
        self.sticker_label.place(relx=0.5, rely=0.5, anchor="center")
        self.load_random_mascot()
        self.load_random_sticker()

    def setup_notice_ui(self) -> None:
        layout = ttk.Frame(self.frame_notice, style="Main.TFrame")
        layout.pack(fill="both", expand=True, padx=40, pady=40)
        ttk.Label(layout, text="📌 官方公告板", font=("Microsoft YaHei UI", 24, "bold"),
                  background="#F0F2F5").pack(anchor="w", pady=(0, 20))
        text = tk.Text(layout, font=("Microsoft YaHei UI", 12), bg="#FFFFFF", borderwidth=0,
                       highlightthickness=0, padx=30, pady=30)
        text.insert("end", ANNOUNCEMENT_TEXT)
        text.config(state="disabled")
        text.pack(fill="both", expand=True)

    # ==========================================
    # 界面小工具
    # ==========================================
    def resize_image_to_fit(self, image, max_w: int, max_h: int):
        w, h = image.size
        ratio = min(max_w / w, max_h / h)
        return image.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)

    def load_random_mascot(self) -> None:
        try:
            path = os.path.join(config.IMAGE_DIR, "dy_logo.jpg")
            if os.path.exists(path):
                image = self.resize_image_to_fit(Image.open(path), 250, 280)
                self.photo_mascot = ImageTk.PhotoImage(image)
                self.mascot_label.config(image=self.photo_mascot)
                self.mascot_label.bind("<Button-1>", lambda e: self.load_random_mascot())
        except Exception:
            pass

    def load_random_sticker(self) -> None:
        try:
            names = [f"m{i}.jpg" for i in range(1, 8)]
            path = os.path.join(config.IMAGE_DIR, random.choice(names))
            if os.path.exists(path):
                image = self.resize_image_to_fit(Image.open(path), 250, 140)
                self.photo_sticker = ImageTk.PhotoImage(image)
                self.sticker_label.config(image=self.photo_sticker)
                self.sticker_label.bind("<Button-1>", lambda e: self.load_random_sticker())
        except Exception:
            pass

    def open_display_settings(self) -> None:
        os.system("start ms-settings:display")
        self.log_writer.write(f"[系统] 已打开系统显示设置。请确保：1. 缩放 {config.SCREEN_SCALE}  "
                              f"2. 分辨率 {config.SCREEN_WIDTH}x{config.SCREEN_HEIGHT}\n")

    def check_entry_change(self, name: str, widget, attr_name: str) -> None:
        current = widget.get()
        if current != getattr(self, attr_name):
            self.log_writer.write(f"[配置变更] {name} 已更新为: {current}\n")
            setattr(self, attr_name, current)

    def drop_file(self, event) -> None:
        path = event.data.strip("{}")
        self.entry_path.delete(0, tk.END)
        self.entry_path.insert(0, path)
        self.last_path_val = path
        self.log_writer.write(f"[文件加载] 拖拽加载: {path}\n")

    def browse_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, path)
            self.last_path_val = path
            self.log_writer.write(f"[文件加载] 选择文件: {path}\n")

    def on_fast_change(self) -> None:
        state = "开启" if self.fast_mode_var.get() else "关闭"
        self.log_writer.write(f"[极速模式] 已{state}\n")

    # ==========================================
    # 任务控制（只做控制，不写业务）
    # ==========================================
    def start_task(self) -> None:
        if self.is_running:
            return
        self.controller = TaskController(fast_mode=self.fast_mode_var.get())
        self.is_running = True
        self.btn_start.config(state="disabled", text="运行中...")
        self.btn_pause.config(state="normal", text="⏸ 暂停")
        self.btn_stop.config(state="normal")
        threading.Thread(target=self.run_task, daemon=True).start()

    def run_task(self) -> None:
        excel_path = self.entry_path.get().strip().strip('"\'')
        mode = self.mode_var.get()
        log = self.log_writer.write
        try:
            if not os.path.exists(excel_path):
                messagebox.showerror("错误", f"文件不存在:\n{excel_path}")
                return
            log("\n" + "=" * 40 + f"\n🚀 任务启动 | 模式: {mode}\n" + "=" * 40 + "\n")
            if mode in ("video", "mixed"):
                video_report.run_video_tasks(self.controller, log=log, excel_path=excel_path,
                                             sheet_name=self.entry_video_sheet.get().strip())
            if mode in ("user", "mixed"):
                user_report.run_user_tasks(self.controller, log=log, excel_path=excel_path,
                                           sheet_name=self.entry_user_sheet.get().strip())
            log("✅ 所有任务处理完毕！\n")
        except TaskStopped:
            log("🛑 任务已被强制停止。\n")
        except Exception as exc:
            log(f"❌ 发生错误: {exc}\n{traceback.format_exc()}\n")
        finally:
            self.is_running = False
            self.root.after(0, self.reset_ui)

    def toggle_pause(self) -> None:
        if not self.is_running:
            return
        if self.controller.paused:
            self.controller.resume()
            self.btn_pause.config(text="⏸ 暂停")
            self.log_writer.write("▶ 任务继续\n")
        else:
            self.controller.pause()
            self.btn_pause.config(text="▶ 继续")
            self.log_writer.write("⏸ 任务已暂停\n")

    def stop_task(self) -> None:
        if self.is_running and messagebox.askyesno("确认", "强制停止任务？"):
            self.controller.request_stop()
            self.log_writer.write("🛑 正在停止...\n")

    def reset_ui(self) -> None:
        self.btn_start.config(state="normal", text="🚀 开始任务")
        self.btn_pause.config(state="disabled", text="⏸ 暂停")
        self.btn_stop.config(state="disabled")


def main() -> None:
    try:
        try:
            from tkinterdnd2 import TkinterDnD
        except ImportError:
            raise ImportError("缺少库: tkinterdnd2 (pip install tkinterdnd2)")
        root = TkinterDnD.Tk()
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        ReportApp(root)
        root.mainloop()
    except Exception as exc:
        root = tk.Tk() if "root" not in locals() or not root else root
        root.withdraw()
        messagebox.showerror("发生错误", f"程序启动失败！\n\n{exc}\n\n{traceback.format_exc()}")


if __name__ == "__main__":
    main()

