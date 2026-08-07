"""集中管理所有可调参数，消除散落在业务代码里的魔法数。"""
from __future__ import annotations

import os
import sys

# ---------------------------------------------------------------- 屏幕环境
# 旧版基于图像识别，依赖固定屏幕环境；打包后仍需保持此要求
SCREEN_WIDTH = 1920
SCREEN_HEIGHT = 1080
SCREEN_SCALE = "100%"

# ---------------------------------------------------------------- 文件与工作表
DEFAULT_EXCEL_PATH = r"D:\抖音举报.xlsx"
VIDEO_SHEET_NAME = "举报指定视频"
USER_SHEET_NAME = "举报指定用户"
EXCEL_START_ROW = 2  # 第一行为表头

# 各流程读取的 Excel 列（A/B/C/D）
VIDEO_COLUMNS = (1, 2, 3, 4)
USER_COLUMNS = (1, 2, 3)

# ---------------------------------------------------------------- 资源路径
_BASE = sys._MEIPASS if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
IMAGE_DIR = os.path.join(_BASE, "images")

# ---------------------------------------------------------------- 图像识别
CONFIDENCE_DEFAULT = 0.7    # 通用匹配阈值
CONFIDENCE_MENU = 0.8       # “更多”菜单
CONFIDENCE_CONFIRM = 0.9    # 确认按钮
CONFIDENCE_VIDEO = 0.9      # 视频流程默认
CONFIDENCE_REPORT_BUTTON = 0.99  # 视频“举报”按钮（高阈值防误点）

# 重试次数
RETRY_DEFAULT = 15
RETRY_FAST = 10             # 需要快速定位的场景（搜索框等）
RETRY_VIDEO = 20            # 视频流程

# 点击偏移（“left”模式用于复选/单选行）
CLICK_OFFSET_LEFT = 15
CHECKBOX_OFFSET_SECONDARY = 14  # 二级原因（文件名含 _）
CHECKBOX_OFFSET_PRIMARY = 10    # 一级原因
CLICK_JITTER = 3                 # 随机微调范围，模拟真人
VIDEO_CLICK_JITTER = 1
POST_CLICK_SLEEP = 1.0           # 视频流程点击后固定停顿

# ---------------------------------------------------------------- 随机等待范围（秒）
DELAY_CLICK = (0.2, 0.4)        # 移动鼠标
DELAY_RETRY = (0.3, 1.4)        # 未找到图片时重试间隔
DELAY_INPUT = (1.8, 2.5)        # 粘贴文本后
DELAY_LAUNCH = (2.5, 3.5)       # 浏览器启动后
DELAY_PAGE = (2.5, 3.8)         # 用户主页跳转后
DELAY_FANS = (2.2, 2.7)         # 粉丝定位后
DELAY_MENU = (1.8, 2.2)         # 更多菜单展开后
DELAY_DIALOG = (2.1, 2.6)       # 举报弹窗打开后
DELAY_REASON = (2.5, 2.9)       # 选择原因后
DELAY_SUBMIT = (2.5, 3.5)       # 提交后
DELAY_CLEANUP = (2.2, 4.0)      # 关闭标签页后

# 视频流程专用
VIDEO_NAV = (2.7, 3.9)          # 新开标签页后
VIDEO_SEARCH_WAIT = (15.0, 16.4)  # 搜索结果加载
VIDEO_SCROLL_WAIT = (2.5, 3.4)  # 模拟浏览滚动
VIDEO_INPUT = (2.3, 3.4)        # 视频描述粘贴后
VIDEO_CLEANUP = (2.3, 3.0)      # 关闭视频标签页后
VIDEO_BROWSE_COORDINATES = [(703, 415), (1000, 490), (800, 600)]  # 模拟真人浏览的点击点
VIDEO_SCROLL_AMOUNT = -250
VIDEO_BROWSER_START_SLEEP = 3.0

# ---------------------------------------------------------------- 浏览器
BROWSER_URL = "https://www.douyin.com/"

# ---------------------------------------------------------------- 举报原因 → 图片（用户流程）
USER_REPORT_REASONS = {
    "内容违规": "nr.png",
    "色情低俗": "nr_sqds.png",
    "不实信息": "nr_bsxx.png",
    "政治敏感": "nr_zzmg.png",
}

# ---------------------------------------------------------------- 图片映射（视频流程）
VIDEO_IMAGE_MAP = {
    "report": "report_button.png",
    "search": "search_box.png",
    "input": "click_put_in.png",
    "confirm": "confirm.png",
    # 一级原因
    "政治敏感": "zz.png",
    "色情低俗": "sq.png",
    "不实信息": "bs.png",
    # 二级原因
    "涉政不当言论": "zz_szbdyl.png",
    "涉政不实信息": "zz_szbsxx.png",
    "色情裸露内容": "sq_sqllnr.png",
    "未成年低俗": "sq_wcnds.png",
    "疑似招嫖": "sq_yszp.png",
    "刻意抹黑": "bs_kymh.png",
    "虚假摆拍演绎": "bs_xjbpyy.png",
    "疑似虚假时事": "bs_ysxjss.png",
    "疑似虚假知识": "bs_ysxjzs.png",
}

# 用户流程用到的界面图片
USER_IMAGE_SEARCH_BOX = "search_box_2.png"
USER_IMAGE_USER = "user.png"
USER_IMAGE_FANS = "fans.png"
USER_IMAGE_MORE = "more.png"
USER_IMAGE_REPORT_BUTTON = "report_button_2.png"
USER_IMAGE_MAIN_REASON = "nr.png"
USER_IMAGE_INPUT = "click_put_in_2.png"
USER_IMAGE_CONFIRM = "confirm_2.png"

