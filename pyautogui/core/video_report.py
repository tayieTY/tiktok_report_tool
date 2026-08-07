"""视频举报业务流程（纯业务，不依赖任何 UI 代码）。"""
from __future__ import annotations

import os
import random
from typing import Callable, Optional

import pyautogui

from pyautogui import config
from pyautogui.core import automation
from pyautogui.core.controller import TaskController
from pyautogui.core.excel_reader import load_rows


def run_video_tasks(
    controller: TaskController,
    log: Callable[[str], None] = print,
    excel_path: Optional[str] = None,
    sheet_name: Optional[str] = None,
    launch_browser: bool = True,
) -> None:
    """读取 Excel 并逐行执行视频举报。"""
    excel_path = excel_path or config.DEFAULT_EXCEL_PATH
    sheet_name = sheet_name or config.VIDEO_SHEET_NAME
    rows = load_rows(excel_path, sheet_name, config.VIDEO_COLUMNS)
    log(f"📋 任务已就绪，工作表「{sheet_name}」，总计 {len(rows)} 条数据")

    if launch_browser:
        os.system("start chrome")
        controller.sleep(config.VIDEO_BROWSER_START_SLEEP)
        pyautogui.hotkey("alt", "space")
        pyautogui.hotkey("x")
        pyautogui.hotkey("alt")

    for row_index, values in rows:
        controller.check()
        process_video_row(row_index, values, controller, log)
    log(f"\n✅ 「{sheet_name}」任务处理完毕")


def process_video_row(row_index: int, data: list, controller: TaskController, log: Callable[[str], None]) -> None:
    search_q, reason_main, reason_sub, detail = [str(v or "").strip() for v in (data + [None, None, None, None])[:4]]
    log(f"\n🚀 处理第 {row_index} 行: {search_q}")

    # 1. 新开标签页
    pyautogui.hotkey("ctrl", "t")
    controller.sleep(random.uniform(*config.VIDEO_NAV))

    # 2. 搜索并等待结果
    if automation.universal_click("search", controller=controller, log=log):
        automation.input_text(search_q, controller=controller)
        pyautogui.press("enter")
        pyautogui.hotkey("ctrl", "0")
        log("  [🔍] 搜索结果加载中，等待页面刷新...")
        controller.sleep(random.uniform(*config.VIDEO_SEARCH_WAIT))

        # 3. 模拟真人浏览
        pyautogui.moveTo(*random.choice(config.VIDEO_BROWSE_COORDINATES), duration=0.5)
        controller.sleep(random.uniform(*config.VIDEO_SCROLL_WAIT))
        pyautogui.scroll(config.VIDEO_SCROLL_AMOUNT)
        controller.sleep(random.uniform(*config.VIDEO_SCROLL_WAIT))

        # 4. 打开举报
        if automation.universal_click("report", conf=config.CONFIDENCE_REPORT_BUTTON, controller=controller, log=log):
            pyautogui.press("enter")
            if automation.universal_click(reason_main, mode="left", controller=controller, log=log):
                if reason_sub:
                    automation.universal_click(reason_sub, mode="left", controller=controller, log=log)
                if detail:
                    if automation.universal_click("input", controller=controller, log=log):
                        automation.input_text(detail, controller=controller)
                automation.universal_click("confirm", controller=controller, log=log)
                controller.sleep(random.uniform(*config.DELAY_SUBMIT))

    # 5. 关闭标签页，准备下一行
    pyautogui.hotkey("ctrl", "w")
    controller.sleep(random.uniform(*config.VIDEO_CLEANUP))

