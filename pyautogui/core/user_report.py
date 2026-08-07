"""用户举报业务流程（纯业务，不依赖任何 UI 代码）。"""
from __future__ import annotations

import os
import random
from typing import Callable, Optional

import pyautogui

from pyautogui import config
from pyautogui.core import automation
from pyautogui.core.controller import TaskController
from pyautogui.core.excel_reader import load_rows


def run_user_tasks(
    controller: TaskController,
    log: Callable[[str], None] = print,
    excel_path: Optional[str] = None,
    sheet_name: Optional[str] = None,
    launch_browser: bool = True,
) -> None:
    """读取 Excel 并逐行执行用户举报。"""
    excel_path = excel_path or config.DEFAULT_EXCEL_PATH
    sheet_name = sheet_name or config.USER_SHEET_NAME
    rows = load_rows(excel_path, sheet_name, config.USER_COLUMNS)
    log(f"📋 任务已就绪，工作表「{sheet_name}」，总计 {len(rows)} 条数据")

    if launch_browser:
        os.system(f"start chrome {config.BROWSER_URL}")
        log("  [🔍] 正在启动浏览器并打开抖音...")
        controller.sleep(random.uniform(*config.DELAY_LAUNCH))
        pyautogui.hotkey("alt", "space")
        pyautogui.press("x")
        controller.sleep(random.uniform(2.0, 4.0))

    for row_index, values in rows:
        controller.check()
        process_user_row(row_index, values, controller, log)
    log(f"\n✅ 「{sheet_name}」任务处理完毕")


def process_user_row(row_index: int, data: list, controller: TaskController, log: Callable[[str], None]) -> None:
    search_text, report_type, detail = [str(v or "").strip() for v in (data + [None, None, None])[:3]]
    log(f"\n>>> 正在处理第 {row_index} 行: {search_text}")

    # 1. 点击搜索框并输入
    if not automation.find_and_click(config.USER_IMAGE_SEARCH_BOX, attempts=config.RETRY_FAST, controller=controller, log=log):
        log("  [X] 找不到搜索框，跳过本行")
        return

    automation.input_text(search_text, controller=controller)
    pyautogui.press("enter")
    pyautogui.hotkey("ctrl", "0")
    log("  [🔍] 搜索结果加载中，等待页面刷新...")
    controller.sleep(random.uniform(*config.DELAY_PAGE))

    # 2. 点击用户头像进入主页
    if automation.find_and_click(config.USER_IMAGE_USER, confidence=config.CONFIDENCE_DEFAULT, attempts=config.RETRY_DEFAULT, controller=controller, log=log):
        log("  [✓] 已点击用户头像")
        controller.sleep(random.uniform(*config.DELAY_PAGE))

        # 3. 点击“粉丝”辅助定位
        if automation.find_and_click(config.USER_IMAGE_FANS, confidence=config.CONFIDENCE_DEFAULT, attempts=config.RETRY_DEFAULT, controller=controller, log=log):
            log("  [✓] 已定位粉丝入口")
            controller.sleep(random.uniform(*config.DELAY_FANS))
        else:
            log("  [!] 未找到粉丝图标，尝试直接寻找菜单...")

        # 4. 更多 → 举报
        if automation.find_and_click(config.USER_IMAGE_MORE, confidence=config.CONFIDENCE_MENU, attempts=config.RETRY_DEFAULT, controller=controller, log=log):
            controller.sleep(random.uniform(*config.DELAY_MENU))

            if automation.find_and_click(config.USER_IMAGE_REPORT_BUTTON, attempts=config.RETRY_FAST, controller=controller, log=log):
                controller.sleep(random.uniform(*config.DELAY_DIALOG))

                # 5. 选择分类：先点“内容违规”，再点具体类型
                automation.find_and_click(config.USER_IMAGE_MAIN_REASON, click_offset="left", controller=controller, log=log)
                if report_type in config.USER_REPORT_REASONS:
                    controller.sleep(random.uniform(*config.DELAY_REASON))
                    automation.find_and_click(
                        config.USER_REPORT_REASONS[report_type], click_offset="left", controller=controller, log=log
                    )

                # 6. 填写详细描述
                if detail and automation.find_and_click(config.USER_IMAGE_INPUT, controller=controller, log=log):
                    automation.input_text(detail, controller=controller)

                # 7. 提交确认
                if automation.find_and_click(config.USER_IMAGE_CONFIRM, confidence=config.CONFIDENCE_CONFIRM, attempts=config.RETRY_DEFAULT, controller=controller, log=log):
                    log(f"  [✓] 第 {row_index} 行举报提交成功")
                    controller.sleep(random.uniform(*config.DELAY_SUBMIT))
    else:
        log("  [X] 搜索结果中未找到目标用户")

    # 8. 清理现场：关闭当前标签页
    pyautogui.hotkey("ctrl", "w")
    controller.sleep(random.uniform(*config.DELAY_CLEANUP))

