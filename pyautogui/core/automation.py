"""底层自动化能力：图像识别点击、文本输入。

UI 与业务流程都不直接调用 pyautogui，而是使用本模块封装好的能力。
"""
from __future__ import annotations

import os
import random
import time
from typing import Callable, Optional

import pyautogui
import pyperclip

from pyautogui import config
from pyautogui.core.controller import TaskController


def image_path(filename: str) -> str:
    return os.path.join(config.IMAGE_DIR, filename)


def _sleep(controller: Optional[TaskController], seconds: float) -> None:
    if controller is not None:
        controller.sleep(seconds)
    else:
        time.sleep(seconds)


def find_and_click(
    filename: str,
    confidence: float = config.CONFIDENCE_DEFAULT,
    click_offset: Optional[str] = None,
    attempts: int = config.RETRY_DEFAULT,
    controller: Optional[TaskController] = None,
    log: Callable[[str], None] = print,
) -> bool:
    """统一查找并点击图片；找不到时按重试间隔重试。"""
    path = image_path(filename)
    if not os.path.exists(path):
        log(f"  [!] 警告：找不到图片文件: {filename}")
        return False

    for _ in range(attempts):
        if controller is not None:
            controller.check()
        try:
            pos = pyautogui.locateOnScreen(path, confidence=confidence)
            if pos:
                center = pyautogui.center(pos)
                if click_offset == "left":
                    x, y = pos.left + config.CLICK_OFFSET_LEFT, center.y
                else:
                    x, y = center
                pyautogui.moveTo(
                    x + random.randint(-config.CLICK_JITTER, config.CLICK_JITTER),
                    y + random.randint(-config.CLICK_JITTER, config.CLICK_JITTER),
                    duration=random.uniform(*config.DELAY_CLICK),
                )
                pyautogui.click()
                return True
        except Exception:
            pass
        _sleep(controller, random.uniform(*config.DELAY_RETRY))
    return False


def input_text(text: str, controller: Optional[TaskController] = None, clear: bool = True) -> None:
    """通过系统剪贴板粘贴文本，支持中文。"""
    if not text:
        return
    pyperclip.copy(text)
    if clear:
        pyautogui.hotkey("ctrl", "a")
        pyautogui.press("backspace")
    pyautogui.hotkey("ctrl", "v")
    _sleep(controller, random.uniform(*config.DELAY_INPUT))


def universal_click(
    img_key: str,
    conf: float = config.CONFIDENCE_VIDEO,
    mode: str = "center",
    retry: int = config.RETRY_VIDEO,
    controller: Optional[TaskController] = None,
    log: Callable[[str], None] = print,
) -> bool:
    """视频流程专用的图片识别点击（支持 center / left 两种模式）。"""
    filename = config.VIDEO_IMAGE_MAP.get(img_key, img_key)
    path = image_path(filename)
    if not os.path.exists(path):
        log(f"❌ [错误] 找不到图片资源: {filename}")
        return False

    for index in range(retry):
        if controller is not None:
            controller.check()
        try:
            res = pyautogui.locateOnScreen(path, confidence=conf)
            if res:
                if mode == "center":
                    target_x = res.left + res.width // 2 + random.randint(-config.VIDEO_CLICK_JITTER, config.VIDEO_CLICK_JITTER)
                    target_y = res.top + res.height // 2 + random.randint(-config.VIDEO_CLICK_JITTER, config.VIDEO_CLICK_JITTER)
                else:
                    offset = config.CHECKBOX_OFFSET_SECONDARY if "_" in filename else config.CHECKBOX_OFFSET_PRIMARY
                    target_x = res.left + offset
                    target_y = res.top + res.height // 2

                pyautogui.moveTo(target_x, target_y, duration=config.DELAY_CLICK[0])
                pyautogui.click()
                log(f"✅ 成功点击 ({'二级' if '_' in filename else '一级'}): {img_key}")
                time.sleep(config.POST_CLICK_SLEEP)
                return True
        except Exception:
            pass

        if index % 5 == 0 and index > 0:
            log(f"正在重试查找: {img_key}...")
        _sleep(controller, random.uniform(*config.DELAY_RETRY))

    log(f"❌ 识别失败: {img_key}")
    return False

