"""统一任务控制：暂停 / 停止 / 极速模式，UI 与业务通过同一个对象协作。"""
from __future__ import annotations

import threading
import time


class TaskStopped(Exception):
    """用户主动停止任务。"""


class TaskController:
    def __init__(self, fast_mode: bool = False):
        self.fast_mode = fast_mode
        self._pause = threading.Event()
        self._pause.set()
        self._stop = threading.Event()

    # ---------------- 状态控制 ----------------
    def pause(self) -> None:
        self._pause.clear()

    def resume(self) -> None:
        self._pause.set()

    def request_stop(self) -> None:
        self._stop.set()
        self._pause.set()

    @property
    def paused(self) -> bool:
        return not self._pause.is_set()

    def check(self) -> None:
        """每次操作前调用；已请求停止则抛出 TaskStopped。"""
        if self._stop.is_set():
            raise TaskStopped("任务已被手动停止")

    # ---------------- 等待 ----------------
    def sleep(self, seconds: float) -> None:
        """可暂停、可停止、可被极速模式缩短的等待。"""
        while not self._pause.is_set():
            self.check()
            time.sleep(0.1)
        rest = self._scaled(seconds)
        while rest > 0:
            self.check()
            step = min(rest, 0.5)
            time.sleep(step)
            rest -= step

    def _scaled(self, seconds: float) -> float:
        if not self.fast_mode:
            return seconds
        if seconds > 15:
            return max(1.0, seconds - 10)
        if 2 < seconds < 3:
            return max(0.5, seconds - 1)
        if 3 < seconds < 4:
            return max(0.5, seconds - 2)
        if 4 < seconds < 5:
            return max(0.5, seconds - 3)
        return seconds
