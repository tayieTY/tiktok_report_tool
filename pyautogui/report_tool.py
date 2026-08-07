"""命令行入口：不依赖界面也能运行（供脚本或调试使用）。"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pyautogui import config
from pyautogui.core import user_report, video_report
from pyautogui.core.controller import TaskController


def main() -> None:
    print("=" * 70)
    print("      抖音自动化举报综合工具箱 (v1.1 CLI)")
    print("=" * 70)
    path_input = input("Excel 绝对路径 (回车默认 D:\\抖音举报.xlsx): ").strip()
    excel_path = path_input.strip('"\'') if path_input else config.DEFAULT_EXCEL_PATH
    choice = input("模式 (1:视频 2:用户 3:混合): ").strip()
    fast = input("极速模式？(y/n): ").strip().lower() == "y"
    controller = TaskController(fast_mode=fast)

    try:
        if choice == "1":
            video_report.run_video_tasks(controller, excel_path=excel_path, sheet_name=config.VIDEO_SHEET_NAME)
        elif choice == "2":
            user_report.run_user_tasks(controller, excel_path=excel_path, sheet_name=config.USER_SHEET_NAME)
        elif choice == "3":
            video_report.run_video_tasks(controller, excel_path=excel_path, sheet_name=config.VIDEO_SHEET_NAME)
            user_report.run_user_tasks(controller, excel_path=excel_path, sheet_name=config.USER_SHEET_NAME)
        else:
            print("输入无效")
    except Exception as exc:
        print(f"运行出错: {exc}")

    input("\n按回车退出...")


if __name__ == "__main__":
    main()

