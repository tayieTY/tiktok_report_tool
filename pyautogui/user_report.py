import time, pyautogui, pyperclip, openpyxl, os, random, sys

"""
User Report Module
功能：处理抖音用户（账号）举报的自动化流程
逻辑：搜索用户 -> 进入主页 -> 点击粉丝 -> 更多 -> 举报
"""

# 基础配置
DEFAULT_EXCEL_PATH = r"D:\抖音举报.xlsx"
DEFAULT_SHEET_NAME = "举报指定用户"

# 处理打包后的资源路径问题
BASE_PATH = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
IMAGE_DIR = os.path.join(BASE_PATH, "images")

# 举报原因映射表
REPORT_REASONS = {
    "内容违规": "nr.png", "色情低俗": "nr_sqds.png",
    "不实信息": "nr_bsxx.png", "政治敏感": "nr_zzmg.png"
}


def find_and_click(img_name, confidence=0.7, click_offset=None, attempts=15):
    """
    统一查找并点击图像函数。

    Args:
        img_name (str): 图片文件名。
        confidence (float): 图像识别的匹配度阈值。
        click_offset (str): 点击位置偏移，None为中心，'left'为左侧。
        attempts (int): 尝试查找的次数。

    Returns:
        bool: 操作是否成功。
    """
    path = os.path.join(IMAGE_DIR, img_name)
    if not os.path.exists(path):
        print(f"  [!] 警告：找不到图片文件: {img_name}")
        return False

    for i in range(attempts):
        try:
            pos = pyautogui.locateOnScreen(path, confidence=confidence)
            if pos:
                center = pyautogui.center(pos)
                # 根据偏移设置确定点击坐标
                if click_offset == "left":
                    x, y = (pos.left + 15, center.y)
                else:
                    x, y = center

                # 随机微调点击坐标，防止被判定为机器人
                pyautogui.moveTo(x + random.randint(-3, 3), y + random.randint(-3, 3),
                                 duration=random.uniform(0.2, 0.4))
                pyautogui.click()
                return True
        except Exception:
            pass

        # 循环内的等待时间保持随机
        time.sleep(random.uniform(0.3, 1.4))

    return False


def input_text(text, clear=True):
    """
    输入文本。
    通过系统剪贴板进行粘贴，比 typewrite 支持更好的中文输入。
    """
    if not text: return
    pyperclip.copy(text)
    if clear:
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')

    pyautogui.hotkey('ctrl', 'v')
    # 粘贴后随机等待，给系统反应时间
    time.sleep(random.uniform(1.8, 2.5))


def process_row(search_text, b_col, c_col, row_idx):
    """
    执行单行数据的举报流程。

    Args:
        search_text: 搜索关键词（用户名/抖音号）
        b_col: 举报大类
        c_col: 详细描述
        row_idx: 当前行号
    """
    print(f"\n>>> 正在处理第 {row_idx} 行: {search_text}")

    # 1. 点击搜索框并输入
    if not find_and_click("search_box_2.png", attempts=10):
        print("  [X] 找不到搜索框，跳过本条")
        return

    input_text(search_text)
    pyautogui.press('enter')
    pyautogui.hotkey("ctrl", "0")

    print("  [⏳] 搜索结果加载中，等待页面刷新...")
    # 长时间等待点，极速模式下会被显著优化
    time.sleep(random.uniform(2.5, 3.5))

    # 2. 点击用户头像进入主页
    if find_and_click("user.png", confidence=0.7, attempts=15):
        print("  [√] 已点击用户头像")
        time.sleep(random.uniform(2.5, 3.8))

        # 3. 点击“粉丝”位置以定位（辅助定位）
        if find_and_click("fans.png", confidence=0.7, attempts=15):
            print("  [√] 已定位粉丝入口")
            time.sleep(random.uniform(2.2, 2.7))
        else:
            print("  [!] 未找到粉丝图标，尝试直接寻找菜单...")

        # 4. 点击“更多”(...) 并执行举报
        if find_and_click("more.png", confidence=0.8, attempts=15):
            time.sleep(random.uniform(1.8, 2.2))

            if find_and_click("report_button_2.png", attempts=10):
                time.sleep(random.uniform(2.1, 2.6))

                # 选择举报分类 (默认先点第一项，再点具体项)
                find_and_click("nr.png", click_offset="left")
                if b_col in REPORT_REASONS:
                    time.sleep(random.uniform(2.5, 2.9))
                    find_and_click(REPORT_REASONS[b_col], click_offset="left")

                # 填写详细描述
                if c_col and find_and_click("click_put_in_2.png"):
                    input_text(c_col)

                # 提交确认
                if find_and_click("confirm_2.png", confidence=0.9, attempts=15):
                    print(f"  [√] 第 {row_idx} 行举报提交成功")
                    time.sleep(random.uniform(2.5, 3.5))
    else:
        print("  [X] 搜索结果中未找到目标用户")

    # 5. 清理现场：关闭当前标签页，准备下一条
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(random.uniform(2.2, 4.0))


def main(excel_path=None, sheet_name=None, auto_exit=False):
    if not excel_path:
        raw_input = input("请输入Excel路径: ").strip()
        excel_path = raw_input.replace('"', '').replace("'", "") or DEFAULT_EXCEL_PATH
    if not sheet_name:
        sheet_name = input(f"Sheet名 (默认 {DEFAULT_SHEET_NAME}): ").strip() or DEFAULT_SHEET_NAME

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

        # 启动浏览器
        os.system("start chrome https://www.douyin.com/")
        print("  [⏳] 正在启动浏览器并打开抖音...")
        time.sleep(random.uniform(2.5, 3.5))

        pyautogui.hotkey('alt', 'space')
        pyautogui.press('x')  # 窗口最大化
        time.sleep(random.uniform(2.0, 4.0))

        for row in range(2, sheet.max_row + 1):
            search_val = str(sheet.cell(row, 1).value or "").strip()
            if not search_val: continue

            process_row(search_val,
                        str(sheet.cell(row, 2).value or "").strip(),
                        str(sheet.cell(row, 3).value or "").strip(),
                        row)

        wb.close()
        print(f"\n[√] {sheet_name} 任务处理完毕")

    except Exception as e:
        print(f"❌ 运行中出错: {e}")

    if not auto_exit:
        input("\n按回车键返回...")


if __name__ == "__main__":
    main()