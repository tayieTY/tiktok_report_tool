import time, os, sys, random, pyautogui, pyperclip, openpyxl

"""
Video Report Module
功能：处理抖音视频举报的自动化流程
包含：图像识别、鼠标模拟、表单填写
"""


# 配置与路径管理
class Config:
    DEFAULT_EXCEL = r"D:\抖音举报.xlsx"
    DEFAULT_SHEET = "举报指定视频"
    # 浏览时的随机点击坐标池
    COORDINATES = [(703, 415), (1000, 490), (800, 600)]

    @staticmethod
    def get_path(filename):
        """处理图像资源路径，兼容 PyInstaller 打包后的临时目录"""
        base = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(base, "images", filename)
        # 备用路径逻辑
        if not os.path.exists(path) and getattr(sys, 'frozen', False):
            path = os.path.join(os.path.dirname(sys.executable), "images", filename)
        return path


# 图片文件名映射表
IMG_MAP = {
    'report': "report_button.png", 'search': "search_box.png",
    'input': "click_put_in.png", 'confirm': "confirm.png",
    # 一级原因
    '政治敏感': "zz.png", '色情低俗': "sq.png", '不实信息': "bs.png",
    # 二级原因
    '涉政不当言论': "zz_szbdyl.png", '涉政不实信息': "zz_szbsxx.png",
    '色情裸露内容': "sq_sqllnr.png", '未成年低俗': "sq_wcnds.png", '疑似招嫖': "sq_yszp.png",
    '刻意抹黑': "bs_kymh.png", '虚假摆拍演绎': "bs_xjbpyy.png",
    '疑似虚假时事': "bs_ysxjss.png", '疑似虚假知识': "bs_ysxjzs.png"
}


def universal_click(img_key, conf=0.9, mode='center', retry=20):
    """
    通用图像识别点击函数。

    Args:
        img_key (str): IMG_MAP 中的键名，或直接的图片文件名。
        conf (float): 识别置信度 (0.0 - 1.0)。
        mode (str): 点击模式。'center' 点击图片中心；'left' 点击图片左侧偏移处（用于复选框）。
        retry (int): 最大重试次数。

    Returns:
        bool: 是否点击成功。
    """
    filename = IMG_MAP.get(img_key, img_key)
    path = Config.get_path(filename)

    if not os.path.exists(path):
        print(f"❌ [错误] 找不到图片资源: {filename}")
        return False

    for i in range(retry):
        try:
            res = pyautogui.locateOnScreen(path, confidence=conf)
            if res:
                if mode == 'center':
                    # 添加微小的随机偏移，模拟真人操作
                    target_x = res.left + res.width // 2 + random.randint(-1, 1)
                    target_y = res.top + res.height // 2 + random.randint(-1, 1)
                else:
                    # 'left' 模式：针对勾选框，根据图片类型调整偏移量
                    offset = 14 if "_" in filename else 10
                    target_x = res.left + offset
                    target_y = res.top + res.height // 2

                pyautogui.moveTo(target_x, target_y, duration=0.2)
                pyautogui.click()
                print(f"✓ 成功点击 ({'二级' if '_' in filename else '一级'}): {img_key}")
                time.sleep(1)
                return True
        except Exception:
            pass

        if i % 5 == 0 and i > 0:
            print(f"正在重试寻找: {img_key}...")

        time.sleep(random.uniform(0.5, 1.0))

    print(f"✗ 识别失败: {img_key}")
    return False


def input_text(text):
    """模拟粘贴文本"""
    if not text: return
    pyperclip.copy(str(text))
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(random.uniform(2.3, 3.4))


def process_report(row_idx, data):
    """
    处理单行举报逻辑。

    Args:
        row_idx (int): Excel 行号。
        data (list): [搜索词, 一级原因, 二级原因, 详细描述]。
    """
    search_q, reason_main, reason_sub, detail = data
    print(f"\n🚀 处理第 {row_idx} 行: {search_q}")

    # 新开标签页
    pyautogui.hotkey('ctrl', 't')
    time.sleep(random.uniform(2.7, 3.9))

    if universal_click('search'):
        input_text(search_q)
        pyautogui.press('enter')
        pyautogui.hotkey("ctrl", "0")
        # 等待搜索结果加载，此时间较长，会被极速模式优化
        time.sleep(random.uniform(15.0, 16.4))

        # 模拟真人浏览行为：移动鼠标 + 滚动
        pyautogui.moveTo(*random.choice(Config.COORDINATES), duration=0.5)
        time.sleep(random.uniform(2.5, 3.4))
        pyautogui.scroll(-250)
        time.sleep(random.uniform(2.3, 3.4))

        if universal_click('report',conf=0.99):
            pyautogui.press('enter')
            # 点击一级原因
            if universal_click(reason_main, mode='left'):
                # 点击二级类型（如果有）
                if reason_sub:
                    universal_click(reason_sub, mode='left')

                # 填写详细描述（如果有）
                if detail:
                    if universal_click('input'):
                        input_text(detail)

                universal_click('confirm')
                time.sleep(random.uniform(2.5, 3.4))

    # 关闭当前标签页
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(random.uniform(2.3, 3.0))


def main(excel_path=None, sheet_name=None, auto_exit=False):
    """
    Args:
        excel_path (str): Excel 文件路径
        sheet_name (str): 工作表名称
        auto_exit (bool): 任务完成后是否自动退出（UI模式下为True）
    """
    if not excel_path:
        raw_input = input("请输入Excel路径: ").strip()
        excel_path = raw_input.replace('"', '').replace("'", "") or r"D:\抖音举报.xlsx"

    if not sheet_name:
        sheet_name = "举报指定视频"

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        print(f"📅 任务已就绪，工作表: {sheet.title}，总计 {sheet.max_row - 1} 条数据")

        # 启动 Chrome
        os.system("start chrome")
        time.sleep(3)
        # 窗口最大化操作
        pyautogui.hotkey('alt', 'space')
        pyautogui.hotkey('x')
        pyautogui.hotkey('alt')

        for r in range(2, sheet.max_row + 1):
            row_vals = [sheet.cell(r, c).value for c in range(1, 5)]
            if not row_vals[0]: continue

            process_report(r, row_vals)

        wb.close()
        print(f"\n[√] {sheet_name} 任务处理完毕")

    except Exception as e:
        print(f"❌ 运行中出错: {e}")

    if not auto_exit:
        input("\n按回车键返回...")


if __name__ == "__main__":
    main()