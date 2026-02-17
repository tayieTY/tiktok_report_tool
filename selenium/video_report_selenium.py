import time,os,random,openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ================= 配置区 =================
DEFAULT_EXCEL_PATH = r"D:\抖音举报.xlsx"
DEFAULT_SHEET_NAME = "举报指定视频"


def get_driver():
    """启动带用户配置文件的 Chrome"""
    print("⏳ 正在初始化浏览器...")
    co = Options()
    co.page_load_strategy = 'eager'
    base_path = os.path.dirname(os.path.abspath(__file__))
    user_data_dir = os.path.join(base_path, "User_Data")
    co.add_argument(f"--user-data-dir={user_data_dir}")
    co.add_experimental_option("excludeSwitches", ["enable-automation"])
    co.add_experimental_option('useAutomationExtension', False)
    co.add_argument("--start-maximized")
    co.add_argument("--log-level=3")

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=co)
        return driver
    except Exception as e:
        print(f"\n❌ 浏览器启动失败: {e}")
        return None


def js_click(driver, element):
    """使用 JS 强制点击，绕过 React 层级拦截"""
    driver.execute_script("arguments[0].click();", element)


def process_one_video(driver, input_content, reason_main, reason_sub, detail):
    print(f"\n>>> 正在处理: {input_content}")
    is_url = "http" in str(input_content)
    main_window = driver.current_window_handle

    try:
        # 1. 访问页面逻辑 (保持不变)
        if is_url:
            driver.get(input_content)
            time.sleep(3)
        else:
            driver.get("https://www.douyin.com/")
            search_input = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[data-e2e='search-input']")))
            search_input.send_keys(Keys.CONTROL, "a", Keys.BACK_SPACE)
            search_input.send_keys(input_content, Keys.ENTER)
            # 点击第一个结果并切换窗口
            first_video = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "ul[data-e2e='scroll-list'] li:first-child a")))
            first_video.click()
            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(2)

        # 2. 呼出举报弹窗
        print("    正在打开举报界面...")
        more_btn = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, "//*[@data-e2e='video-share-container'] | //*[contains(@class, 'more')]")))
        driver.execute_script("arguments[0].click();", more_btn)

        report_btn = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[text()='举报']")))
        driver.execute_script("arguments[0].click();", report_btn)

        # --- 核心：强力点击函数 ---
        def force_react_click(text):
            # 找到文字 span
            target_span = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.XPATH, f"//span[contains(text(), '{text}')]"))
            )
            # 通过 JS 获取圆圈和容器
            driver.execute_script("""
                var span = arguments[0];
                var circle = span.previousElementSibling;
                var container = span.parentNode;

                var evOpts = {bubbles: true, cancelable: true, view: window};

                // 对圆圈和容器同时派发完整的事件链
                [circle, container].forEach(function(el) {
                    if(!el) return;
                    el.dispatchEvent(new MouseEvent('mousedown', evOpts));
                    el.dispatchEvent(new MouseEvent('mouseup', evOpts));
                    el.click();
                    el.dispatchEvent(new Event('change', evOpts));
                });
            """, target_span)
            return True

        # 3. 选择一级理由
        print(f"    尝试点击一级理由: {reason_main}")
        time.sleep(1.5)
        force_react_click(reason_main)

        # 4. 选择二级理由
        if reason_sub:
            print(f"    尝试点击二级理由: {reason_sub}")
            time.sleep(1.5)  # 给二级菜单渲染留够时间
            try:
                force_react_click(reason_sub)
            except:
                print(f"    ⚠️ 未找到二级理由: {reason_sub}")

        # 5. 填写描述
        if detail:
            try:
                textarea = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.TAG_NAME, "textarea")))
                textarea.click()
                textarea.clear()
                textarea.send_keys(detail)
                # 必须派发 input 事件，React 才会把文字存入 state
                driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", textarea)
            except:
                pass

        # 6. 提交
        print("    正在尝试提交...")
        try:
            # 这里的定位器只找“提交”
            submit_btn = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.XPATH, "//button[contains(., '提交')] | //*[text()='提交']"))
            )

            # 强行破解：有些按钮被 React 设为 disabled，JS 可以强行点击
            driver.execute_script("""
                var btn = arguments[0];
                btn.disabled = false;
                btn.removeAttribute('disabled');
                btn.click();
                // 如果是 div 模拟的按钮
                var ev = new MouseEvent('click', {bubbles: true, cancelable: true});
                btn.dispatchEvent(ev);
            """, submit_btn)

            print(f"    🚀 [提交指令已发出]")
            time.sleep(2)
        except:
            print("    ❌ 无法找到提交按钮")

    except Exception as e:
        print(f"    ❌ 发生异常: {str(e)[:100]}")
    finally:
        # 清理逻辑
        if not is_url and len(driver.window_handles) > 1:
            driver.close()
            driver.switch_to.window(main_window)


def main():
    try:
        path_input = input(f"Excel路径 (默认 {DEFAULT_EXCEL_PATH}): ").strip()
        excel_path = path_input.replace('"', '') if path_input else DEFAULT_EXCEL_PATH

        driver = get_driver()
        if not driver: return

        print("🚨 准备就绪...")
        input(">>> 按【回车键】开始...")

        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet = wb[DEFAULT_SHEET_NAME]

        for row in range(2, sheet.max_row + 1):
            data = [sheet.cell(row, c).value for c in range(1, 5)]
            if not data[0]: continue
            process_one_video(driver, data[0], data[1], data[2], data[3])
            time.sleep(random.uniform(2, 4))

    except Exception as e:
        print(f"\n❌ 错误: {e}")
    input("\n程序结束，按回车退出...")


if __name__ == "__main__":
    main()