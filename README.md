# 抖音自动化举报工具箱（v1.1，pyautogui 版）

【开源声明】本软件为免费开源项目，仅供个人学习与技术研究，严禁用于任何商业用途。

## v1.1 重构说明

本仓库在保留原技术（Tkinter + pyautogui 图像识别）的基础上，修复了两个结构性问题：

1. **UI 与业务逻辑解耦**
   - 旧版：`ui_main.py` 直接调用业务模块，并通过“猴子补丁”`video_report.time.sleep = self.smart_sleep` 把界面逻辑注入业务代码。
   - 新版：界面（`ui_main.py`）只负责展示与输入；业务（`pyautogui/core/`）完全独立，通过统一对象 `TaskController` 响应暂停/停止/极速模式，通过回调输出日志。
2. **消灭魔法数**
   - 旧版：屏幕坐标、匹配阈值、重试次数、等待时间散落在业务代码里。
   - 新版：全部集中到 `pyautogui/config.py`，统一命名、一眼可调。

> 更现代的 v2 重构版（Playwright DOM 自动化 + PySide6）见 [tiktok-report-tool-v2](https://github.com/tayieTY/tiktok-report-tool-v2)；本仓库的原始未重构版本保存在 git 标签 `v1-legacy`。

## 项目结构

```text
pyautogui/
├── config.py            # 配置中心：所有可调参数（阈值/重试/等待/图片映射）
├── ui_main.py           # 界面层：只负责展示与用户输入
├── report_tool.py       # CLI 入口：不需要界面也能运行
├── core/
│   ├── controller.py    # 统一控制：暂停 / 停止 / 极速模式
│   ├── automation.py    # 底层自动化：图像识别点击、文本输入
│   ├── excel_reader.py  # Excel 读取（视频/用户共用）
│   ├── video_report.py  # 视频举报业务流程
│   └── user_report.py   # 用户举报业务流程
└── images/              # 图像识别素材
selenium/                # 半成品 Selenium 方案（开发中）
```

## 使用须知

1. 提前下载 Google Chrome 浏览器，登录个人抖音账号网页版。
2. 在模板 excel 里按分类依次录入举报的账号（视频）、举报类型、举报理由。
3. 运行方式（二选一）：
   - 图形界面：`python -m pyautogui.ui_main`（或直接双击 `pyautogui/ui_main.py`）
   - 命令行：`python -m pyautogui.report_tool`
4. 屏幕分辨率必须为 **1920x1080、缩放 100%**（图像识别依赖固定屏幕环境，界面里也有“打开屏幕设置”按钮）。
5. 软件运行期间请勿移动鼠标和触碰键盘。
6. 若电脑性能不足，请慎重选择极速模式。

## 已知限制

- 举报用户时，“举报原因”只有“内容违规”，“举报类型”只有“不实信息”“色情低俗”“政治敏感”三种选择。
- 举报视频时，“举报原因”只有“不实信息”“色情低俗”“政治敏感”三种选择，但三个大类下的子类“举报类型”是完善的。
- 图像识别对屏幕环境敏感，这是 pyautogui 方案的固有限制；需要跨分辨率、抗改版能力请使用 v2（Playwright 版）。

⚠️ 免责声明：请合理合法使用本工具，用户使用本工具产生的一切后果由用户自行承担。

