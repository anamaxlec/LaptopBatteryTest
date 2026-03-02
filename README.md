# LaptopBatteryTest - 笔记本续航与性能自动化测试系统

这是一个基于 Python 开发的自动化测试工具，旨在模拟真实用户日常使用场景（办公、网页浏览、流媒体、即时通讯），并在评估续航时间的同时，实时记录笔记本在离电状态下的性能分数表现。

## 1. 主要功能

* **Office 生产力评估**：深度调用 UL Procyon 办公自动化测试，覆盖 Word、Excel、PowerPoint 真实负载，并自动解析导出分数。
* **性能实时汇总**：每轮测试结束后自动提取性能得分并追加至 `performance_summary.txt`，直观展现离电性能下降情况。
* **网页负载模拟**：基于 Playwright 驱动 Edge 浏览器，循环加载京东、淘宝、IT 之家、知乎等主流站点，模拟真实翻页操作。
* **高清视频回放**：自动跳转至指定 Bilibili 高清视频链接并进行长时间连续播放。
* **即时通讯模拟**：自动查找并激活 QQ 与微信窗口，模拟键鼠输入发送消息，还原后台通讯软件的真实耗电。
* **进度自动恢复**：内置断点保护，如遇异常崩溃或系统重启，脚本启动后可自动从上一轮进度继续测试。

## 2. 环境要求

### 硬件与软件环境

* **操作系统**：Windows 10/11
* **必要软件**：
* Microsoft Office (Word, Excel, PowerPoint)
* [UL Procyon](https://benchmarks.ul.com/procyon) (必须安装并激活)
* 微信/QQ (如需测试通讯模块，需提前登录并保持窗口开启)



### Python 依赖

推荐使用 Python 3.10+。需安装以下依赖库：

```bash
pip install psutil playwright pyautogui pypiwin32
python -m playwright install chromium

```

## 3. 快速开始

* **放置脚本**：将打包好的 `test_script.exe` 放置在 UL Procyon 的安装根目录下（通常为 `C:\Program Files\UL\Procyon\`）。
* **准备配置**：确保目录下存在 `office_productivity.def`，这是执行 Office 测试的必要定义文件。
* **运行测试**：
1. 右键点击 `test_script.exe`，选择**以管理员身份运行**。
2. 脚本会自动启动循环测试直至电量耗尽。


* **查看结果**：测试结束后，在同级目录下查看 `performance_summary.txt` 获取各轮次的性能分数汇总。

## 4. 文件说明

* **test_script.py**：主程序源码。
* **office_productivity.def**：Procyon 测试定义配置文件。
* **performance_summary.txt**：自动生成的性能分数记录表。
* **test_progress.txt**：自动生成的断点进度文件。

## 5. 注意事项

* **窗口焦点**：由于涉及键鼠模拟，测试开始后请勿人工干扰键鼠，否则可能导致通讯模块发送失败。
* **安全策略**：部分通讯软件可能会拦截自动化输入，运行前请确保脚本已获得管理员权限。
* **免责声明**：本脚本仅用于个人硬件测试，请勿用于任何违规用途。

## 6. 许可证

[MIT License](https://www.google.com/search?q=LICENSE)
