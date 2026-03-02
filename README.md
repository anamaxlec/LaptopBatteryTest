# BatteryAutoTest - 笔记本续航与性能自动化测试系统

这是一个基于Python开发的自动化续航测试工具，旨在模拟真实用户使用场景（办公、网页、视频、通讯），并在测试过程中实时记录硬件性能表现。

## 1.主要功能
* **Office生产力测试**：调用UL Procyon执行Word、Excel、PowerPoint自动化负载，并提取离电性能分数。
* **网页浏览模拟**：基于Playwright自动化驱动Edge浏览器，循环加载主流电商、社交及资讯网站并模拟翻页操作。
* **流媒体播放**：模拟长时间观看Bilibili高清在线视频。
* **通讯软件模拟**：自动定位QQ与微信窗口，模拟真实键鼠输入发送消息。
* **数据监控与恢复**：实时监测电池电量，电量低于10%自动进入休眠；支持崩溃自动恢复，脚本重启后可衔接上一轮进度。
* **性能汇总**：自动解析测试结果并生成`performance_summary.txt`性能简报。

## 2.环境要求
### 硬件与软件
* **操作系统**：Windows 10/11
* **必要软件**：
    * Microsoft Office (Word, Excel, PowerPoint)
    * [UL Procyon](https://benchmarks.ul.com/procyon) (必须安装并激活，且脚本需要放置在UL Procyon根目录)
    * 微信/QQ (如需测试通讯模块，需提前登录并保持窗口开启)

### Python环境
推荐使用Python 3.10+。需安装以下依赖库：
```bash
pip install psutil playwright pyautogui pypiwin32
python -m playwright install chromium
