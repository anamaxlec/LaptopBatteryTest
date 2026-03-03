# 笔记本续航测试工具

一个功能完善的笔记本续航与性能自动化测试系统，支持 Office 生产力测试、网页浏览、视频播放、社交软件模拟等多种测试场景。

## 功能特性

### 测试模块

| 模块 | 说明 | 依赖 |
|------|------|------|
| Office 生产力测试 | Word、Excel、PowerPoint 性能测试 | UL Procyon |
| 网页浏览测试 | 自动打开多个主流网站并模拟滚动浏览 | Playwright |
| 视频播放测试 | Bilibili 视频播放测试 | Playwright |
| 社交软件模拟 | QQ、微信等窗口激活与交互模拟 | PyAutoGUI |
| 实时性能监控 | CPU 功耗、频率、温度监控 | HWiNFO / psutil |

### 核心功能

- ✅ **模块化测试** - 测试前可选择开启/关闭各功能模块
- ✅ **浏览器登录状态保持** - 预登录模式保存 Cookie，避免重复登录
- ✅ **精准功耗统计** - 支持 HWiNFO CSV 日志读取，计算平均 CPU 功耗
- ✅ **续航数据记录** - 每轮测试记录电量、功耗、运行时间
- ✅ **进度保存恢复** - 支持测试中断后恢复继续
- ✅ **低电量自动休眠** - 电量低于 10% 自动进入休眠

## 安装依赖

```bash
pip install playwright psutil pyautogui pywin32
playwright install msedge
```

### 可选依赖

- **UL Procyon** - 用于 Office 生产力测试
- **HWiNFO64** - 用于精准 CPU 功耗监控

## 使用方法

### 1. 直接运行 Python 脚本

```bash
python test_script-V3.py
```

### 2. 使用打包后的 EXE

```bash
# 打包命令
python -m PyInstaller test_script-V3.spec --clean

# 运行
.\dist\test_script-V3.exe
```

## 使用流程

### 首次运行

1. **选择测试模块** - 根据提示选择要启用的测试功能
2. **预登录**（如启用网页/视频测试）- 在浏览器中登录各网站，完成后按回车
3. **开始测试** - 脚本自动循环执行测试直到电量耗尽

### 后续运行

- 自动使用已保存的登录状态
- 可选择是否重新登录
- 支持从上次进度恢复

## HWiNFO 配置（推荐）

为获得精准的 CPU 功耗数据，建议配置 HWiNFO：

1. 打开 HWiNFO64
2. 菜单：`File` → `Preferences` → `Log to CSV`
3. 勾选 `Enabled`
4. 设置轮询间隔为 2-5 秒
5. 日志文件默认保存在 `C:\Program Files\HWiNFO64\`

脚本会自动识别 CSV 文件并计算每轮测试的平均功耗。

## 输出文件

| 文件 | 说明 |
|------|------|
| `battery_stats.csv` | 电池统计数据（轮次、时间、电量、功耗） |
| `performance_summary.txt` | 测试摘要报告 |
| `performance_details.txt` | 详细性能监控数据 |
| `browser_storage_state.json` | 浏览器登录状态 |
| `test_progress.txt` | 测试进度（用于恢复） |

## 测试报告示例

```
============================================================
                  续航测试统计报告
============================================================
开始时间: 2024-01-15 09:00:00
结束时间: 2024-01-15 14:30:00
总运行时长: 330.5 分钟 (5.51 小时)
完成轮数: 12 轮
初始电量: 100%
结束电量: 8%
总耗电: 92%
平均耗电速率: 16.72%/小时
============================================================
```

## 配置说明

脚本内置 `CONFIG` 字典可自定义默认设置：

```python
CONFIG = {
    "ENABLE_OFFICE": True,       # Office 生产力测试
    "ENABLE_WEB": True,          # 网页浏览测试
    "ENABLE_VIDEO": True,        # 视频播放测试
    "ENABLE_CHAT": True,         # 社交软件模拟
    "ENABLE_MONITOR": True,      # 实时性能监控
    "MONITOR_INTERVAL": 10,      # 性能抓取间隔（秒）
}
```

## 支持的社交软件

- QQ
- 微信
- TIM
- 钉钉
- 飞书

## 注意事项

1. **管理员权限** - 部分功能需要管理员权限（UAC）
2. **Edge 浏览器** - 网页测试需要安装 Microsoft Edge
3. **登录状态** - 首次运行需要手动登录各网站
4. **HWiNFO** - 如不使用 HWiNFO，脚本会自动回退到 psutil 估算功耗

## 打包说明

使用 PyInstaller 打包为单文件 EXE：

```bash
# 使用 spec 文件（推荐）
python -m PyInstaller test_script-V3.spec --clean

# 或命令行
python -m PyInstaller --onefile --uac-admin --collect-all playwright test_script-V3.py
```

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！
