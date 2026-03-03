import os
import sys
import time
import subprocess
import psutil
import traceback
import csv
import threading
import random
import glob
from datetime import datetime
from playwright.sync_api import sync_playwright
import pyautogui
import win32gui
import win32con

# ===================== [模块开关与自定义配置] =====================
CONFIG = {
    "ENABLE_OFFICE": True,       # Office生产力测试 (需安装UL Procyon)
    "ENABLE_WEB": True,          # 网页浏览测试 (京东/淘宝/IT之家等)
    "ENABLE_VIDEO": True,        # 视频播放测试 (Bilibili指定链接)
    "ENABLE_CHAT": True,         # 社交软件模拟 (QQ/微信窗口激活)
    "ENABLE_MONITOR": True,      # 实时性能与功耗监控
    "MONITOR_INTERVAL": 10,      # 性能抓取间隔 (秒)
}

# 测试目标定义
WEB_TEST_URLS = [
    "https://www.jd.com", "https://www.taobao.com", "https://www.sina.com.cn",
    "https://www.163.com", "https://www.sohu.com", "https://www.ithome.com",
     "https://www.gamersky.com","https://www.3dmgame.com", "https://weibo.com",
    "https://www.bilibili.com"
]
VIDEO_URL = "https://www.bilibili.com/video/BV13vPMzKEzd/"

# 路径配置
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

SUMMARY_LOG = os.path.join(base_dir, "performance_summary.txt")
PERF_DETAIL_LOG = os.path.join(base_dir, "performance_details.txt")
PROGRESS_FILE = os.path.join(base_dir, "test_progress.txt")
STORAGE_STATE_FILE = os.path.join(base_dir, "browser_storage_state.json")
BATTERY_LOG = os.path.join(base_dir, "battery_stats.csv")
# ================================================================

class BatteryTracker:
    """电池续航统计追踪器"""
    def __init__(self):
        self.start_time = None
        self.start_battery = None
        self.round_data = []
        self.lock = threading.Lock()
        self._init_log_file()
    
    def _init_log_file(self):
        """初始化CSV日志文件"""
        if not os.path.exists(BATTERY_LOG):
            with open(BATTERY_LOG, "w", encoding="utf-8") as f:
                f.write("timestamp,round,elapsed_minutes,battery_percent,power_plugged,round_duration_sec,avg_power_w\n")
    
    def start_test(self):
        """开始测试时记录初始状态"""
        battery = psutil.sensors_battery()
        self.start_time = time.time()
        self.start_battery = battery.percent
        self.round_data = []
        print(f"\n=== 续航测试开始 ===")
        print(f"开始时间: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"初始电量: {self.start_battery}%")
        print(f"电源状态: {'已连接' if battery.power_plugged else '电池供电'}\n")
        self._log_round(0, 0, battery.percent, battery.power_plugged, 0, 0)
    
    def record_round(self, round_num, round_duration_sec, hwinfo_reader=None):
        """记录每轮测试后的电池状态
        
        Args:
            round_num: 轮数
            round_duration_sec: 本轮持续时间（秒）
            hwinfo_reader: HWiNFOReader实例，用于获取准确的平均功耗
        """
        battery = psutil.sensors_battery()
        elapsed = (time.time() - self.start_time) / 60  # 分钟
        round_start_time = time.time() - round_duration_sec
        
        # 获取平均功耗
        power_consumption = 0
        power_details = ""
        
        if hwinfo_reader and hwinfo_reader.hwinfo_available:
            # 从HWiNFO CSV计算本轮平均功耗
            avg_data = hwinfo_reader.get_average_since(round_start_time)
            if avg_data and 'cpu_power_avg' in avg_data:
                power_consumption = avg_data['cpu_power_avg']
                samples = avg_data.get('cpu_power_samples', 0)
                power_min = avg_data.get('cpu_power_min', 0)
                power_max = avg_data.get('cpu_power_max', 0)
                power_details = f" | CPU功耗:{power_consumption:.1f}W(avg, n={samples})"
            else:
                # 回退到电池估算
                power_consumption = self._estimate_power_from_battery(round_duration_sec)
                power_details = f" | 估算功耗:{power_consumption:.1f}W"
        else:
            # 使用电池估算
            power_consumption = self._estimate_power_from_battery(round_duration_sec)
            power_details = f" | 估算功耗:{power_consumption:.1f}W"
        
        with self.lock:
            self.round_data.append({
                'round': round_num,
                'elapsed': elapsed,
                'battery': battery.percent,
                'drop': self.start_battery - battery.percent,
                'power': power_consumption
            })
        
        self._log_round(round_num, elapsed, battery.percent, battery.power_plugged, round_duration_sec, power_consumption)
        
        # 打印当前状态
        print(f"\n>>> 第 {round_num} 轮完成 | 已运行: {elapsed:.1f}分钟 | 电量: {battery.percent}% | 累计耗电: {self.start_battery - battery.percent}%{power_details}")
        
        # 估算剩余续航
        if elapsed > 0 and battery.percent > 10:
            drop_rate = (self.start_battery - battery.percent) / elapsed  # %/分钟
            if drop_rate > 0:
                remaining_min = (battery.percent - 5) / drop_rate  # 预留5%安全电量
                print(f">>> 估算剩余续航: {remaining_min:.0f}分钟 ({remaining_min/60:.1f}小时)")
        
        return battery.percent
    
    def _estimate_power_from_battery(self, round_duration_sec):
        """通过电池下降估算功耗"""
        if len(self.round_data) < 1:
            return 0
        
        prev_battery = self.round_data[-1]['battery']
        curr_battery = psutil.sensors_battery().percent
        battery_drop = prev_battery - curr_battery
        
        if battery_drop > 0 and round_duration_sec > 0:
            # 假设电池容量 50Wh
            return (battery_drop / 100 * 50) / (round_duration_sec / 3600)
        return 0
    
    def _log_round(self, round_num, elapsed, battery_pct, plugged, duration_sec, power_w):
        """写入CSV日志"""
        with open(BATTERY_LOG, "a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')},{round_num},{elapsed:.2f},{battery_pct},{int(plugged)},{duration_sec:.0f},{power_w:.2f}\n")
    
    def generate_summary(self, final_round, terminated_by_battery=False):
        """生成最终统计报告"""
        if not self.round_data:
            return
        
        total_time = (time.time() - self.start_time) / 60
        final_battery = self.round_data[-1]['battery']
        total_drop = self.start_battery - final_battery
        
        summary = []
        summary.append("\n" + "="*60)
        summary.append("                  续航测试统计报告")
        summary.append("="*60)
        summary.append(f"开始时间: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(self.start_time))}")
        summary.append(f"结束时间: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        summary.append(f"总运行时长: {total_time:.1f} 分钟 ({total_time/60:.2f} 小时)")
        summary.append(f"完成轮数: {final_round} 轮")
        summary.append(f"初始电量: {self.start_battery}%")
        summary.append(f"结束电量: {final_battery}%")
        summary.append(f"总耗电: {total_drop}%")
        if total_time > 0:
            summary.append(f"平均耗电速率: {total_drop/total_time:.2f}%/小时")
        summary.append("="*60)
        
        summary_text = "\n".join(summary)
        print(summary_text)
        
        # 写入摘要日志
        with open(SUMMARY_LOG, "a", encoding="utf-8") as f:
            f.write(summary_text + "\n")
        
        return summary_text


class HWiNFOReader:
    """通过HWiNFO CSV日志读取传感器数据 - 支持时间同步和平均计算"""
    
    def __init__(self):
        self.hwinfo_available = False
        self.csv_path = None
        self.sensor_columns = {
            'cpu_power': None,
            'cpu_freq': None,
            'cpu_temp': None,
        }
        self.time_column = None
        self._last_read_time = None
        self._find_hwinfo_csv()
    
    def _find_hwinfo_csv(self):
        """查找HWiNFO CSV日志文件"""
        possible_paths = [
            os.path.join(base_dir, "HWiNFO_LOG.csv"),
            os.path.join(os.path.expanduser("~"), "Documents", "HWiNFO", "*.csv"),
            os.path.join(os.path.expanduser("~"), "HWiNFO", "*.csv"),
            "C:\\HWiNFO\\*.csv",
            "C:\\Program Files\\HWiNFO64\\HWiNFO_LOG_*.csv",
            "C:\\Program Files (x86)\\HWiNFO64\\HWiNFO_LOG_*.csv",
        ]
        
        # 收集所有匹配的CSV文件，选择最新的
        all_matches = []
        for path in possible_paths:
            matches = glob.glob(path) if '*' in path else [path]
            for match in matches:
                if os.path.exists(match) and os.path.getsize(match) > 100:
                    # 获取文件修改时间
                    mtime = os.path.getmtime(match)
                    all_matches.append((match, mtime))
        
        # 按修改时间排序，选择最新的
        if all_matches:
            all_matches.sort(key=lambda x: x[1], reverse=True)
            self.csv_path = all_matches[0][0]
            self.hwinfo_available = True
            print(f"[HWiNFO] 找到CSV日志: {self.csv_path}")
            self._scan_csv_header()
            return
        
        print("[HWiNFO] 未找到CSV日志文件，将使用psutil数据")
        print("[HWiNFO] 提示: 在HWiNFO中启用CSV日志: File -> Preferences -> Log to CSV")
    
    def _read_csv_with_encoding(self, read_all=False):
        """尝试多种编码读取CSV文件"""
        encodings = ['utf-8', 'utf-16', 'utf-16-le', 'utf-16-be', 'gbk', 'gb2312', 'cp1252', 'latin1']
        
        for encoding in encodings:
            try:
                with open(self.csv_path, 'r', encoding=encoding) as f:
                    if read_all:
                        return f.readlines()
                    else:
                        return [f.readline().strip()]
            except (UnicodeDecodeError, UnicodeError):
                continue
            except Exception:
                continue
        
        # 如果都失败，使用 latin1（不会报错但可能有乱码）
        try:
            with open(self.csv_path, 'r', encoding='latin1') as f:
                if read_all:
                    return f.readlines()
                else:
                    return [f.readline().strip()]
        except:
            return None
    
    def _scan_csv_header(self):
        """扫描CSV表头，识别传感器列和时间列"""
        try:
            lines = self._read_csv_with_encoding(read_all=False)
            if not lines:
                return
            
            header_line = lines[0]
            if not header_line:
                return
            
            columns = header_line.split(',')
            
            for i, col in enumerate(columns):
                col_lower = col.lower().strip()
                
                # 识别时间列
                if col_lower in ['time', 'timestamp', 'date'] or 'time' in col_lower:
                    self.time_column = i
                
                # 识别CPU功耗
                elif any(k in col_lower for k in ['cpu package power', 'cpu power', 'ppt']):
                    self.sensor_columns['cpu_power'] = i
                    print(f"[HWiNFO] 发现CPU功耗列: {col}")
                
                # 识别CPU频率
                elif any(k in col_lower for k in ['cpu clock', 'core 0 t0', 'cpu freq']):
                    if self.sensor_columns['cpu_freq'] is None:
                        self.sensor_columns['cpu_freq'] = i
                        print(f"[HWiNFO] 发现CPU频率列: {col}")
                
                # 识别CPU温度
                elif 'cpu' in col_lower and 'temperature' in col_lower and 'package' in col_lower:
                    self.sensor_columns['cpu_temp'] = i
                    print(f"[HWiNFO] 发现CPU温度列: {col}")
        
        except Exception as e:
            print(f"[HWiNFO] 扫描CSV表头失败: {e}")
    
    def read_sensors(self):
        """读取最新的传感器数据"""
        if not self.hwinfo_available or not self.csv_path:
            return None
        
        try:
            lines = self._read_csv_with_encoding(read_all=True)
            if not lines or len(lines) < 2:
                return None
            
            # 获取最新数据行（最后一行）
            last_line = lines[-1].strip()
            if not last_line or last_line.lower().startswith('date'):
                last_line = lines[-2].strip()
            
            values = last_line.split(',')
            
            # 更新时间戳
            if self.time_column is not None and self.time_column < len(values):
                self._last_read_time = values[self.time_column]
            
            result = {}
            for sensor_name, col_index in self.sensor_columns.items():
                if col_index is not None and col_index < len(values):
                    try:
                        result[sensor_name] = float(values[col_index])
                    except:
                        result[sensor_name] = None
                else:
                    result[sensor_name] = None
            
            return result
        
        except Exception as e:
            return None
    
    def get_average_since(self, start_time, end_time=None):
        """计算指定时间范围内的平均功耗
        
        Args:
            start_time: 开始时间戳（time.time()格式）
            end_time: 结束时间戳，默认为当前时间
        
        Returns:
            dict: 各传感器的平均值
        """
        if not self.hwinfo_available or not self.csv_path:
            return None
        
        if end_time is None:
            end_time = time.time()
        
        try:
            lines = self._read_csv_with_encoding(read_all=True)
            if not lines or len(lines) < 2:
                return None
            
            # 收集时间范围内的数据
            power_values = []
            freq_values = []
            temp_values = []
            
            for line in lines[1:]:  # 跳过表头
                line = line.strip()
                if not line:
                    continue
                
                values = line.split(',')
                if len(values) < max(self.sensor_columns.values()) + 1:
                    continue
                
                # 解析时间戳
                if self.time_column is not None and self.time_column < len(values):
                    try:
                        # HWiNFO时间格式通常是 "2024-01-15 14:30:15"
                        time_str = values[self.time_column]
                        row_time = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S").timestamp()
                        
                        # 检查是否在时间范围内
                        if start_time <= row_time <= end_time:
                            if self.sensor_columns['cpu_power'] is not None:
                                try:
                                    power_values.append(float(values[self.sensor_columns['cpu_power']]))
                                except:
                                    pass
                            if self.sensor_columns['cpu_freq'] is not None:
                                try:
                                    freq_values.append(float(values[self.sensor_columns['cpu_freq']]))
                                except:
                                    pass
                            if self.sensor_columns['cpu_temp'] is not None:
                                try:
                                    temp_values.append(float(values[self.sensor_columns['cpu_temp']]))
                                except:
                                    pass
                    except:
                        continue
            
            # 计算平均值
            result = {}
            if power_values:
                result['cpu_power_avg'] = sum(power_values) / len(power_values)
                result['cpu_power_min'] = min(power_values)
                result['cpu_power_max'] = max(power_values)
                result['cpu_power_samples'] = len(power_values)
            if freq_values:
                result['cpu_freq_avg'] = sum(freq_values) / len(freq_values)
            if temp_values:
                result['cpu_temp_avg'] = sum(temp_values) / len(temp_values)
            
            return result if result else None
        
        except Exception as e:
            return None


class PerformanceMonitor:
    """实时性能监控模块 - 支持HWiNFO数据"""
    def __init__(self):
        self.is_running = False
        self.data_buffer = []
        self.lock = threading.Lock()
        self.hwinfo = HWiNFOReader()

    def _monitor_loop(self):
        while self.is_running:
            try:
                timestamp = time.strftime("%H:%M:%S")
                battery = psutil.sensors_battery().percent
                
                # 尝试从HWiNFO读取数据
                hw_data = self.hwinfo.read_sensors()
                
                if hw_data and hw_data.get('cpu_power') is not None:
                    # 使用HWiNFO数据
                    power = hw_data.get('cpu_power', 0)
                    freq = hw_data.get('cpu_freq', 0)
                    temp = hw_data.get('cpu_temp', 0)
                    
                    log_line = f"[{timestamp}] 电量:{battery}% | CPU功耗:{power:.2f}W | 频率:{freq:.0f}MHz | 温度:{temp:.0f}°C"
                else:
                    # 回退到psutil数据
                    freq = psutil.cpu_freq().current if psutil.cpu_freq() else 0
                    usage = psutil.cpu_percent(interval=1)
                    log_line = f"[{timestamp}] 电量:{battery}% | 频率:{freq:.0f}MHz | 占用:{usage}% [psutil]"
                
                with self.lock:
                    self.data_buffer.append(log_line)
                    
            except Exception as e:
                pass
            
            time.sleep(CONFIG["MONITOR_INTERVAL"])

    def start(self):
        if not CONFIG["ENABLE_MONITOR"]: return
        self.is_running = True
        self.data_buffer = []
        self.thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.thread.start()

    def stop(self, tag, round_num):
        if not CONFIG["ENABLE_MONITOR"]: return
        self.is_running = False
        if hasattr(self, 'thread'):
            self.thread.join(timeout=1)
        with self.lock:
            if self.data_buffer:
                with open(PERF_DETAIL_LOG, "a", encoding="utf-8") as f:
                    f.write(f"\n>>> Round {round_num} - {tag} 监控明细 <<<\n")
                    f.write("\n".join(self.data_buffer) + "\n")

# 实例化全局监控器
monitor = PerformanceMonitor()
battery_tracker = BatteryTracker()

def focus_window(title_keyword):
    def enum_windows_callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title_keyword in title:
                windows.append(hwnd)
    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)
    if windows:
        hwnd = windows[0]
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        try:
            win32gui.SetForegroundWindow(hwnd)
        except:
            pyautogui.press('alt')
            win32gui.SetForegroundWindow(hwnd)
        return True
    return False

def parse_procyon_csv(csv_path):
    scores = {"Total": "0", "Word": "0", "Excel": "0", "PPT": "0"}
    try:
        if os.path.exists(csv_path):
            with open(csv_path, mode='r', encoding='utf-16') as f:
                reader = csv.reader(f)
                for row in reader:
                    if not row: continue
                    line = row[0]
                    if "Office Productivity Score" in line: scores["Total"] = line.split(',')[-1]
                    if "Word Score" in line: scores["Word"] = line.split(',')[-1]
                    if "Excel Score" in line: scores["Excel"] = line.split(',')[-1]
                    if "PowerPoint Score" in line: scores["PPT"] = line.split(',')[-1]
    except:
        pass
    return scores

def test_office(round_num):
    if not CONFIG["ENABLE_OFFICE"]: return
    print(f"[{time.strftime('%H:%M:%S')}] 开始第{round_num}轮 Office 测试...")
    monitor.start()
    def_file = os.path.join(base_dir, "office_productivity.def")
    csv_file = os.path.join(base_dir, f"round_{round_num}.csv")
    cmd = ["ProcyonCmd.exe", f"--definition={def_file}", f"--out=round_{round_num}.procyon-result", f"--export-csv={csv_file}"]
    try:
        subprocess.run(cmd, check=True)
        s = parse_procyon_csv(csv_file)
        result_str = f"Round {round_num} Office 分数: 总分 {s['Total']} (W:{s['Word']} E:{s['Excel']} P:{s['PPT']})"
        with open(SUMMARY_LOG, "a", encoding="utf-8") as f:
            f.write(result_str + "\n")
        print(result_str)
    finally:
        monitor.stop("Office", round_num)

def test_web(round_num):
    if not CONFIG["ENABLE_WEB"]: return
    print(f"[{time.strftime('%H:%M:%S')}] 开始第{round_num}轮 网页浏览测试...")
    monitor.start()
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False, channel="msedge")
            # 加载已保存的登录状态（如果存在）
            storage_state = STORAGE_STATE_FILE if os.path.exists(STORAGE_STATE_FILE) else None
            context = browser.new_context(storage_state=storage_state)
            pages = [context.new_page() for _ in range(len(WEB_TEST_URLS))]
            for i, url in enumerate(WEB_TEST_URLS):
                pages[i].goto(url)
            for _ in range(3): # 循环滚动查看
                for page in pages:
                    page.bring_to_front()
                    for _ in range(4):
                        page.mouse.wheel(0, 600)
                        time.sleep(1)
            # 保存登录状态供下次使用
            context.storage_state(path=STORAGE_STATE_FILE)
            browser.close()
    finally:
        monitor.stop("Web", round_num)

def test_video(round_num):
    if not CONFIG["ENABLE_VIDEO"]: return
    print(f"[{time.strftime('%H:%M:%S')}] 开始第{round_num}轮 视频播放测试...")
    monitor.start()
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False, channel="msedge")
            # 加载已保存的登录状态（如果存在）
            storage_state = STORAGE_STATE_FILE if os.path.exists(STORAGE_STATE_FILE) else None
            context = browser.new_context(storage_state=storage_state)
            page = context.new_page()
            page.goto(VIDEO_URL)
            time.sleep(690) # 模拟播放11.5分钟
            # 保存登录状态供下次使用
            context.storage_state(path=STORAGE_STATE_FILE)
            browser.close()
    finally:
        monitor.stop("Video", round_num)

class ChatAppTester:
    """社交软件测试器 - 支持多种触发方式和交互模拟"""
    
    # 支持的应用配置: (窗口标题关键词, 进程名, 启动命令)
    SUPPORTED_APPS = [
        ("QQ", ["QQ.exe", "QQ", "腾讯QQ"], None),
        ("微信", ["WeChat.exe", "WeChat", "微信"], None),
        ("TIM", ["TIM.exe", "TIM"], None),
        ("钉钉", ["DingTalk.exe", "DingTalk", "钉钉"], None),
        ("飞书", ["Feishu.exe", "Feishu", "飞书", "Lark"], None),
    ]
    
    def __init__(self):
        self.tested_apps = []
        self.skipped_apps = []
    
    def find_window(self, title_keywords):
        """查找窗口，支持多个关键词"""
        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                title = win32gui.GetWindowText(hwnd)
                for keyword in title_keywords:
                    if keyword in title:
                        results.append((hwnd, title))
                        break
        
        results = []
        win32gui.EnumWindows(enum_callback, results)
        return results
    
    def find_process(self, process_names):
        """检查进程是否运行"""
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                proc_name = proc.info['name'].lower()
                for name in process_names:
                    if name.lower() in proc_name:
                        return proc.info['pid']
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        return None
    
    def activate_window(self, hwnd, title):
        """激活窗口，带重试机制"""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                # 恢复窗口（如果最小化）
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.2)
                
                # 尝试设置前台窗口
                win32gui.SetForegroundWindow(hwnd)
                
                # 验证是否成功
                time.sleep(0.1)
                if win32gui.GetForegroundWindow() == hwnd:
                    return True
                    
            except Exception as e:
                if attempt < max_retries - 1:
                    # 使用Alt键绕过Windows限制
                    pyautogui.keyDown('alt')
                    pyautogui.keyUp('alt')
                    time.sleep(0.1)
                else:
                    print(f"  [Chat] 激活窗口失败: {title}")
        return False
    
    def simulate_chat_activity(self, app_name, duration_sec=60):
        """模拟聊天活动"""
        print(f"  [Chat] 开始模拟 {app_name} 活动 ({duration_sec}秒)")
        
        start_time = time.time()
        message_count = 0
        
        # 模拟随机聊天行为
        actions = [
            self._action_scroll_history,
            self._action_switch_chat,
            self._action_type_message,
            self._action_click_sticker,
            self._action_idle,
        ]
        
        while time.time() - start_time < duration_sec:
            # 随机选择一个动作
            action = random.choice(actions)
            action(app_name)
            
            # 随机间隔 3-8 秒
            time.sleep(random.uniform(3, 8))
            message_count += 1
            
            if message_count % 5 == 0:
                print(f"  [Chat] {app_name} 已执行 {message_count} 个动作")
        
        print(f"  [Chat] {app_name} 活动模拟完成")
    
    def _action_scroll_history(self, app_name):
        """滚动聊天记录"""
        # 在聊天区域滚动
        pyautogui.scroll(random.randint(-5, 5), 960, 600)
    
    def _action_switch_chat(self, app_name):
        """切换聊天对象"""
        # 点击左侧聊天列表（大致位置）
        pyautogui.click(random.randint(100, 300), random.randint(200, 700))
        time.sleep(0.3)
    
    def _action_type_message(self, app_name):
        """输入消息"""
        messages = [
            "测试消息",
            "在吗",
            "收到",
            "好的",
            "哈哈",
            "[表情]",
            "1",
        ]
        msg = random.choice(messages)
        pyautogui.typewrite(msg, interval=0.01)
        time.sleep(0.2)
        if random.random() > 0.3:  # 70%概率发送
            pyautogui.press('enter')
    
    def _action_click_sticker(self, app_name):
        """点击表情/贴纸按钮"""
        # 点击表情按钮（大致位置）
        pyautogui.click(random.randint(400, 600), random.randint(800, 850))
        time.sleep(0.3)
        # 随机选择一个表情
        pyautogui.click(random.randint(500, 900), random.randint(600, 750))
        time.sleep(0.2)
    
    def _action_idle(self, app_name):
        """空闲等待"""
        time.sleep(random.uniform(1, 3))
    
    def test_app(self, app_config, duration_per_app=60):
        """测试单个应用"""
        app_name, process_names, launch_cmd = app_config
        
        print(f"\n  [Chat] 检查 {app_name}...")
        
        # 检查进程是否在运行
        pid = self.find_process(process_names)
        if not pid:
            print(f"  [Chat] {app_name} 未运行，跳过")
            self.skipped_apps.append(app_name)
            return False
        
        print(f"  [Chat] {app_name} 进程已找到 (PID: {pid})")
        
        # 查找窗口
        windows = self.find_window(process_names)
        if not windows:
            print(f"  [Chat] {app_name} 窗口未找到，尝试备用关键词...")
            # 尝试用应用名作为窗口标题关键词
            windows = self.find_window([app_name])
        
        if not windows:
            print(f"  [Chat] {app_name} 窗口查找失败，跳过")
            self.skipped_apps.append(app_name)
            return False
        
        # 激活主窗口（通常是第一个）
        hwnd, title = windows[0]
        print(f"  [Chat] 找到窗口: {title}")
        
        if not self.activate_window(hwnd, title):
            self.skipped_apps.append(app_name)
            return False
        
        print(f"  [Chat] {app_name} 窗口已激活")
        time.sleep(0.5)
        
        # 模拟聊天活动
        self.simulate_chat_activity(app_name, duration_per_app)
        self.tested_apps.append(app_name)
        return True
    
    def run_test(self, round_num, apps_to_test=None, duration_per_app=60):
        """运行社交软件测试"""
        if apps_to_test is None:
            apps_to_test = ["QQ", "微信"]
        
        self.tested_apps = []
        self.skipped_apps = []
        
        print(f"[{time.strftime('%H:%M:%S')}] 开始第{round_num}轮 社交软件测试...")
        print(f"[Chat] 目标应用: {', '.join(apps_to_test)}")
        
        # 查找配置
        app_configs = []
        for app_name in apps_to_test:
            for config in self.SUPPORTED_APPS:
                if config[0] == app_name:
                    app_configs.append(config)
                    break
        
        if not app_configs:
            print("[Chat] 没有配置的应用需要测试")
            return
        
        monitor.start()
        try:
            for config in app_configs:
                self.test_app(config, duration_per_app)
                time.sleep(2)  # 应用间切换间隔
            
            # 打印测试摘要
            print(f"\n[Chat] 测试完成: 成功 {len(self.tested_apps)} 个, 跳过 {len(self.skipped_apps)} 个")
            if self.tested_apps:
                print(f"[Chat] 已测试: {', '.join(self.tested_apps)}")
            if self.skipped_apps:
                print(f"[Chat] 已跳过: {', '.join(self.skipped_apps)}")
                
        finally:
            monitor.stop("Chat", round_num)


def test_chat(round_num):
    """社交软件测试入口"""
    if not CONFIG["ENABLE_CHAT"]:
        return
    
    tester = ChatAppTester()
    # 每轮每个应用测试60秒
    tester.run_test(round_num, apps_to_test=["QQ", "微信"], duration_per_app=60)

def pre_login():
    """预登录流程：在正式测试前让用户登录所有网站"""
    if os.path.exists(STORAGE_STATE_FILE):
        print("检测到已保存的登录状态，是否重新登录？")
        choice = input("输入 'y' 重新登录，直接回车跳过: ").strip().lower()
        if choice != 'y':
            print("使用已保存的登录状态继续...")
            return
    
    print("\n=== 预登录模式 ===")
    print("将依次打开所有测试网站，请手动登录各账号")
    print("登录完成后请按回车键，脚本将保存登录状态\n")
    
    all_urls = WEB_TEST_URLS + [VIDEO_URL]
    
    browser = None
    context = None
    
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False, channel="msedge")
            context = browser.new_context()
            
            # 打开所有网页（限制最多10个标签页避免崩溃）
            max_tabs = 10
            pages = []
            for i, url in enumerate(all_urls[:max_tabs]):
                page = context.new_page()
                try:
                    page.goto(url, wait_until='domcontentloaded', timeout=30000)
                except:
                    pass  # 忽略加载超时
                pages.append(page)
                print(f"已打开 ({i+1}/{min(len(all_urls), max_tabs)}): {url}")
                time.sleep(0.5)
            
            if len(all_urls) > max_tabs:
                print(f"\n注：还有 {len(all_urls) - max_tabs} 个网站未打开，登录状态通常可共享")
            
            print("\n请在所有网页中完成登录...")
            print("登录完成后请在此窗口按回车键继续")
            
            # 使用input等待用户，而不是检测浏览器状态
            try:
                input()
            except KeyboardInterrupt:
                pass
            
            # 保存登录状态
            try:
                context.storage_state(path=STORAGE_STATE_FILE)
                print(f"\n登录状态已保存到: {STORAGE_STATE_FILE}")
            except Exception as e:
                print(f"\n保存登录状态失败: {e}")
            
            try:
                browser.close()
            except:
                pass
    
    except Exception as e:
        print(f"预登录过程出错: {e}")
    
    print("预登录完成，准备开始正式测试...\n")
    input("按回车键开始续航测试...")

def run_one_round(round_num):
    """运行一轮完整的测试，返回该轮耗时（秒）"""
    round_start = time.time()
    
    if CONFIG["ENABLE_OFFICE"]:
        test_office(round_num)
    
    if CONFIG["ENABLE_WEB"]:
        test_web(round_num)
    
    if CONFIG["ENABLE_VIDEO"]:
        test_video(round_num)
    
    if CONFIG["ENABLE_CHAT"]:
        test_chat(round_num)
    
    return time.time() - round_start

def select_test_modules():
    """交互式选择测试模块"""
    print("\n" + "="*60)
    print("              测试模块配置")
    print("="*60)
    print("请选择要启用的测试模块（输入 y/n）：\n")
    
    modules = [
        ("ENABLE_OFFICE", "Office生产力测试", "需安装UL Procyon"),
        ("ENABLE_WEB", "网页浏览测试", "京东/淘宝/IT之家等"),
        ("ENABLE_VIDEO", "视频播放测试", "Bilibili指定链接"),
        ("ENABLE_CHAT", "社交软件模拟", "QQ/微信窗口激活"),
        ("ENABLE_MONITOR", "实时性能监控", "HWiNFO/psutil数据"),
    ]
    
    for key, name, desc in modules:
        current = "开启" if CONFIG[key] else "关闭"
        choice = input(f"  {name} ({desc}) [{current}]: ").strip().lower()
        
        if choice == 'y':
            CONFIG[key] = True
            print(f"    → 已开启")
        elif choice == 'n':
            CONFIG[key] = False
            print(f"    → 已关闭")
        else:
            status = "开启" if CONFIG[key] else "关闭"
            print(f"    → 保持{status}")
    
    # 显示最终配置
    print("\n" + "="*60)
    print("              最终配置")
    print("="*60)
    for key, name, _ in modules:
        status = "✓ 开启" if CONFIG[key] else "✗ 关闭"
        print(f"  {name}: {status}")
    print("="*60)
    
    confirm = input("\n确认开始测试？(y/n): ").strip().lower()
    return confirm == 'y'


def main():
    print("=== 笔记本续航与性能自动化测试系统 ===")
    
    # 选择测试模块
    if not select_test_modules():
        print("\n已取消测试")
        return
    
    # 预登录流程（仅在网页测试开启且首次或用户选择时执行）
    if CONFIG["ENABLE_WEB"] or CONFIG["ENABLE_VIDEO"]:
        pre_login()
    
    round_num = 1
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r") as f:
            round_num = int(f.read().strip() or 1)
            print(f"恢复进度: 第 {round_num} 轮")
    
    # 启动电池追踪
    battery_tracker.start_test()
    
    # 启动性能监控（如果启用）
    if CONFIG["ENABLE_MONITOR"]:
        hwinfo_reader = HWiNFOReader()
    else:
        hwinfo_reader = None

    while True:
        curr_battery = psutil.sensors_battery().percent
        if curr_battery < 10:
            print(f"\n电量剩余 {curr_battery}%，达到阈值，准备休眠。")
            battery_tracker.generate_summary(round_num - 1, terminated_by_battery=True)
            if os.path.exists(PROGRESS_FILE): os.remove(PROGRESS_FILE)
            subprocess.run(["rundll32.exe", "powrprof.dll,SetSuspendState", "0,1,0"])
            break
        
        try:
            round_duration = run_one_round(round_num)
            
            # 记录本轮电池状态，传入hwinfo_reader获取准确功耗
            battery_tracker.record_round(round_num, round_duration, hwinfo_reader)
            
            round_num += 1
            with open(PROGRESS_FILE, "w") as f: f.write(str(round_num))
            time.sleep(10)
        except KeyboardInterrupt:
            print("\n\n用户中断测试...")
            battery_tracker.generate_summary(round_num - 1)
            break
        except Exception:
            traceback.print_exc()
            battery_tracker.generate_summary(round_num - 1)
            input("遇到错误，排查后按回车继续...")
            break
    
    print(f"\n详细数据已保存到: {BATTERY_LOG}")

if __name__ == "__main__":
    main()