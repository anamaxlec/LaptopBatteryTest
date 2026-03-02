import os
import sys
import time
import subprocess
import psutil
import traceback
from playwright.sync_api import sync_playwright
import pyautogui
import win32gui
import win32con

if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

def get_battery():
    battery = psutil.sensors_battery()
    return battery.percent if battery else 100

def focus_window_by_title(title_keyword):
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
        except Exception as e:
            print(f"抢占焦点失败，尝试模拟Alt键后重试:{e}")
            pyautogui.press('alt')
            time.sleep(0.5)
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception as inner_e:
                print(f"最终获取焦点失败:{inner_e}")
                return False
        time.sleep(1) 
        return True
    return False

def test_office_procyon(round_num):
    print(f"进行第{round_num}轮办公测试…………")
    
    def_file = os.path.join(base_dir, "office_productivity.def")
    result_file = os.path.join(base_dir, f"round_{round_num}.procyon-result")
    csv_file = os.path.join(base_dir, f"round_{round_num}.csv")
    recovered_file = os.path.join(base_dir, f"recovered_round_{round_num}.procyon-result")
    
    procyon_path = r"C:\Program Files\UL\Procyon\ProcyonCmd.exe"
    if not os.path.exists(procyon_path):
        raise Exception("未能找到ULProcyon主程序，请检查绝对路径是否正确")
    
    cmd = [
        procyon_path, 
        f"--definition={def_file}", 
        f"--out={result_file}",
        f"--export-csv={csv_file}"
    ]
    
    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        print(f"Procyon测试异常中断，尝试恢复数据:{e}")
        recovery_cmd = [
            procyon_path,
            "--recovery",
            f"--out={recovered_file}"
        ]
        subprocess.run(recovery_cmd)
        raise Exception("办公测试执行失败")

def test_web_browsing():
    print("进行实际网页浏览测试……………")
    urls = [
        "https://www.jd.com", "https://www.taobao.com", "https://www.sina.com.cn",
        "https://www.163.com", "https://www.sohu.com", "https://www.ithome.com",
        "https://www.chiphell.com", "https://bbs.nga.cn", "https://www.gamersky.com",
        "https://www.3dmgame.com", "https://www.4399.com", "https://weibo.com",
        "https://www.zhihu.com"
    ]
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="msedge")
        context = browser.new_context()
        pages = []
        
        for url in urls:
            page = context.new_page()
            page.goto(url)
            pages.append(page)
        
        for _ in range(5):
            for page in pages:
                page.bring_to_front()
                page.reload()
                time.sleep(2)
                for _ in range(5):
                    page.mouse.wheel(0, 500)
                    time.sleep(1)
        
        browser.close()

def test_video_playback():
    print("进行视频浏览测试……………")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="msedge")
        page = browser.new_page()
        page.goto("https://v.qq.com/") 
        time.sleep(690)
        browser.close()

def test_chat_software():
    chat_apps = {"QQ": "QQ", "WeChat": "微信"}
    for app_name, window_title in chat_apps.items():
        print(f"进行{app_name}聊天测试……………")
        if focus_window_by_title(window_title):
            for i in range(25):
                pyautogui.typewrite(f"Test message {i}")
                pyautogui.press('enter')
                time.sleep(12) 
        else:
            print(f"未找到{app_name}的窗口，跳过本次测试。")

def run_test_cycle(round_num):
    test_office_procyon(round_num)
    time.sleep(5)
    test_web_browsing()
    time.sleep(5)
    test_video_playback()
    time.sleep(5)
    test_chat_software()
    time.sleep(5)

def main():
    print("测试已开始!")
    round_num = 1
    progress_file = os.path.join(base_dir, "test_progress.txt")
    
    if os.path.exists(progress_file):
        with open(progress_file, "r") as f:
            content = f.read().strip()
            if content.isdigit():
                round_num = int(content)
                print(f"检测到崩溃恢复记录，从第{round_num}轮继续测试。")
            
    while True:
        battery = get_battery()
        print(f"剩余电量百分比:{battery}%")
        
        if battery < 10:
            print("电池电量已低于10%，结束测试!")
            print("60秒后将使电脑进入休眠/睡眠模式——")
            time.sleep(60)
            subprocess.run(["rundll32.exe", "powrprof.dll,SetSuspendState", "0,1,0"])
            if os.path.exists(progress_file):
                os.remove(progress_file)
            break
            
        print(f"开始第{round_num}轮循环测试…………")
        
        try:
            run_test_cycle(round_num)
            round_num += 1
            with open(progress_file, "w") as f:
                f.write(str(round_num))
        except Exception as e:
            print("测试脚本发生异常崩溃，详细错误信息如下:")
            traceback.print_exc()
            input("\n按回车键关闭窗口...")
            break

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("程序发生致命级全局错误:")
        traceback.print_exc()
        input("\n按回车键关闭窗口...")