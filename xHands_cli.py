#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xHands CLI - 工作流自动化工具命令行版本
"""

import cv2
import numpy as np
import pyautogui
import time
import json
import os
import sys
import argparse
import subprocess
from datetime import datetime
import threading
import signal

import base64
def _d(s):
    return base64.b64decode(s).decode('utf-8')
_A = _d('5oqW6Z+zQOm7keWMlueahOi2hee6p+WltueIuA==')
_V = _d('5bel5L2c5rWB6Ieq5Yqo5YyW5bel5YW3IENMSSB2MS4w')

class Colors:
    HEADER = '\033[95m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    MAGENTA = '\033[95m'
    BOLD = '\033[1m'
    DIM = '\033[2m'
    RESET = '\033[0m'
    
    @classmethod
    def disable(cls):
        cls.HEADER = ''
        cls.BLUE = ''
        cls.CYAN = ''
        cls.GREEN = ''
        cls.YELLOW = ''
        cls.RED = ''
        cls.MAGENTA = ''
        cls.BOLD = ''
        cls.DIM = ''
        cls.RESET = ''

if os.name == 'nt':
    import ctypes
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)

def print_banner():
    banner = f"""
{Colors.CYAN}{Colors.BOLD}╔══════════════════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                               
║   {Colors.YELLOW}██╗   ██╗{Colors.CYAN} ██╗   ██╗{Colors.GREEN}  █████╗{Colors.BLUE} ███╗   ██╗{Colors.MAGENTA} ██████╗{Colors.RED}  ███████╗{Colors.CYAN}                    
║   {Colors.YELLOW}╚██╗ ██╔╝{Colors.CYAN} ██║   ██║{Colors.GREEN} ██╔══██╗{Colors.BLUE} ████╗  ██║{Colors.MAGENTA} ██╔══██╗{Colors.RED}  ██╔════╝{Colors.CYAN}                    
║   {Colors.YELLOW} ╚████╔╝ {Colors.CYAN} ████████║{Colors.GREEN} ███████║{Colors.BLUE} ██╔██╗ ██║{Colors.MAGENTA} ██║  ██║{Colors.RED}  ███████╗{Colors.CYAN}                    
║   {Colors.YELLOW} ██╔██╗  {Colors.CYAN} ██╔═══██║{Colors.GREEN} ██╔══██║{Colors.BLUE} ██║╚██╗██║{Colors.MAGENTA} ██║  ██║{Colors.RED}  ╚════██║{Colors.CYAN}                    
║   {Colors.YELLOW}██╔╝ ██╗ {Colors.CYAN} ██║   ██║{Colors.GREEN} ██║  ██║{Colors.BLUE} ██║ ╚████║{Colors.MAGENTA} ██████╔╝{Colors.RED}  ███████║{Colors.CYAN}                    
║   {Colors.YELLOW}╚═╝  ╚═╝ {Colors.CYAN} ╚═╝   ╚═╝{Colors.GREEN} ╚═╝  ╚═╝{Colors.BLUE} ╚═╝  ╚═══╝{Colors.MAGENTA} ╚═════╝ {Colors.RED}  ╚══════╝{Colors.CYAN}                    
║                                                                                                
║   {Colors.GREEN}{_V}{Colors.CYAN}                                                     
║   {Colors.DIM}{_A}{Colors.CYAN}                                                         
║                                                                                                
╚════════════════════════════════════════════════════════════════════════════════════════════════╝{Colors.RESET}
"""
    print(banner)

current_dir = os.path.dirname(os.path.abspath(__file__))
config_file = os.path.join(current_dir, "config.json")

DEFAULT_FLOW_SCRIPTS_DIR = os.path.join(current_dir, "flowScripts")
DEFAULT_SCREEN_SHOT_DIR = os.path.join(current_dir, "screenShot")

def load_config():
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config
        except Exception as e:
            print(f"加载配置文件失败: {str(e)}")
    return {}

def save_config(config):
    try:
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存配置文件失败: {str(e)}")

config = load_config()

flow_scripts_dir = config.get('flow_scripts_dir', DEFAULT_FLOW_SCRIPTS_DIR)
screen_shot_dir = config.get('screen_shot_dir', DEFAULT_SCREEN_SHOT_DIR)
flow_py_dir = os.path.join(current_dir, "flowPY")

if not os.path.isabs(flow_scripts_dir):
    flow_scripts_dir = os.path.join(current_dir, flow_scripts_dir)
if not os.path.isabs(screen_shot_dir):
    screen_shot_dir = os.path.join(current_dir, screen_shot_dir)

os.makedirs(flow_scripts_dir, exist_ok=True)
os.makedirs(screen_shot_dir, exist_ok=True)
os.makedirs(flow_py_dir, exist_ok=True)

try:
    import pyperclip
except ImportError:
    import subprocess
    print("正在安装pyperclip...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyperclip"])
    print("pyperclip安装完成")
    import pyperclip

try:
    import pytesseract
    from PIL import Image
except ImportError:
    import subprocess
    print("正在安装OCR依赖...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pytesseract", "Pillow"])
    print("OCR依赖安装完成")
    import pytesseract
    from PIL import Image

stop_requested = False

def signal_handler(sig, frame):
    global stop_requested
    print("\n收到停止信号，正在停止执行...")
    stop_requested = True

signal.signal(signal.SIGINT, signal_handler)

def find_image(image_path, accuracy=0.9):
    screen = pyautogui.screenshot()
    screen_gray = cv2.cvtColor(np.array(screen), cv2.COLOR_RGB2GRAY)

    processed_path = image_path
    
    if "D:/Xhand/screenShot/" in processed_path or "D:/xWin/xTools/xHands/screenShot/" in processed_path:
        filename = os.path.basename(processed_path)
        processed_path = os.path.join(screen_shot_dir, filename)
    elif not os.path.isabs(processed_path):
        filename = os.path.basename(processed_path)
        processed_path = os.path.join(screen_shot_dir, filename)
    
    processed_path = os.path.normpath(processed_path)

    if not os.path.isfile(processed_path):
        raise Exception(f"模板文件不存在: {processed_path}")

    template = cv2.imread(processed_path, 0)
    if template is None:
        raise Exception(f"无法读取模板图像: {processed_path}")

    result = cv2.matchTemplate(screen_gray, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

    threshold = accuracy
    if max_val >= threshold:
        h, w = template.shape
        center_x = max_loc[0] + w // 2
        center_y = max_loc[1] + h // 2
        return (center_x, center_y)
    else:
        return None

def find_text(text):
    screen = pyautogui.screenshot()
    try:
        extracted_text = pytesseract.image_to_string(screen)
        return text in extracted_text
    except Exception as e:
        print(f"OCR识别失败: {str(e)}")
        return False

def log(message, level="info"):
    timestamp = datetime.now().strftime("%H:%M:%S")
    if level == "success":
        print(f"{Colors.GREEN}[{timestamp}] ✓ {message}{Colors.RESET}")
    elif level == "error":
        print(f"{Colors.RED}[{timestamp}] ✗ {message}{Colors.RESET}")
    elif level == "warning":
        print(f"{Colors.YELLOW}[{timestamp}] ⚠ {message}{Colors.RESET}")
    else:
        print(f"{Colors.CYAN}[{timestamp}] {message}{Colors.RESET}")

def launch_gui():
    if hasattr(sys, 'frozen'):
        print(f"{Colors.GREEN}正在启动 xHands GUI...{Colors.RESET}")
        import tkinter as tk
        from xHands import WorkflowAutomationTool
        root = tk.Tk()
        app = WorkflowAutomationTool(root)
        root.mainloop()
        return True
    
    gui_path = os.path.join(current_dir, "Hands.py")
    if not os.path.exists(gui_path):
        gui_path = os.path.join(current_dir, "xHands.py")
    if not os.path.exists(gui_path):
        print(f"{Colors.RED}错误: 找不到 GUI 程序文件{Colors.RESET}")
        return False
    
    print(f"{Colors.GREEN}正在启动 xHands GUI...{Colors.RESET}")
    subprocess.Popen([sys.executable, gui_path])
    print(f"{Colors.GREEN}GUI 已在后台启动{Colors.RESET}")
    return True

def execute_py_script(script_file, count=1, duration=0, interval=0):
    script_path = os.path.join(flow_py_dir, script_file)
    if not os.path.exists(script_path):
        print(f"错误: 脚本文件不存在: {script_file}")
        return False
    
    cmd_args = [sys.executable, script_path]
    if count != 1:
        cmd_args.extend(['-c', str(count)])
    if duration > 0:
        cmd_args.extend(['-d', str(duration)])
    if interval > 0:
        cmd_args.extend(['-i', str(interval)])
    
    log(f"执行脚本: {script_file}")
    log(f"参数: 次数={count}, 时长={duration}分钟, 间隔={interval}秒")
    
    result = subprocess.run(cmd_args)
    
    if result.returncode == 0:
        log("脚本执行完成")
        return True
    else:
        log(f"脚本执行失败，返回码: {result.returncode}")
        return False

def list_workflows():
    print(f"\n{Colors.BOLD}{Colors.CYAN}可用的工作流:{Colors.RESET}")
    print(f"{Colors.DIM}{'─' * 50}{Colors.RESET}")
    workflows = []
    for file in os.listdir(flow_scripts_dir):
        if file.endswith(".json"):
            try:
                with open(os.path.join(flow_scripts_dir, file), "r", encoding="utf-8") as f:
                    data = json.load(f)
                    name = data.get('name', file.replace('.json', ''))
                    steps_count = len(data.get('steps', []))
                    workflows.append((file, name, steps_count))
                    print(f"  {Colors.GREEN}{file:<25}{Colors.RESET} {Colors.YELLOW}名称:{Colors.RESET} {name:<15} {Colors.YELLOW}步骤:{Colors.RESET} {steps_count}")
            except:
                pass
    print(f"{Colors.DIM}{'─' * 50}{Colors.RESET}")
    if not workflows:
        print(f"  {Colors.DIM}(无工作流文件){Colors.RESET}")
    return workflows

def show_workflow_details(workflow_file):
    file_path = os.path.join(flow_scripts_dir, workflow_file)
    if not os.path.exists(file_path):
        print(f"{Colors.RED}错误: 工作流文件不存在: {workflow_file}{Colors.RESET}")
        return
    
    with open(file_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    print(f"\n{Colors.BOLD}{Colors.CYAN}工作流: {data.get('name', workflow_file)}{Colors.RESET}")
    print(f"{Colors.DIM}{'═' * 50}{Colors.RESET}")
    print(f"  {Colors.YELLOW}创建时间:{Colors.RESET} {data.get('created_at', '未知')}")
    print(f"  {Colors.YELLOW}步骤数:{Colors.RESET}   {len(data.get('steps', []))}")
    print(f"\n{Colors.BOLD}步骤详情:{Colors.RESET}")
    print(f"{Colors.DIM}{'─' * 50}{Colors.RESET}")
    
    for i, step in enumerate(data.get('steps', [])):
        step_type = step.get('type', 'unknown')
        type_map = {
            'mouse': '鼠标点击',
            'keyboard': '键盘按键',
            'delay': '延迟',
            'clipboard_write': '写入剪贴板',
            'clipboard_read': '读取剪贴板',
            'condition': '条件判断',
            'system_command': '系统命令',
            'doubao_screenshot': '豆包截图'
        }
        type_display = type_map.get(step_type, step_type)
        
        details = ""
        if step_type == "mouse":
            image_path = step.get('image_path', '')
            details = f"图像: {os.path.basename(image_path) if image_path else '无'}"
        elif step_type == "keyboard":
            keys = step.get('keys', '')
            details = f"按键: {keys}"
        elif step_type == "delay":
            duration = step.get('duration', 1)
            details = f"时长: {duration}秒"
        elif step_type == "clipboard_write":
            text = step.get('text', '')
            details = f"内容: {text[:30]}..." if len(text) > 30 else f"内容: {text}"
        elif step_type == "condition":
            image_path = step.get('image_path', '')
            details = f"图像: {os.path.basename(image_path) if image_path else '无'}"
        elif step_type == "system_command":
            cmd_type = step.get('cmd_type', 'launch_app')
            if cmd_type == "launch_app":
                app_path = step.get('app_path', '')
                details = f"启动: {os.path.basename(app_path) if app_path else '无'}"
            elif cmd_type == "set_window":
                window_title = step.get('window_title', '')
                window_x = step.get('window_x', 0)
                window_y = step.get('window_y', 0)
                window_width = step.get('window_width', 800)
                window_height = step.get('window_height', 600)
                details = f"固定窗体: ({window_x},{window_y}) {window_width}x{window_height}"
        elif step_type == "doubao_screenshot":
            details = "执行豆包截图功能"
        
        print(f"  {Colors.GREEN}{i+1:2}.{Colors.RESET} [{Colors.CYAN}{type_display:<8}{Colors.RESET}] {step.get('name', '未命名'):<20} {Colors.DIM}{details}{Colors.RESET}")
    
    print(f"{Colors.DIM}{'─' * 50}{Colors.RESET}")

def execute_workflow(workflow_file, count=1, duration=0, interval=0, retry=False, verbose=True):
    global stop_requested
    
    file_path = os.path.join(flow_scripts_dir, workflow_file)
    if not os.path.exists(file_path):
        print(f"错误: 工作流文件不存在: {workflow_file}")
        return False
    
    script_name = workflow_file.replace('.json', '.py')
    script_path = os.path.join(flow_py_dir, script_name)
    
    if os.path.exists(script_path):
        log(f"发现已导出的脚本: {script_name}")
        return execute_py_script(script_name, count, duration, interval)
    
    with open(file_path, "r", encoding="utf-8") as f:
        workflow_data = json.load(f)
    
    execution_count = count
    execution_duration = duration * 60
    execution_interval = interval
    retry_on_error = retry
    
    log(f"开始执行工作流: {workflow_data['name']}")
    log(f"执行设置: 次数={execution_count}, 时长={duration}分钟, 间隔={execution_interval}秒, 错误重试={retry_on_error}")
    
    exec_count = 0
    start_time = time.time()
    
    while True:
        if stop_requested:
            log("收到停止请求，正在停止执行...")
            break
        
        if execution_count > 0 and exec_count >= execution_count:
            log(f"已达到执行次数限制: {exec_count}")
            break
        
        if execution_duration > 0 and time.time() - start_time > execution_duration:
            log(f"已达到执行时长限制: {duration}分钟")
            break
        
        log(f"执行第 {exec_count + 1} 轮")
        success = True
        
        try:
            i = 0
            steps = workflow_data['steps']
            while i < len(steps):
                if stop_requested:
                    log("收到停止请求，正在停止执行...")
                    success = False
                    break
                
                step = steps[i]
                if verbose:
                    log(f"执行步骤 {i + 1}: {step['name']}")
                
                if step['type'] == "mouse":
                    target_type = step.get('target_type', 'image')
                    click_type = step.get('click_type', 'left')
                    accuracy = step.get('accuracy', 0.9)
                    click_type_names = {"left": "单击左键", "right": "单击右键", "double": "双击左键"}
                    
                    if target_type == "image":
                        image_path = step.get('image_path', '')
                        position = find_image(image_path, accuracy)
                        if position:
                            if verbose:
                                log(f"找到图像，位置: {position}")
                            if click_type == "left":
                                pyautogui.click(position)
                            elif click_type == "right":
                                pyautogui.rightClick(position)
                            elif click_type == "double":
                                pyautogui.doubleClick(position)
                            if verbose:
                                log(f"{click_type_names.get(click_type, '单击左键')}成功")
                        else:
                            log(f"未找到图像: {os.path.basename(image_path)}")
                            if not retry_on_error:
                                return False
                            success = False
                            break
                    else:
                        text = step.get('text', '')
                        if verbose:
                            log(f"查找文字: '{text}'")
                        if click_type == "left":
                            pyautogui.click()
                        elif click_type == "right":
                            pyautogui.rightClick()
                        elif click_type == "double":
                            pyautogui.doubleClick()
                        if verbose:
                            log(f"执行{click_type_names.get(click_type, '单击左键')}: '{text}'")
                        
                elif step['type'] == "system_command":
                    cmd_type = step.get('cmd_type', 'launch_app')
                    if cmd_type == "launch_app":
                        app_path = step.get('app_path', '')
                        if verbose:
                            log(f"启动应用: {os.path.basename(app_path)}")
                        try:
                            os.startfile(app_path)
                            if verbose:
                                log(f"应用启动成功: {os.path.basename(app_path)}")
                        except Exception as e:
                            log(f"应用启动失败: {str(e)}")
                            if not retry_on_error:
                                return False
                            success = False
                            break
                    elif cmd_type == "set_window":
                        window_title = step.get('window_title', '')
                        window_class = step.get('window_class', '')
                        window_x = step.get('window_x', 0)
                        window_y = step.get('window_y', 0)
                        window_width = step.get('window_width', 800)
                        window_height = step.get('window_height', 600)
                        
                        if verbose:
                            log(f"固定窗体: 标题='{window_title}', 类名='{window_class}'")
                        
                        try:
                            import win32gui
                            result_hwnd = [None]
                            
                            def enum_callback(hwnd, _):
                                if win32gui.IsWindowVisible(hwnd):
                                    win_title = win32gui.GetWindowText(hwnd)
                                    win_class = win32gui.GetClassName(hwnd)
                                    
                                    title_match = not window_title or window_title.lower() in win_title.lower()
                                    class_match = not window_class or window_class.lower() == win_class.lower()
                                    
                                    if title_match and class_match and win_title:
                                        result_hwnd[0] = hwnd
                                        return False
                                return True
                            
                            win32gui.EnumWindows(enum_callback, None)
                            hwnd = result_hwnd[0]
                            
                            if hwnd:
                                win32gui.SetWindowPos(hwnd, None, window_x, window_y, window_width, window_height, 0)
                                win32gui.SetForegroundWindow(hwnd)
                                if verbose:
                                    log(f"窗口已移动到 ({window_x}, {window_y}), 大小 {window_width}x{window_height}")
                            else:
                                log("未找到指定窗口")
                                if not retry_on_error:
                                    return False
                                success = False
                                break
                        except Exception as e:
                            log(f"固定窗体失败: {str(e)}")
                            if not retry_on_error:
                                return False
                            success = False
                            break
                        
                elif step['type'] == "keyboard":
                    trigger_type = step.get('trigger_type', 'none')
                    keys = step.get('keys', '')
                    
                    log(f"准备执行按键: {keys}")
                    try:
                        key_parts = keys.split('+')
                        normalized_keys = []
                        for key in key_parts:
                            key_lower = key.lower()
                            normalized_keys.append(key_lower)
                        
                        log(f"标准化按键: {normalized_keys}")
                        
                        if trigger_type == "none":
                            pyautogui.hotkey(*normalized_keys)
                            log(f"按键组合执行成功: {keys}")
                        else:
                            text = step.get('text', '')
                            if verbose:
                                log(f"查找文字: '{text}' 以触发按键操作")
                            pyautogui.hotkey(*normalized_keys)
                            log(f"按键组合执行成功: {keys}")
                    except Exception as e:
                        log(f"按键执行失败: {keys}, 错误: {str(e)}")
                        if not retry_on_error:
                            return False
                        success = False
                        break
                            
                elif step['type'] == "delay":
                    duration_sec = step.get('duration', 1)
                    if verbose:
                        log(f"执行延迟操作: {duration_sec}秒")
                    remaining = duration_sec
                    while remaining > 0 and not stop_requested:
                        sleep_time = min(0.5, remaining)
                        time.sleep(sleep_time)
                        remaining -= sleep_time
                    if stop_requested:
                        success = False
                        break
                        
                elif step['type'] == "clipboard_write":
                    text = step.get('text', '')
                    if verbose:
                        log(f"执行写入剪贴板操作: {text[:20]}..." if len(text) > 20 else f"执行写入剪贴板操作: {text}")
                    try:
                        pyperclip.copy(text)
                        if verbose:
                            log("写入剪贴板成功")
                    except Exception as e:
                        log(f"写入剪贴板失败: {str(e)}")
                        if not retry_on_error:
                            return False
                        success = False
                        break
                        
                elif step['type'] == "clipboard_read":
                    try:
                        clipboard_text = pyperclip.paste()
                        if verbose:
                            log(f"执行读取剪贴板操作: {clipboard_text[:20]}..." if len(clipboard_text) > 20 else f"执行读取剪贴板操作: {clipboard_text}")
                    except Exception as e:
                        log(f"读取剪贴板失败: {str(e)}")
                        if not retry_on_error:
                            return False
                        success = False
                        break
                        
                elif step['type'] == "doubao_screenshot":
                    log("执行豆包截图...")
                    try:
                        time.sleep(0.5)
                        pyautogui.hotkey('shift', 'alt', 'a')
                        time.sleep(0.5)
                        log("豆包截图执行成功")
                    except Exception as e:
                        log(f"豆包截图执行失败: {str(e)}")
                        if not retry_on_error:
                            return False
                        success = False
                        break
                        
                elif step['type'] == "condition":
                    condition_type = step.get('condition_type', 'image')
                    accuracy = step.get('accuracy', 0.9)
                    found_behavior = step.get('found_behavior', 'continue')
                    not_found_behavior = step.get('not_found_behavior', 'jump')
                    jump_to = step.get('jump_to', 1)
                    
                    found = False
                    if condition_type == "image":
                        image_path = step.get('image_path', '')
                        if verbose:
                            log(f"执行条件判断: 等待图像 {os.path.basename(image_path)} (精度: {int(accuracy * 100)}%)")
                        position = find_image(image_path, accuracy)
                        if position:
                            if verbose:
                                log(f"找到图像，位置: {position}")
                            found = True
                        else:
                            log(f"未找到图像 {os.path.basename(image_path)}")
                    else:
                        text = step.get('text', '')
                        if verbose:
                            log(f"执行条件判断: 等待文字 '{text}'")
                        found = find_text(text)
                        if found:
                            if verbose:
                                log(f"找到文字: '{text}'")
                        else:
                            log(f"未找到文字: '{text}'")
                    
                    if found:
                        if found_behavior == "continue":
                            if verbose:
                                log("条件满足，继续执行")
                        elif found_behavior == "jump":
                            log(f"跳转到步骤 {jump_to}")
                            i = jump_to - 2
                            if i < -1:
                                i = -1
                        elif found_behavior == "restart":
                            log("重新开始工作流")
                            break
                    else:
                        if not_found_behavior == "jump":
                            log(f"跳转到步骤 {jump_to}")
                            i = jump_to - 2
                            if i < -1:
                                i = -1
                        elif not_found_behavior == "wait":
                            log("继续找，直到找到为止...")
                            wait_start = time.time()
                            timeout = 60
                            while not found:
                                if stop_requested:
                                    log("收到停止请求...")
                                    success = False
                                    break
                                
                                if condition_type == "image":
                                    image_path = step.get('image_path', '')
                                    position = find_image(image_path, accuracy)
                                    if position:
                                        log(f"找到图像，位置: {position}")
                                        found = True
                                else:
                                    text = step.get('text', '')
                                    found = find_text(text)
                                    if found:
                                        log(f"找到文字: '{text}'")
                                
                                if found:
                                    log("条件满足，继续执行")
                                else:
                                    if time.time() - wait_start > timeout:
                                        log(f"超时: 未找到目标，超过 {timeout} 秒")
                                        if not retry_on_error:
                                            return False
                                        success = False
                                        break
                                    remaining = 2
                                    while remaining > 0 and not stop_requested:
                                        sleep_time = min(0.5, remaining)
                                        time.sleep(sleep_time)
                                        remaining -= sleep_time
                
                remaining = 1
                while remaining > 0 and not stop_requested:
                    sleep_time = min(0.5, remaining)
                    time.sleep(sleep_time)
                    remaining -= sleep_time
                
                i += 1
                
        except Exception as e:
            log(f"执行过程中出现错误: {str(e)}")
            if not retry_on_error:
                return False
            success = False
        
        if success:
            exec_count += 1
            
            if execution_interval > 0 and not stop_requested:
                log(f"等待 {execution_interval} 秒后继续...")
                remaining = execution_interval
                while remaining > 0 and not stop_requested:
                    sleep_time = min(0.5, remaining)
                    time.sleep(sleep_time)
                    remaining -= sleep_time
    
    log(f"工作流执行完成! 共执行 {exec_count} 轮")
    return True

def show_config():
    print(f"\n{Colors.BOLD}{Colors.CYAN}当前配置:{Colors.RESET}")
    print(f"{Colors.DIM}{'─' * 50}{Colors.RESET}")
    print(f"  {Colors.YELLOW}工作流目录:{Colors.RESET} {flow_scripts_dir}")
    print(f"  {Colors.YELLOW}找图目录:{Colors.RESET}   {screen_shot_dir}")
    print(f"  {Colors.YELLOW}脚本目录:{Colors.RESET}   {flow_py_dir}")
    print(f"{Colors.DIM}{'─' * 50}{Colors.RESET}")

def interactive_mode():
    global stop_requested
    
    print_banner()
    print(f"{Colors.BOLD}可用命令:{Colors.RESET}")
    print(f"  {Colors.GREEN}gui{Colors.RESET}              - 启动图形界面")
    print(f"  {Colors.GREEN}list, ls, l{Colors.RESET}     - 列出所有工作流")
    print(f"  {Colors.GREEN}show <name>{Colors.RESET}    - 显示工作流详情")
    print(f"  {Colors.GREEN}run <name>{Colors.RESET}     - 执行工作流")
    print(f"  {Colors.GREEN}exec <script>{Colors.RESET}  - 执行脚本")
    print(f"  {Colors.GREEN}config{Colors.RESET}         - 显示当前配置")
    print(f"  {Colors.GREEN}help, h, ?{Colors.RESET}     - 显示帮助")
    print(f"  {Colors.GREEN}quit, q, exit{Colors.RESET}  - 退出程序")
    print(f"{Colors.DIM}{'─' * 50}{Colors.RESET}")
    
    while True:
        try:
            user_input = input(f"\n{Colors.BOLD}{Colors.CYAN}xHands>{Colors.RESET} ").strip()
        except EOFError:
            break
        
        if not user_input:
            continue
        
        parts = user_input.split()
        cmd = parts[0].lower()
        args = parts[1:] if len(parts) > 1 else []
        
        if cmd in ['quit', 'q', 'exit']:
            print(f"{Colors.YELLOW}再见!{Colors.RESET}")
            break
        elif cmd == 'gui':
            launch_gui()
        elif cmd in ['list', 'ls', 'l']:
            list_workflows()
        elif cmd == 'show' and args:
            show_workflow_details(args[0])
        elif cmd == 'run' and args:
            stop_requested = False
            count = 1
            duration = 0
            interval = 0
            retry = False
            
            i = 1
            while i < len(args):
                if args[i] == '-c' and i + 1 < len(args):
                    count = int(args[i + 1])
                    i += 2
                elif args[i] == '-d' and i + 1 < len(args):
                    duration = int(args[i + 1])
                    i += 2
                elif args[i] == '-i' and i + 1 < len(args):
                    interval = int(args[i + 1])
                    i += 2
                elif args[i] == '-r':
                    retry = True
                    i += 1
                else:
                    i += 1
            
            execute_workflow(args[0], count, duration, interval, retry)
        elif cmd == 'exec' and args:
            count = 1
            duration = 0
            interval = 0
            
            i = 1
            while i < len(args):
                if args[i] == '-c' and i + 1 < len(args):
                    count = int(args[i + 1])
                    i += 2
                elif args[i] == '-d' and i + 1 < len(args):
                    duration = int(args[i + 1])
                    i += 2
                elif args[i] == '-i' and i + 1 < len(args):
                    interval = int(args[i + 1])
                    i += 2
                else:
                    i += 1
            
            execute_py_script(args[0], count, duration, interval)
        elif cmd == 'config':
            show_config()
        elif cmd in ['help', 'h', '?']:
            print(f"\n{Colors.BOLD}{Colors.CYAN}可用命令:{Colors.RESET}")
            print(f"  {Colors.GREEN}gui{Colors.RESET}                      - 启动图形界面")
            print(f"  {Colors.GREEN}list, ls, l{Colors.RESET}           - 列出所有工作流")
            print(f"  {Colors.GREEN}show <name>{Colors.RESET}          - 显示工作流详情")
            print(f"  {Colors.GREEN}run <name> [options]{Colors.RESET} - 执行工作流")
            print(f"    {Colors.DIM}选项:{Colors.RESET}")
            print(f"      {Colors.YELLOW}-c <次数>{Colors.RESET}   执行次数 (默认1, 0=无限)")
            print(f"      {Colors.YELLOW}-d <分钟>{Colors.RESET}   执行时长 (默认0, 0=不限制)")
            print(f"      {Colors.YELLOW}-i <秒>{Colors.RESET}     执行间隔 (默认0)")
            print(f"      {Colors.YELLOW}-r{Colors.RESET}          错误时重试")
            print(f"  {Colors.GREEN}exec <script.py> [options]{Colors.RESET} - 执行flowPY目录中的脚本")
            print(f"    {Colors.DIM}选项:{Colors.RESET}")
            print(f"      {Colors.YELLOW}-c <次数>{Colors.RESET}   执行次数 (默认1, 0=无限)")
            print(f"      {Colors.YELLOW}-d <分钟>{Colors.RESET}   执行时长 (默认0, 0=不限制)")
            print(f"      {Colors.YELLOW}-i <秒>{Colors.RESET}     执行间隔 (默认0)")
            print(f"  {Colors.GREEN}config{Colors.RESET}                 - 显示当前配置")
            print(f"  {Colors.GREEN}help, h, ?{Colors.RESET}           - 显示帮助")
            print(f"  {Colors.GREEN}quit, q, exit{Colors.RESET}        - 退出程序")
            print(f"\n{Colors.DIM}{_A}{Colors.RESET}")
        else:
            print(f"{Colors.RED}未知命令: {cmd}。输入 'help' 查看帮助。{Colors.RESET}")

def main():
    parser = argparse.ArgumentParser(
        description='xHands CLI - 工作流自动化工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
示例:
  python xHands_cli.py                    进入交互模式
  python xHands_cli.py gui                启动图形界面
  python xHands_cli.py list               列出所有工作流
  python xHands_cli.py show test.json     显示工作流详情
  python xHands_cli.py run test.json      执行工作流
  python xHands_cli.py run test.json -c 5 -d 10 -i 2 -r
                                          执行5次，时长10分钟，间隔2秒，错误重试
  python xHands_cli.py exec script.py     执行flowPY目录中的脚本
  python xHands_cli.py exec script.py -c 5 -d 10 -i 2
                                          执行脚本，5次，时长10分钟，间隔2秒
  
{_A}
        '''
    )
    
    parser.add_argument('command', nargs='?', default=None,
                        help='命令: gui/list/show/run/exec/config')
    parser.add_argument('target', nargs='?', default=None,
                        help='目标文件名 (用于show/run/exec)')
    parser.add_argument('-c', '--count', type=int, default=1,
                        help='执行次数 (0=无限循环)')
    parser.add_argument('-d', '--duration', type=int, default=0,
                        help='执行时长(分钟, 0=不限制)')
    parser.add_argument('-i', '--interval', type=int, default=0,
                        help='执行间隔(秒)')
    parser.add_argument('-r', '--retry', action='store_true',
                        help='出现错误时重试')
    parser.add_argument('-q', '--quiet', action='store_true',
                        help='静默模式，减少输出')
    
    args = parser.parse_args()
    
    if args.command is None:
        interactive_mode()
    elif args.command.lower() == 'gui':
        launch_gui()
    elif args.command.lower() in ['list', 'ls', 'l']:
        list_workflows()
    elif args.command.lower() == 'show':
        if args.target:
            show_workflow_details(args.target)
        else:
            print(f"{Colors.RED}错误: 请指定工作流文件名{Colors.RESET}")
            print("用法: python xHands_cli.py show <workflow.json>")
    elif args.command.lower() == 'run':
        if args.target:
            execute_workflow(
                args.target,
                count=args.count,
                duration=args.duration,
                interval=args.interval,
                retry=args.retry,
                verbose=not args.quiet
            )
        else:
            print(f"{Colors.RED}错误: 请指定工作流文件名{Colors.RESET}")
            print("用法: python xHands_cli.py run <workflow.json> [options]")
    elif args.command.lower() == 'exec':
        if args.target:
            execute_py_script(
                args.target,
                count=args.count,
                duration=args.duration,
                interval=args.interval
            )
        else:
            print(f"{Colors.RED}错误: 请指定脚本文件名{Colors.RESET}")
            print("用法: python xHands_cli.py exec <script.py> [options]")
    elif args.command.lower() == 'config':
        show_config()
    else:
        print(f"{Colors.RED}未知命令: {args.command}{Colors.RESET}")
        print(f"可用命令: {Colors.GREEN}gui, list, show, run, exec, config{Colors.RESET}")
        print("或直接运行进入交互模式")

if __name__ == "__main__":
    main()
