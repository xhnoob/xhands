#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工作流: douyinpinglun
创建时间: 2026-02-21T03:16:49.726893
导出时间: 2026-02-21 03:24:30
"""

import cv2
import numpy as np
import pyautogui
import time
import os
import sys

try:
    import pyperclip
except ImportError:
    print('正在安装pyperclip...')
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyperclip'])
    import pyperclip

try:
    import win32gui
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

SCREEN_SHOT_DIR = r'./screenShot'

def find_image(image_path, accuracy=0.9):
    screen = pyautogui.screenshot()
    screen_gray = cv2.cvtColor(np.array(screen), cv2.COLOR_RGB2GRAY)
    
    processed_path = image_path
    if not os.path.isabs(processed_path):
        processed_path = os.path.join(SCREEN_SHOT_DIR, os.path.basename(processed_path))
    
    processed_path = os.path.normpath(processed_path)
    if not os.path.isfile(processed_path):
        print(f'模板文件不存在: {processed_path}')
        return None
    
    template = cv2.imread(processed_path, 0)
    if template is None:
        print(f'无法读取模板图像: {processed_path}')
        return None
    
    result = cv2.matchTemplate(screen_gray, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
    
    if max_val >= accuracy:
        h, w = template.shape
        center_x = max_loc[0] + w // 2
        center_y = max_loc[1] + h // 2
        print(f'找到图像，位置: ({center_x}, {center_y})，置信度: {max_val:.2f}')
        return (center_x, center_y)
    else:
        print(f'未找到图像: {os.path.basename(image_path)}，置信度: {max_val:.2f}')
        return None

def find_window(title, class_name):
    if not HAS_WIN32:
        return None
    result_hwnd = [None]
    
    def enum_callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            win_title = win32gui.GetWindowText(hwnd)
            win_class = win32gui.GetClassName(hwnd)
            title_match = not title or title.lower() in win_title.lower()
            class_match = not class_name or class_name.lower() == win_class.lower()
            if title_match and class_match and win_title:
                result_hwnd[0] = hwnd
                return False
        return True
    
    win32gui.EnumWindows(enum_callback, None)
    return result_hwnd[0]

def log(message):
    timestamp = time.strftime('%H:%M:%S')
    print(f'[{timestamp}] {message}')

def execute_workflow():
    steps = [
        {'id': 'action_11', 'name': '启动抖音应用', 'type': 'system_command', 'cmd_type': 'launch_app', 'app_path': 'YOUR_DOUYIN_PATH/douyin.exe'},
        {'id': 'delay_1771591317', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_14', 'name': '固定抖音窗体', 'type': 'system_command', 'cmd_type': 'set_window', 'window_title': '抖音', 'window_class': 'Chrome_WidgetWin_1', 'window_x': 97, 'window_y': 43, 'window_width': 960, 'window_height': 1283},
        {'id': 'delay_1771591270', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_16', 'name': '启动豆包应用', 'type': 'system_command', 'cmd_type': 'launch_app', 'app_path': 'YOUR_DOUBAO_PATH/Doubao.exe'},
        {'id': 'delay_1771591352', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_18', 'name': '固定豆包窗体', 'type': 'system_command', 'cmd_type': 'set_window', 'window_title': '豆包', 'window_class': 'Chrome_WidgetWin_1', 'window_x': 1535, 'window_y': 194, 'window_width': 978, 'window_height': 969},
        {'id': 'delay_1771591371', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_3', 'name': '下一个视频', 'type': 'mouse', 'image_path': 'screenShot/next.png'},
        {'id': 'delay_1771269823', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_40', 'name': '判断是否直播间', 'type': 'condition', 'condition_type': 'image', 'image_path': 'screenShot/panduanzhibojian.png', 'accuracy': 0.9, 'found_behavior': 'restart', 'not_found_behavior': 'jump', 'jump_to': 1},
        {'id': 'action_43', 'name': '判断是否广告', 'type': 'condition', 'condition_type': 'image', 'image_path': 'screenShot/panduanguanggao.png', 'accuracy': 0.9, 'found_behavior': 'restart', 'not_found_behavior': 'jump', 'jump_to': 1},
        {'id': 'delay_1771135497', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_26', 'name': '豆包截图', 'type': 'doubao_screenshot'},
        {'id': 'action_3', 'name': '下一个视频', 'type': 'mouse', 'image_path': 'screenShot/next.png'},
        {'id': 'delay_1771135507', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_9', 'name': '问问豆包', 'type': 'mouse', 'image_path': 'screenShot/wenwendoubao.png', 'accuracy': 0.9},
        {'id': 'delay_1771136027', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_25', 'name': '鼠标点击豆包发消息框', 'type': 'mouse', 'image_path': 'screenShot/faxxiaoxi.png', 'accuracy': 0.9},
        {'id': 'delay_1771136147', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_10', 'name': '豆包提示词', 'type': 'clipboard_write', 'text': '这是一张抖音某视频的评论区截图，请你根据图片上的所有信息，并参考别人的评论，帮我写一条可能成为热评的评论，给出的结果不要包含任何其他任何文字，只要单纯的结果就行，让我能直接复制这个结果'},
        {'id': 'delay_1771136201', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_14', 'name': '粘贴', 'type': 'keyboard', 'keys': 'ctrl+V'},
        {'id': 'delay_1771136208', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_41', 'name': '按下回车键', 'type': 'keyboard', 'trigger_type': 'none', 'keys': 'enter'},
        {'id': 'delay_1771137432', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_37', 'name': '判断是否思考完毕', 'type': 'condition', 'image_path': 'screenShot/douobaosikao.png', 'accuracy': 0.9, 'found_behavior': 'continue', 'not_found_behavior': 'wait', 'jump_to': 1},
        {'id': 'delay_1771137437', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_20', 'name': '鼠标点击豆包复制', 'type': 'mouse', 'image_path': 'screenShot/doubaofuzhi.png', 'accuracy': 0.9},
        {'id': 'delay_1771137489', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_22', 'name': '鼠标点击抖音留下评论', 'type': 'mouse', 'image_path': 'screenShot/liuxiapinglun.png', 'accuracy': 0.9},
        {'id': 'delay_1771137573', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_14', 'name': '粘贴', 'type': 'keyboard', 'keys': 'ctrl+V'},
        {'id': 'delay_1771137660', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
        {'id': 'action_27', 'name': '鼠标点击抖音发送评论', 'type': 'mouse', 'image_path': 'screenShot/douyinfasong.png', 'accuracy': 0.9},
        {'id': 'delay_1771137670', 'name': '延迟操作', 'type': 'delay', 'duration': 1},
    ]
    
    log('开始执行工作流')
    
    for i, step in enumerate(steps):
        log(f'执行步骤 {i + 1}: {step["name"]}')
        
        if step['type'] == 'mouse':
            target_type = step.get('target_type', 'image')
            click_type = step.get('click_type', 'left')
            accuracy = step.get('accuracy', 0.9)
            
            if target_type == 'image':
                image_path = step.get('image_path', '')
                position = find_image(image_path, accuracy)
                if position:
                    if click_type == 'left':
                        pyautogui.click(position)
                    elif click_type == 'right':
                        pyautogui.rightClick(position)
                    elif click_type == 'double':
                        pyautogui.doubleClick(position)
                    log('点击成功')
                else:
                    log('未找到图像，跳过')
            else:
                if click_type == 'left':
                    pyautogui.click()
                elif click_type == 'right':
                    pyautogui.rightClick()
                elif click_type == 'double':
                    pyautogui.doubleClick()
        
        elif step['type'] == 'keyboard':
            keys = step.get('keys', '')
            key_parts = keys.split('+')
            normalized_keys = [k.lower() for k in key_parts]
            pyautogui.hotkey(*normalized_keys)
            log(f'按键执行成功: {keys}')
        
        elif step['type'] == 'system_command':
            cmd_type = step.get('cmd_type', 'launch_app')
            if cmd_type == 'launch_app':
                app_path = step.get('app_path', '')
                os.startfile(app_path)
                log(f'启动应用: {os.path.basename(app_path)}')
            elif cmd_type == 'set_window':
                if HAS_WIN32:
                    window_title = step.get('window_title', '')
                    window_class = step.get('window_class', '')
                    window_x = step.get('window_x', 0)
                    window_y = step.get('window_y', 0)
                    window_width = step.get('window_width', 800)
                    window_height = step.get('window_height', 600)
                    hwnd = find_window(window_title, window_class)
                    if hwnd:
                        win32gui.SetWindowPos(hwnd, None, window_x, window_y, window_width, window_height, 0)
                        win32gui.SetForegroundWindow(hwnd)
                        log('固定窗体成功')
        
        elif step['type'] == 'delay':
            duration = step.get('duration', 1)
            time.sleep(duration)
            log(f'延迟 {duration} 秒')
        
        elif step['type'] == 'clipboard_write':
            text = step.get('text', '')
            pyperclip.copy(text)
            log('写入剪贴板成功')
        
        elif step['type'] == 'clipboard_read':
            clipboard_text = pyperclip.paste()
            log(f'读取剪贴板: {clipboard_text[:30]}...' if len(clipboard_text) > 30 else f'读取剪贴板: {clipboard_text}')
        
        elif step['type'] == 'doubao_screenshot':
            pyautogui.hotkey('shift', 'alt', 'a')
            log('豆包截图执行成功')
        
        time.sleep(0.5)
    
    log('工作流执行完成')

if __name__ == '__main__':
    import sys
    print('=' * 50)
    print('工作流执行脚本')
    print('=' * 50)
    
    # 解析参数
    count = 1  # 默认执行次数
    interval = 0  # 默认间隔
    
    i = 1
    while i < len(sys.argv):
        if sys.argv[i] == '-c' and i + 1 < len(sys.argv):
            count = int(sys.argv[i + 1])
            i += 2
        elif sys.argv[i] == '-i' and i + 1 < len(sys.argv):
            interval = int(sys.argv[i + 1])
            i += 2
        elif sys.argv[i] == '--auto':
            i += 1
        else:
            i += 1
    
    # 支持命令行参数跳过输入
    if len(sys.argv) > 1 and sys.argv[1] == '--auto':
        # 自动模式
        for c in range(count):
            print(f'执行第 {c+1}/{count} 次')
            execute_workflow()
            if c < count - 1 and interval > 0:
                print(f'等待 {interval} 秒...')
                time.sleep(interval)
    else:
        input('按回车键开始执行...')
        execute_workflow()
        input('按回车键退出...')
