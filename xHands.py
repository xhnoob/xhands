import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import cv2
import numpy as np
import pyautogui
import time
import json
import os
import sys
from datetime import datetime
import base64

def _d(s):
    return base64.b64decode(s).decode('utf-8')
_A = _d('5oqW6Z+zQOm7keWMlueahOi2hee6p+WltueIuA==')
_T = _d('eGhhbmRzIC0g5oqW6Z+zQOm7keWMlueahOi2hee6p+WltueIuA==')

try:
    import win32gui
    import win32process
    import psutil
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# 获取当前脚本所在目录的绝对路径
current_dir = os.path.dirname(os.path.abspath(__file__))

# 配置文件路径
config_file = os.path.join(current_dir, "config.json")

# 默认目录
DEFAULT_FLOW_SCRIPTS_DIR = os.path.join(current_dir, "flowScripts")
DEFAULT_SCREEN_SHOT_DIR = os.path.join(current_dir, "screenShot")

# 加载配置
def load_config():
    """加载配置文件"""
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config
        except Exception as e:
            print(f"加载配置文件失败: {str(e)}")
    return {}

# 保存配置
def save_config(config):
    """保存配置文件"""
    try:
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存配置文件失败: {str(e)}")

# 加载配置
config = load_config()

# 读取目录配置
flow_scripts_dir = config.get('flow_scripts_dir', DEFAULT_FLOW_SCRIPTS_DIR)
screen_shot_dir = config.get('screen_shot_dir', DEFAULT_SCREEN_SHOT_DIR)
flow_py_dir = os.path.join(current_dir, "flowPY")

# 处理相对路径
if not os.path.isabs(flow_scripts_dir):
    flow_scripts_dir = os.path.join(current_dir, flow_scripts_dir)
if not os.path.isabs(screen_shot_dir):
    screen_shot_dir = os.path.join(current_dir, screen_shot_dir)

# 确保目录存在
os.makedirs(flow_scripts_dir, exist_ok=True)
os.makedirs(screen_shot_dir, exist_ok=True)
os.makedirs(flow_py_dir, exist_ok=True)

# 尝试导入pyperclip，如果没有则安装
try:
    import pyperclip
except ImportError:
    import subprocess
    import sys
    print("正在安装pyperclip...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyperclip"])
    print("pyperclip安装完成")
    import pyperclip

# 尝试导入pytesseract和PIL，如果没有则安装
try:
    import pytesseract
    from PIL import Image
except ImportError:
    import subprocess
    import sys
    print("正在安装OCR依赖...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pytesseract", "Pillow"])
    print("OCR依赖安装完成")
    import pytesseract
    from PIL import Image

class WorkflowAutomationTool:
    def __init__(self, root):
        self.root = root
        self.root.title(_T)
        self.root.geometry("800x750")
        self.root.resizable(True, True)

        # 创建操作列表
        self.actions = []
        self.workflow = []
        self.action_counter = 1
        
        # 执行控制标志
        self.is_executing = False
        self.stop_requested = False

        # 创建界面
        self.create_widgets()

        # 加载保存的操作
        self.load_actions()
        
        # 添加读取剪贴板预设操作
        self.add_clipboard_read_preset()
        
        # 添加按下回车键预设操作
        self.add_enter_key_preset()
        
        # 添加豆包截图预设操作
        self.add_doubao_screenshot_preset()
        
        # 添加快捷键绑定
        self.root.bind('<F8>', lambda event: self.execute_workflow())
        self.root.bind('<F9>', lambda event: self.stop_execution())

    def create_widgets(self):
        # 创建标签页
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 操作管理标签页
        action_frame = ttk.Frame(notebook)
        notebook.add(action_frame, text="操作管理")

        # 工作流编辑标签页
        workflow_frame = ttk.Frame(notebook)
        notebook.add(workflow_frame, text="工作流编辑")

        # 工作流执行标签页
        execution_frame = ttk.Frame(notebook)
        notebook.add(execution_frame, text="工作流执行")

        # 设置标签页
        settings_frame = ttk.Frame(notebook)
        notebook.add(settings_frame, text="设置")

        # 相关标签页
        about_frame = ttk.Frame(notebook)
        notebook.add(about_frame, text="相关")

        # 操作管理界面
        self.create_action_management_ui(action_frame)

        # 工作流编辑界面
        self.create_workflow_editor_ui(workflow_frame)

        # 工作流执行界面
        self.create_execution_ui(execution_frame)

        # 设置界面
        self.create_settings_ui(settings_frame)

        # 相关界面
        self.create_about_ui(about_frame)

    def create_action_management_ui(self, parent):
        # 操作类型选择 - 改为选项卡形式
        action_type_frame = ttk.LabelFrame(parent, text="操作类型")
        action_type_frame.pack(fill=tk.X, padx=10, pady=10)

        self.action_notebook = ttk.Notebook(action_type_frame)
        self.action_notebook.pack(fill=tk.BOTH, expand=True)

        # 鼠标点击选项卡
        mouse_tab = ttk.Frame(self.action_notebook)
        self.action_notebook.add(mouse_tab, text="鼠标点击")

        # 定位方式选择
        ttk.Label(mouse_tab, text="定位方式:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.mouse_target_var = tk.StringVar(value="image")
        mouse_target_frame = ttk.Frame(mouse_tab)
        mouse_target_frame.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(mouse_target_frame, text="图像识别", variable=self.mouse_target_var, value="image").pack(side=tk.TOP, anchor=tk.W)
        ttk.Radiobutton(mouse_target_frame, text="文字识别", variable=self.mouse_target_var, value="text").pack(side=tk.TOP, anchor=tk.W)

        # 点击类型选择
        ttk.Label(mouse_tab, text="点击类型:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.click_type_var = tk.StringVar(value="left")
        click_type_frame = ttk.Frame(mouse_tab)
        click_type_frame.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(click_type_frame, text="单击左键", variable=self.click_type_var, value="left").pack(side=tk.LEFT, anchor=tk.W)
        ttk.Radiobutton(click_type_frame, text="单击右键", variable=self.click_type_var, value="right").pack(side=tk.LEFT, anchor=tk.W, padx=10)
        ttk.Radiobutton(click_type_frame, text="双击左键", variable=self.click_type_var, value="double").pack(side=tk.LEFT, anchor=tk.W, padx=10)

        # 图像识别选项
        ttk.Label(mouse_tab, text="截图路径:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.image_path_var = tk.StringVar()
        ttk.Entry(mouse_tab, textvariable=self.image_path_var, width=50).grid(row=2, column=1, padx=5, pady=5)
        screenshot_btn_frame = ttk.Frame(mouse_tab)
        screenshot_btn_frame.grid(row=2, column=2, padx=5, pady=5)
        ttk.Button(screenshot_btn_frame, text="浏览", command=self.browse_image).pack(side=tk.LEFT)
        ttk.Button(screenshot_btn_frame, text="截图", command=self.start_screenshot).pack(side=tk.LEFT, padx=5)

        # 文字识别选项
        ttk.Label(mouse_tab, text="目标文字:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.mouse_text_var = tk.StringVar(value="示例文字")
        ttk.Entry(mouse_tab, textvariable=self.mouse_text_var, width=50).grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(mouse_tab, text="操作名称:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        self.mouse_name_var = tk.StringVar(value="鼠标点击")
        ttk.Entry(mouse_tab, textvariable=self.mouse_name_var).grid(row=4, column=1, padx=5, pady=5)

        ttk.Label(mouse_tab, text="识别精度:").grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)
        self.accuracy_var = tk.DoubleVar(value=0.9)
        ttk.Scale(mouse_tab, from_=0.5, to=1.0, variable=self.accuracy_var, orient=tk.HORIZONTAL).grid(row=5, column=1, padx=5, pady=5, sticky=tk.W)
        self.accuracy_label = ttk.Label(mouse_tab, text="90%")
        self.accuracy_label.grid(row=5, column=2, padx=5, pady=5, sticky=tk.W)

        # 绑定精度变化事件
        self.accuracy_var.trace_add("write", self.update_accuracy_label)

        # 键盘按键选项卡
        keyboard_tab = ttk.Frame(self.action_notebook)
        self.action_notebook.add(keyboard_tab, text="键盘按键")

        # 触发条件选择
        ttk.Label(keyboard_tab, text="触发条件:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.keyboard_trigger_var = tk.StringVar(value="none")
        keyboard_trigger_frame = ttk.Frame(keyboard_tab)
        keyboard_trigger_frame.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(keyboard_trigger_frame, text="直接执行", variable=self.keyboard_trigger_var, value="none").pack(side=tk.TOP, anchor=tk.W)
        ttk.Radiobutton(keyboard_trigger_frame, text="文字识别", variable=self.keyboard_trigger_var, value="text").pack(side=tk.TOP, anchor=tk.W)

        # 文字识别选项
        ttk.Label(keyboard_tab, text="目标文字:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.keyboard_text_var = tk.StringVar(value="示例文字")
        ttk.Entry(keyboard_tab, textvariable=self.keyboard_text_var, width=50).grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(keyboard_tab, text="按键组合:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.keys_var = tk.StringVar(value="ctrl+c")
        ttk.Entry(keyboard_tab, textvariable=self.keys_var).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(keyboard_tab, text="读取键盘操作", command=self.read_keyboard_input).grid(row=2, column=2, padx=5, pady=5)

        ttk.Label(keyboard_tab, text="操作名称:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.keyboard_name_var = tk.StringVar(value="键盘按键")
        ttk.Entry(keyboard_tab, textvariable=self.keyboard_name_var).grid(row=3, column=1, padx=5, pady=5)

        # 系统命令选项卡
        system_tab = ttk.Frame(self.action_notebook)
        self.action_notebook.add(system_tab, text="系统命令")

        # 命令类型选择
        ttk.Label(system_tab, text="命令类型:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.system_cmd_type_var = tk.StringVar(value="launch_app")
        system_cmd_frame = ttk.Frame(system_tab)
        system_cmd_frame.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(system_cmd_frame, text="启动应用", variable=self.system_cmd_type_var, value="launch_app", command=self.on_system_cmd_type_change).pack(side=tk.TOP, anchor=tk.W)
        ttk.Radiobutton(system_cmd_frame, text="固定窗体", variable=self.system_cmd_type_var, value="set_window", command=self.on_system_cmd_type_change).pack(side=tk.TOP, anchor=tk.W)

        # 启动应用选项
        self.launch_app_frame = ttk.Frame(system_tab)
        self.launch_app_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(self.launch_app_frame, text="应用路径:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.app_path_var = tk.StringVar()
        ttk.Entry(self.launch_app_frame, textvariable=self.app_path_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        app_btn_frame = ttk.Frame(self.launch_app_frame)
        app_btn_frame.grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(app_btn_frame, text="浏览", command=self.browse_app).pack(side=tk.LEFT)
        ttk.Button(app_btn_frame, text="窗体工具", command=self.open_window_tool).pack(side=tk.LEFT, padx=5)

        # 固定窗体选项
        self.set_window_frame = ttk.Frame(system_tab)
        
        ttk.Label(self.set_window_frame, text="窗口标题:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.window_title_var = tk.StringVar()
        ttk.Entry(self.set_window_frame, textvariable=self.window_title_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(self.set_window_frame, text="窗口类名:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.window_class_var = tk.StringVar()
        ttk.Entry(self.set_window_frame, textvariable=self.window_class_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        
        pos_frame = ttk.Frame(self.set_window_frame)
        pos_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(pos_frame, text="窗口位置:").pack(side=tk.LEFT)
        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT, padx=(10, 2))
        self.window_x_var = tk.StringVar(value="0")
        ttk.Entry(pos_frame, textvariable=self.window_x_var, width=8).pack(side=tk.LEFT)
        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT, padx=(10, 2))
        self.window_y_var = tk.StringVar(value="0")
        ttk.Entry(pos_frame, textvariable=self.window_y_var, width=8).pack(side=tk.LEFT)
        
        ttk.Label(pos_frame, text="  大小:").pack(side=tk.LEFT)
        ttk.Label(pos_frame, text="宽:").pack(side=tk.LEFT, padx=(10, 2))
        self.window_width_var = tk.StringVar(value="800")
        ttk.Entry(pos_frame, textvariable=self.window_width_var, width=8).pack(side=tk.LEFT)
        ttk.Label(pos_frame, text="高:").pack(side=tk.LEFT, padx=(10, 2))
        self.window_height_var = tk.StringVar(value="600")
        ttk.Entry(pos_frame, textvariable=self.window_height_var, width=8).pack(side=tk.LEFT)
        
        ttk.Button(self.set_window_frame, text="窗体工具", command=self.open_set_window_tool).grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)

        ttk.Label(system_tab, text="操作名称:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.system_cmd_name_var = tk.StringVar(value="启动应用")
        ttk.Entry(system_tab, textvariable=self.system_cmd_name_var).grid(row=2, column=1, padx=5, pady=5)
        
        self.on_system_cmd_type_change()

        # 剪贴板写入选项卡
        clipboard_write_tab = ttk.Frame(self.action_notebook)
        self.action_notebook.add(clipboard_write_tab, text="写入剪贴板")

        ttk.Label(clipboard_write_tab, text="文字内容:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.NW)
        self.clipboard_text_var = tk.StringVar(value="示例文字")
        self.clipboard_text = tk.Text(clipboard_write_tab, width=50, height=5)
        self.clipboard_text.grid(row=0, column=1, padx=5, pady=5)
        self.clipboard_text.insert(tk.END, "示例文字")

        ttk.Label(clipboard_write_tab, text="操作名称:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.clipboard_write_name_var = tk.StringVar(value="写入剪贴板")
        ttk.Entry(clipboard_write_tab, textvariable=self.clipboard_write_name_var).grid(row=1, column=1, padx=5, pady=5)

        # 条件判断选项卡
        condition_tab = ttk.Frame(self.action_notebook)
        self.action_notebook.add(condition_tab, text="条件判断")

        # 判断类型选择
        ttk.Label(condition_tab, text="判断类型:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.condition_type_var = tk.StringVar(value="image")
        condition_type_frame = ttk.Frame(condition_tab)
        condition_type_frame.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(condition_type_frame, text="图像识别", variable=self.condition_type_var, value="image").pack(side=tk.TOP, anchor=tk.W)
        ttk.Radiobutton(condition_type_frame, text="文字识别", variable=self.condition_type_var, value="text").pack(side=tk.TOP, anchor=tk.W)

        # 图像识别选项
        ttk.Label(condition_tab, text="判断图片:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.condition_image_var = tk.StringVar()
        ttk.Entry(condition_tab, textvariable=self.condition_image_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(condition_tab, text="浏览", command=self.browse_condition_image).grid(row=1, column=2, padx=5, pady=5)

        # 文字识别选项
        ttk.Label(condition_tab, text="判断文字:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.condition_text_var = tk.StringVar(value="示例文字")
        ttk.Entry(condition_tab, textvariable=self.condition_text_var, width=50).grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(condition_tab, text="操作名称:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.condition_name_var = tk.StringVar(value="条件判断")
        ttk.Entry(condition_tab, textvariable=self.condition_name_var).grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(condition_tab, text="识别精度:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        self.condition_accuracy_var = tk.DoubleVar(value=0.9)
        ttk.Scale(condition_tab, from_=0.5, to=1.0, variable=self.condition_accuracy_var, orient=tk.HORIZONTAL).grid(row=4, column=1, padx=5, pady=5, sticky=tk.W)
        self.condition_accuracy_label = ttk.Label(condition_tab, text="90%")
        self.condition_accuracy_label.grid(row=4, column=2, padx=5, pady=5, sticky=tk.W)

        ttk.Label(condition_tab, text="找到时的行为:").grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)
        self.condition_found_behavior_var = tk.StringVar(value="continue")
        found_behavior_frame = ttk.Frame(condition_tab)
        found_behavior_frame.grid(row=5, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(found_behavior_frame, text="继续执行", variable=self.condition_found_behavior_var, value="continue").pack(side=tk.TOP, anchor=tk.W)
        ttk.Radiobutton(found_behavior_frame, text="跳转到步骤", variable=self.condition_found_behavior_var, value="jump").pack(side=tk.TOP, anchor=tk.W)
        ttk.Radiobutton(found_behavior_frame, text="重新开始工作流", variable=self.condition_found_behavior_var, value="restart").pack(side=tk.TOP, anchor=tk.W)

        ttk.Label(condition_tab, text="未找到时的行为:").grid(row=6, column=0, padx=5, pady=5, sticky=tk.W)
        self.condition_not_found_behavior_var = tk.StringVar(value="jump")
        not_found_behavior_frame = ttk.Frame(condition_tab)
        not_found_behavior_frame.grid(row=6, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(not_found_behavior_frame, text="跳转到步骤", variable=self.condition_not_found_behavior_var, value="jump").pack(side=tk.TOP, anchor=tk.W)
        ttk.Radiobutton(not_found_behavior_frame, text="继续找，直到找到为止", variable=self.condition_not_found_behavior_var, value="wait").pack(side=tk.TOP, anchor=tk.W)

        ttk.Label(condition_tab, text="跳转步骤:").grid(row=7, column=0, padx=5, pady=5, sticky=tk.W)
        self.condition_jump_var = tk.IntVar(value=1)
        ttk.Entry(condition_tab, textvariable=self.condition_jump_var, width=10).grid(row=7, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Label(condition_tab, text="（从1开始）").grid(row=7, column=2, padx=5, pady=5, sticky=tk.W)

        # 绑定精度变化事件
        self.condition_accuracy_var.trace_add("write", self.update_condition_accuracy_label)

        # 操作配置
        config_frame = ttk.LabelFrame(parent, text="操作配置")
        config_frame.pack(fill=tk.X, padx=10, pady=10)

        # 操作管理按钮
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(button_frame, text="添加操作", command=self.add_action).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="测试操作", command=self.test_action).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="移除操作", command=self.remove_action).pack(side=tk.LEFT, padx=5)

        # 操作列表
        list_frame = ttk.LabelFrame(parent, text="操作列表")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.action_tree = ttk.Treeview(list_frame, columns=("id", "name", "type", "details"), show="headings")
        self.action_tree.heading("id", text="ID")
        self.action_tree.heading("name", text="名称")
        self.action_tree.heading("type", text="类型")
        self.action_tree.heading("details", text="详情")
        self.action_tree.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.action_tree, orient=tk.VERTICAL, command=self.action_tree.yview)
        self.action_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.action_tree_menu = tk.Menu(self.root, tearoff=0)
        self.action_tree_menu.add_command(label="测试操作", command=self.test_action_from_menu)
        self.action_tree_menu.add_command(label="移除操作", command=self.remove_action_from_menu)
        self.action_tree.bind("<ButtonRelease-3>", self.show_action_tree_menu)

    def create_workflow_editor_ui(self, parent):
        # 可用操作列表
        left_frame = ttk.LabelFrame(parent, text="可用操作")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10, ipadx=5)

        self.available_actions_tree = ttk.Treeview(left_frame, columns=("id", "name", "type"), show="headings")
        self.available_actions_tree.heading("id", text="ID")
        self.available_actions_tree.heading("name", text="名称")
        self.available_actions_tree.heading("type", text="类型")
        self.available_actions_tree.column("id", width=40)
        self.available_actions_tree.column("name", width=120)
        self.available_actions_tree.column("type", width=80)
        self.available_actions_tree.pack(fill=tk.BOTH, expand=True)
        
        self.action_context_menu = tk.Menu(self.root, tearoff=0)
        self.action_context_menu.add_command(label="添加到工作流", command=self.add_to_workflow)
        self.action_context_menu.add_command(label="替换工作流", command=self.replace_workflow_step)
        self.action_context_menu.add_command(label="插入工作流", command=self.insert_workflow_step)
        self.available_actions_tree.bind("<ButtonRelease-3>", self.show_action_context_menu)

        # 工作流编辑区域
        right_frame = ttk.LabelFrame(parent, text="工作流")
        self.workflow_frame = right_frame
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10, ipadx=5)

        self.workflow_tree = ttk.Treeview(right_frame, columns=("id", "name", "type"), show="headings")
        self.workflow_tree.heading("id", text="步骤")
        self.workflow_tree.heading("name", text="名称")
        self.workflow_tree.heading("type", text="类型")
        # 设置列宽
        self.workflow_tree.column("id", width=40)
        self.workflow_tree.column("name", width=150)
        self.workflow_tree.column("type", width=80)
        self.workflow_tree.pack(fill=tk.BOTH, expand=True)

        # 编辑按钮
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        button_grid = ttk.Frame(button_frame)
        button_grid.pack(fill=tk.X)

        ttk.Button(button_grid, text="添加到工作流", command=self.add_to_workflow).pack(fill=tk.X, pady=2)
        ttk.Button(button_grid, text="从工作流移除", command=self.remove_from_workflow).pack(fill=tk.X, pady=2)
        ttk.Button(button_grid, text="替换工作流", command=self.replace_workflow).pack(fill=tk.X, pady=2)
        ttk.Button(button_grid, text="插入到", command=self.insert_workflow).pack(fill=tk.X, pady=2)
        ttk.Button(button_grid, text="上移", command=self.move_up).pack(fill=tk.X, pady=2)
        ttk.Button(button_grid, text="下移", command=self.move_down).pack(fill=tk.X, pady=2)
        ttk.Button(button_grid, text="添加延迟", command=self.add_delay).pack(fill=tk.X, pady=2)
        ttk.Button(button_grid, text="插入延迟", command=self.insert_delay).pack(fill=tk.X, pady=2)
        ttk.Button(button_grid, text="保存工作流", command=self.save_workflow).pack(fill=tk.X, pady=2)
        ttk.Button(button_grid, text="加载工作流", command=self.load_workflow).pack(fill=tk.X, pady=2)
        ttk.Button(button_grid, text="清空工作流", command=self.clear_workflow).pack(fill=tk.X, pady=2)

    def create_execution_ui(self, parent):
        # 工作流选择
        workflow_frame = ttk.LabelFrame(parent, text="工作流选择")
        workflow_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(workflow_frame, text="选择工作流:").pack(side=tk.LEFT, padx=5)
        self.workflow_var = tk.StringVar()
        self.workflow_combobox = ttk.Combobox(workflow_frame, textvariable=self.workflow_var)
        self.workflow_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(workflow_frame, text="刷新", command=self.refresh_workflows).pack(side=tk.LEFT, padx=5)

        # 执行控制
        control_frame = ttk.LabelFrame(parent, text="执行控制")
        control_frame.pack(fill=tk.X, padx=10, pady=10)

        # 执行设置
        settings_frame = ttk.Frame(control_frame)
        settings_frame.pack(fill=tk.X, pady=5)

        ttk.Label(settings_frame, text="执行次数:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.execution_count_var = tk.IntVar(value=1)
        ttk.Entry(settings_frame, textvariable=self.execution_count_var, width=10).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Label(settings_frame, text="(0表示无限循环)").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)

        ttk.Label(settings_frame, text="执行时长:").grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        self.execution_duration_var = tk.IntVar(value=0)
        ttk.Entry(settings_frame, textvariable=self.execution_duration_var, width=10).grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        ttk.Label(settings_frame, text="(分钟，0表示不限制)").grid(row=0, column=5, padx=5, pady=5, sticky=tk.W)

        ttk.Label(settings_frame, text="执行间隔:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.execution_interval_var = tk.IntVar(value=0)
        ttk.Entry(settings_frame, textvariable=self.execution_interval_var, width=10).grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Label(settings_frame, text="(秒)").grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)

        # 错误处理选项
        error_handling_frame = ttk.Frame(control_frame)
        error_handling_frame.pack(fill=tk.X, pady=5)
        self.retry_on_error_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(error_handling_frame, text="出现任何错误，重新执行任务流，不用提示", variable=self.retry_on_error_var).pack(side=tk.LEFT, padx=5)

        # 执行按钮
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill=tk.X, pady=5)

        ttk.Button(button_frame, text="执行工作流", command=self.execute_workflow).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="停止执行", command=self.stop_execution).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="导出为脚本", command=self.export_workflow_as_script).pack(side=tk.LEFT, padx=5)
        ttk.Label(button_frame, text="快捷键: F8 执行, F9 停止").pack(side=tk.LEFT, padx=10, pady=5)

        # 执行日志
        log_frame = ttk.LabelFrame(parent, text="执行日志")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.log_text = tk.Text(log_frame, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(self.log_text, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def create_settings_ui(self, parent):
        dir_frame = ttk.LabelFrame(parent, text="目录设置")
        dir_frame.pack(fill=tk.X, padx=10, pady=10)

        screenshot_dir_frame = ttk.Frame(dir_frame)
        screenshot_dir_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(screenshot_dir_frame, text="找图目录:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.screenshot_dir_var = tk.StringVar(value=screen_shot_dir)
        ttk.Entry(screenshot_dir_frame, textvariable=self.screenshot_dir_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(screenshot_dir_frame, text="浏览", command=self.browse_screenshot_dir).grid(row=0, column=2, padx=5, pady=5)

        workflow_dir_frame = ttk.Frame(dir_frame)
        workflow_dir_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(workflow_dir_frame, text="工作流目录:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.workflow_dir_var = tk.StringVar(value=flow_scripts_dir)
        ttk.Entry(workflow_dir_frame, textvariable=self.workflow_dir_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(workflow_dir_frame, text="浏览", command=self.browse_workflow_dir).grid(row=0, column=2, padx=5, pady=5)

        system_frame = ttk.LabelFrame(parent, text="系统设置")
        system_frame.pack(fill=tk.X, padx=10, pady=10)

        system_btn_frame = ttk.Frame(system_frame)
        system_btn_frame.pack(fill=tk.X, pady=10)

        ttk.Button(system_btn_frame, text="安装 xHands 命令", command=self.install_xhands_command).pack(side=tk.LEFT, padx=10)
        ttk.Button(system_btn_frame, text="卸载 xHands 命令", command=self.uninstall_xhands_command).pack(side=tk.LEFT, padx=10)
        
        self.cmd_status_label = ttk.Label(system_btn_frame, text="")
        self.cmd_status_label.pack(side=tk.LEFT, padx=10)
        
        ttk.Button(system_btn_frame, text="安装必要Python库", command=self.install_required_packages).pack(side=tk.LEFT, padx=10)
        
        self.refresh_cmd_status()

        connection_frame = ttk.LabelFrame(parent, text="连接设置")
        connection_frame.pack(fill=tk.X, padx=10, pady=10)

        feishu_frame = ttk.Frame(connection_frame)
        feishu_frame.pack(fill=tk.X, pady=10, padx=10)

        ttk.Label(feishu_frame, text="飞书ID:").pack(side=tk.LEFT, padx=5)
        self.feishu_id_var = tk.StringVar()
        self.feishu_id_entry = ttk.Entry(feishu_frame, textvariable=self.feishu_id_var, width=50, foreground='gray')
        saved_feishu_id = self.load_feishu_id()
        if saved_feishu_id:
            self.feishu_id_var.set(saved_feishu_id)
            self.feishu_id_entry.config(foreground='black')
        else:
            self.feishu_id_var.set("问飞书机器人，我的飞书ID多少进行获取")
            self.feishu_id_entry.config(foreground='gray')
        self.feishu_id_entry.pack(side=tk.LEFT, padx=5)
        self.feishu_id_entry.bind("<FocusIn>", self.on_feishu_id_focus_in)
        self.feishu_id_entry.bind("<FocusOut>", self.on_feishu_id_focus_out)
        ttk.Button(feishu_frame, text="保存ID", command=self.save_feishu_id).pack(side=tk.LEFT, padx=5)
        ttk.Button(feishu_frame, text="连接测试", command=self.test_connection).pack(side=tk.LEFT, padx=5)

        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="保存设置", command=self.save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="恢复默认", command=self.restore_defaults).pack(side=tk.LEFT, padx=5)

    def create_about_ui(self, parent):
        about_frame = ttk.LabelFrame(parent, text="关于 xHands")
        about_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        about_text = tk.Text(about_frame, wrap=tk.WORD, font=("Microsoft YaHei", 11), padx=15, pady=15)
        about_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        about_content = """
欢迎使用 xHands - 工作流自动化工具

【基本使用说明】

0. 开箱必做
   - 管理权限启动exe
   - 点击安装xHands命令（重启terminal生效）
   - 配置飞书ID

1. 操作管理
   - 创建鼠标点击、键盘按键、延迟等操作
   - 支持图像识别和文字识别定位
   - 可将操作保存为预设，方便复用

2. 工作流编辑
   - 从左侧操作列表选择操作添加到工作流
   - 支持条件判断、循环执行
   - 可保存工作流为JSON文件

3. 工作流执行
   - 选择已保存的工作流执行
   - 支持设置执行次数、时长、间隔
   - F8 开始执行，F9 停止执行

4. CLI 命令行
   - xhands list          列出所有工作流
   - xhands run <name>    执行工作流
   - xhands gui           启动图形界面

【快捷键】
   F8 - 开始执行工作流
   F9 - 停止执行

【作者信息】
   抖音: @黑化的超级奶爸
   
   感谢使用本工具，如有问题欢迎反馈！
"""
        about_text.insert(tk.END, about_content)
        about_text.config(state=tk.DISABLED)
        
        author_frame = ttk.LabelFrame(parent, text="作者")
        author_frame.pack(fill=tk.X, padx=10, pady=10)
        
        author_label = ttk.Label(author_frame, text=_A, font=("Microsoft YaHei", 14, "bold"))
        author_label.pack(pady=15)

    def browse_image(self):
        file_path = filedialog.askopenfilename(
            initialdir=screen_shot_dir,
            filetypes=[("Image files", "*.png;*.jpg;*.jpeg")]
        )
        if file_path:
            relative_path = os.path.relpath(file_path, current_dir)
            self.image_path_var.set(relative_path)

    def browse_condition_image(self):
        file_path = filedialog.askopenfilename(
            initialdir=screen_shot_dir,
            filetypes=[("Image files", "*.png;*.jpg;*.jpeg")]
        )
        if file_path:
            relative_path = os.path.relpath(file_path, current_dir)
            self.condition_image_var.set(relative_path)

    def browse_app(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("可执行文件", "*.exe;*.bat;*.cmd"), ("所有文件", "*.*")]
        )
        if file_path:
            self.app_path_var.set(file_path)

    def start_screenshot(self):
        self.root.withdraw()
        self.root.after(200, self._do_screenshot)

    def _do_screenshot(self):
        ScreenshotSelector(self.root, self.image_path_var, screen_shot_dir)

    def open_window_tool(self):
        if not HAS_WIN32:
            messagebox.showerror("错误", "窗体工具需要安装 pywin32 和 psutil 库\n\n请运行以下命令安装:\npip install pywin32 psutil")
            return
        
        WindowToolDialog(self.root, self.app_path_var)

    def on_system_cmd_type_change(self):
        cmd_type = self.system_cmd_type_var.get()
        if cmd_type == "launch_app":
            self.launch_app_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=5, pady=5)
            self.set_window_frame.grid_forget()
            self.system_cmd_name_var.set("启动应用")
        elif cmd_type == "set_window":
            self.launch_app_frame.grid_forget()
            self.set_window_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=5, pady=5)
            self.system_cmd_name_var.set("固定窗体")

    def open_set_window_tool(self):
        if not HAS_WIN32:
            messagebox.showerror("错误", "窗体工具需要安装 pywin32 和 psutil 库\n\n请运行以下命令安装:\npip install pywin32 psutil")
            return
        
        SetWindowToolDialog(self.root, self.window_title_var, self.window_class_var, 
                           self.window_x_var, self.window_y_var, self.window_width_var, self.window_height_var)

    def find_window(self, title=None, class_name=None):
        if not HAS_WIN32:
            return None
        
        result = [None]
        
        def enum_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                win_title = win32gui.GetWindowText(hwnd)
                win_class = win32gui.GetClassName(hwnd)
                
                title_match = not title or title.lower() in win_title.lower()
                class_match = not class_name or class_name.lower() == win_class.lower()
                
                if title_match and class_match and win_title:
                    result[0] = hwnd
                    return False
            return True
        
        try:
            win32gui.EnumWindows(enum_callback, None)
        except:
            pass
        
        return result[0]

    def browse_screenshot_dir(self):
        """浏览找图目录"""
        dir_path = filedialog.askdirectory(
            initialdir=screen_shot_dir,
            title="选择找图目录"
        )
        if dir_path:
            self.screenshot_dir_var.set(dir_path)

    def browse_workflow_dir(self):
        dir_path = filedialog.askdirectory(
            initialdir=flow_scripts_dir,
            title="选择工作流目录"
        )
        if dir_path:
            self.workflow_dir_var.set(dir_path)

    def get_xhands_command_path(self):
        return current_dir

    def refresh_cmd_status(self):
        bin_dir = self.get_xhands_command_path()
        bat_file = os.path.join(bin_dir, 'xHands.bat')
        if os.path.exists(bat_file):
            user_path = os.environ.get('PATH', '')
            import winreg
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Environment', 0, winreg.KEY_READ)
                reg_path, _ = winreg.QueryValueEx(key, 'PATH')
                winreg.CloseKey(key)
                if bin_dir.lower() in reg_path.lower():
                    self.cmd_status_label.config(text="状态: 已安装", foreground="green")
                else:
                    self.cmd_status_label.config(text="状态: 未安装 (需添加到PATH)", foreground="orange")
            except:
                self.cmd_status_label.config(text="状态: 未安装", foreground="gray")
        else:
            self.cmd_status_label.config(text="状态: 未安装", foreground="gray")

    def install_required_packages(self):
        packages = [
            'opencv-python',
            'numpy',
            'pyautogui',
            'pyperclip',
            'pytesseract',
            'Pillow',
            'psutil',
            'pywin32'
        ]
        
        result = messagebox.askyesno(
            "安装确认",
            f"将安装以下Python库:\n\n{chr(10).join(packages)}\n\n是否继续?"
        )
        
        if not result:
            return
        
        import subprocess
        
        self.root.withdraw()
        
        print("正在安装必要的Python库...")
        print("=" * 50)
        
        failed = []
        for pkg in packages:
            print(f"\n正在安装 {pkg}...")
            try:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', pkg, '--quiet'])
                print(f"✓ {pkg} 安装成功")
            except subprocess.CalledProcessError as e:
                failed.append(pkg)
                print(f"✗ {pkg} 安装失败: {e}")
        
        print("\n" + "=" * 50)
        
        self.root.deiconify()
        
        if failed:
            messagebox.showwarning(
                "安装完成",
                f"部分库安装失败:\n{chr(10).join(failed)}\n\n请手动安装这些库。"
            )
        else:
            messagebox.showinfo(
                "安装完成",
                "所有必要的Python库已安装成功！"
            )

    def install_xhands_command(self):
        bin_dir = self.get_xhands_command_path()
        bat_file = os.path.join(bin_dir, 'xHands.bat')
        cli_path = os.path.join(bin_dir, 'xHands_cli.py')
        
        if not os.path.exists(cli_path):
            messagebox.showerror("错误", f"找不到 xHands_cli.py 文件:\n{cli_path}")
            return
        
        try:
            bat_content = '''@echo off
python "%~dp0xHands_cli.py" %*
'''
            try:
                with open(bat_file, 'w', encoding='utf-8') as f:
                    f.write(bat_content)
            except Exception as write_err:
                messagebox.showerror("错误", f"创建批处理文件失败: {str(write_err)}\n\n目标路径: {bat_file}")
                return
            
            if not os.path.exists(bat_file):
                messagebox.showerror("错误", f"批处理文件创建失败:\n{bat_file}")
                return
            
            import winreg
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Environment', 0, winreg.KEY_ALL_ACCESS)
                current_path, _ = winreg.QueryValueEx(key, 'PATH')
                if bin_dir.lower() not in current_path.lower():
                    new_path = current_path + ';' + bin_dir if current_path else bin_dir
                    winreg.SetValueEx(key, 'PATH', 0, winreg.REG_EXPAND_SZ, new_path)
                winreg.CloseKey(key)
            except Exception as e:
                messagebox.showerror("错误", f"添加到 PATH 失败: {str(e)}\n\n请手动将以下路径添加到系统 PATH:\n{bin_dir}")
                return
            
            self.refresh_cmd_status()
            messagebox.showinfo("成功", f"xHands 命令已安装！\n\n安装位置: {bat_file}\nPATH 已添加: {bin_dir}\n\n请重新打开终端窗口后使用 xHands 命令")
            
        except Exception as e:
            messagebox.showerror("错误", f"安装失败: {str(e)}")

    def uninstall_xhands_command(self):
        bin_dir = self.get_xhands_command_path()
        
        import winreg
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Environment', 0, winreg.KEY_ALL_ACCESS)
            current_path, _ = winreg.QueryValueEx(key, 'PATH')
            paths = current_path.split(';')
            new_paths = [p for p in paths if p.lower() != bin_dir.lower()]
            new_path = ';'.join(new_paths)
            winreg.SetValueEx(key, 'PATH', 0, winreg.REG_EXPAND_SZ, new_path)
            winreg.CloseKey(key)
        except:
            pass
        
        self.refresh_cmd_status()
        messagebox.showinfo("成功", f"xHands 命令已从 PATH 移除！\n\n请重新打开终端窗口以生效")

    def on_feishu_id_focus_in(self, event):
        if self.feishu_id_var.get() == "问飞书机器人，我的飞书ID多少进行获取":
            self.feishu_id_var.set("")
            self.feishu_id_entry.config(foreground='black')

    def on_feishu_id_focus_out(self, event):
        if not self.feishu_id_var.get():
            self.feishu_id_var.set("问飞书机器人，我的飞书ID多少进行获取")
            self.feishu_id_entry.config(foreground='gray')

    def load_feishu_id(self):
        config_file = os.path.join(current_dir, 'config.json')
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return config.get('feishu_id', '')
            except:
                pass
        return ''

    def save_feishu_id(self):
        feishu_id = self.feishu_id_var.get().strip()
        if feishu_id == "问飞书机器人，我的飞书ID多少进行获取":
            feishu_id = ""
        config_file = os.path.join(current_dir, 'config.json')
        config = {}
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            except:
                pass
        config['feishu_id'] = feishu_id
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        messagebox.showinfo("成功", "飞书ID 已保存！")

    def test_connection(self):
        feishu_id = self.feishu_id_var.get().strip()
        if not feishu_id or feishu_id == "问飞书机器人，我的飞书ID多少进行获取":
            messagebox.showwarning("提示", "请先输入飞书ID！")
            return
        try:
            import subprocess
            cmd = f'openclaw message send --channel feishu --account main --target "{feishu_id}" --message "恭喜你！连接很丝滑！"'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='utf-8')
            if result.returncode == 0:
                messagebox.showinfo("连接测试", f"连接成功！\n\n{result.stdout}")
            else:
                messagebox.showwarning("连接测试", f"连接返回非成功状态！\n\n{result.stderr}")
        except Exception as e:
            messagebox.showerror("连接测试", f"连接失败！\n\n错误: {str(e)}")

    def save_settings(self):
        """保存设置"""
        global flow_scripts_dir, screen_shot_dir
        
        # 获取用户设置的目录
        new_screenshot_dir = self.screenshot_dir_var.get()
        new_workflow_dir = self.workflow_dir_var.get()
        
        # 确保目录存在
        try:
            os.makedirs(new_screenshot_dir, exist_ok=True)
            os.makedirs(new_workflow_dir, exist_ok=True)
            
            # 更新全局变量
            screen_shot_dir = new_screenshot_dir
            flow_scripts_dir = new_workflow_dir
            
            # 保存到配置文件
            config = {
                'flow_scripts_dir': new_workflow_dir,
                'screen_shot_dir': new_screenshot_dir
            }
            save_config(config)
            
            messagebox.showinfo("成功", "设置已保存！")
        except Exception as e:
            messagebox.showerror("错误", f"保存设置失败: {str(e)}")

    def restore_defaults(self):
        """恢复默认设置"""
        global flow_scripts_dir, screen_shot_dir
        
        # 恢复默认目录
        flow_scripts_dir = DEFAULT_FLOW_SCRIPTS_DIR
        screen_shot_dir = DEFAULT_SCREEN_SHOT_DIR
        
        # 更新界面
        self.workflow_dir_var.set(DEFAULT_FLOW_SCRIPTS_DIR)
        self.screenshot_dir_var.set(DEFAULT_SCREEN_SHOT_DIR)
        
        # 保存默认设置
        config = {
            'flow_scripts_dir': DEFAULT_FLOW_SCRIPTS_DIR,
            'screen_shot_dir': DEFAULT_SCREEN_SHOT_DIR
        }
        save_config(config)
        
        messagebox.showinfo("成功", "已恢复默认设置！")

    def update_accuracy_label(self, *args):
        accuracy = self.accuracy_var.get()
        self.accuracy_label.config(text=f"{int(accuracy * 100)}%")

    def update_condition_accuracy_label(self, *args):
        accuracy = self.condition_accuracy_var.get()
        self.condition_accuracy_label.config(text=f"{int(accuracy * 100)}%")

    def add_action(self):
        # 根据当前选中的选项卡确定操作类型
        selected_tab = self.action_notebook.select()
        tab_text = self.action_notebook.tab(selected_tab, "text")
        action_id = f"action_{self.action_counter}"
        self.action_counter += 1

        if tab_text == "鼠标点击":
            target_type = self.mouse_target_var.get()
            click_type = self.click_type_var.get()
            name = self.mouse_name_var.get()
            if not name:
                messagebox.showerror("错误", "请输入操作名称")
                return

            click_type_names = {"left": "单击左键", "right": "单击右键", "double": "双击左键"}

            if target_type == "image":
                image_path = self.image_path_var.get()
                if not image_path:
                    messagebox.showerror("错误", "请选择截图文件")
                    return
                action = {
                    "id": action_id,
                    "name": name,
                    "type": "mouse",
                    "target_type": "image",
                    "click_type": click_type,
                    "image_path": image_path,
                    "accuracy": self.accuracy_var.get()
                }
                details = f"{click_type_names[click_type]}, 截图: {os.path.basename(image_path)}"
            else:
                text = self.mouse_text_var.get()
                if not text:
                    messagebox.showerror("错误", "请输入目标文字")
                    return
                action = {
                    "id": action_id,
                    "name": name,
                    "type": "mouse",
                    "target_type": "text",
                    "click_type": click_type,
                    "text": text,
                    "accuracy": self.accuracy_var.get()
                }
                details = f"{click_type_names[click_type]}, 文字: {text}"
            action_type_display = "鼠标点击"

        elif tab_text == "键盘按键":
            trigger_type = self.keyboard_trigger_var.get()
            keys = self.keys_var.get()
            if not keys:
                messagebox.showerror("错误", "请输入按键组合")
                return

            name = self.keyboard_name_var.get()
            if not name:
                messagebox.showerror("错误", "请输入操作名称")
                return

            if trigger_type == "none":
                action = {
                    "id": action_id,
                    "name": name,
                    "type": "keyboard",
                    "trigger_type": "none",
                    "keys": keys
                }
                details = f"按键: {keys}"
            else:  # text
                text = self.keyboard_text_var.get()
                if not text:
                    messagebox.showerror("错误", "请输入目标文字")
                    return
                action = {
                    "id": action_id,
                    "name": name,
                    "type": "keyboard",
                    "trigger_type": "text",
                    "text": text,
                    "keys": keys
                }
                details = f"按键: {keys}, 触发文字: {text}"
            action_type_display = "键盘按键"

        elif tab_text == "系统命令":
            cmd_type = self.system_cmd_type_var.get()
            name = self.system_cmd_name_var.get()
            if not name:
                messagebox.showerror("错误", "请输入操作名称")
                return

            if cmd_type == "launch_app":
                app_path = self.app_path_var.get()
                if not app_path:
                    messagebox.showerror("错误", "请选择要启动的应用")
                    return
                action = {
                    "id": action_id,
                    "name": name,
                    "type": "system_command",
                    "cmd_type": "launch_app",
                    "app_path": app_path
                }
                details = f"启动应用: {os.path.basename(app_path)}"
            elif cmd_type == "set_window":
                window_title = self.window_title_var.get()
                window_class = self.window_class_var.get()
                try:
                    window_x = int(self.window_x_var.get())
                    window_y = int(self.window_y_var.get())
                    window_width = int(self.window_width_var.get())
                    window_height = int(self.window_height_var.get())
                except ValueError:
                    messagebox.showerror("错误", "窗口位置和大小必须是整数")
                    return
                
                if not window_title and not window_class:
                    messagebox.showerror("错误", "请输入窗口标题或窗口类名")
                    return
                
                action = {
                    "id": action_id,
                    "name": name,
                    "type": "system_command",
                    "cmd_type": "set_window",
                    "window_title": window_title,
                    "window_class": window_class,
                    "window_x": window_x,
                    "window_y": window_y,
                    "window_width": window_width,
                    "window_height": window_height
                }
                details = f"固定窗体: ({window_x},{window_y}) {window_width}x{window_height}"
            action_type_display = "系统命令"

        elif tab_text == "写入剪贴板":
            text = self.clipboard_text.get(1.0, tk.END).strip()
            if not text:
                messagebox.showerror("错误", "请输入要写入剪贴板的文字")
                return

            name = self.clipboard_write_name_var.get()
            if not name:
                messagebox.showerror("错误", "请输入操作名称")
                return

            action = {
                "id": action_id,
                "name": name,
                "type": "clipboard_write",
                "text": text
            }
            details = f"文字: {text[:20]}..." if len(text) > 20 else f"文字: {text}"
            action_type_display = "写入剪贴板"

        elif tab_text == "条件判断":
            condition_type = self.condition_type_var.get()
            name = self.condition_name_var.get()
            if not name:
                messagebox.showerror("错误", "请输入操作名称")
                return

            if condition_type == "image":
                image_path = self.condition_image_var.get()
                if not image_path:
                    messagebox.showerror("错误", "请选择判断图片")
                    return
                action = {
                    "id": action_id,
                    "name": name,
                    "type": "condition",
                    "condition_type": "image",
                    "image_path": image_path,
                    "accuracy": self.condition_accuracy_var.get(),
                    "found_behavior": self.condition_found_behavior_var.get(),
                    "not_found_behavior": self.condition_not_found_behavior_var.get(),
                    "jump_to": self.condition_jump_var.get()
                }
                details = f"判断图片: {os.path.basename(image_path)}"
            else:  # text
                text = self.condition_text_var.get()
                if not text:
                    messagebox.showerror("错误", "请输入判断文字")
                    return
                action = {
                    "id": action_id,
                    "name": name,
                    "type": "condition",
                    "condition_type": "text",
                    "text": text,
                    "accuracy": self.condition_accuracy_var.get(),
                    "found_behavior": self.condition_found_behavior_var.get(),
                    "not_found_behavior": self.condition_not_found_behavior_var.get(),
                    "jump_to": self.condition_jump_var.get()
                }
                details = f"判断文字: {text}"
            action_type_display = "条件判断"

        else:
            messagebox.showerror("错误", "未知的操作类型")
            return

        self.actions.append(action)
        self.action_tree.insert("", tk.END, values=(action_id, name, action_type_display, details))
        self.update_available_actions()
        self.save_actions()
        messagebox.showinfo("成功", "操作添加成功")

    def test_action(self):
        # 根据当前选中的选项卡确定操作类型
        selected_tab = self.action_notebook.select()
        tab_text = self.action_notebook.tab(selected_tab, "text")

        if tab_text == "鼠标点击":
            target_type = self.mouse_target_var.get()
            
            try:
                accuracy = self.accuracy_var.get()
                if target_type == "image":
                    image_path = self.image_path_var.get()
                    if not image_path:
                        messagebox.showerror("错误", "请选择截图文件")
                        return
                    self.log(f"测试鼠标点击操作: 查找图像 {os.path.basename(image_path)} (精度: {int(accuracy * 100)}%)")
                    position = self.find_image(image_path, accuracy)
                    if position:
                        self.log(f"找到图像，位置: {position}")
                        click_type = self.click_type_var.get()
                        click_type_names = {"left": "单击左键", "right": "单击右键", "double": "双击左键"}
                        self.log(f"执行{click_type_names[click_type]}操作")
                        if click_type == "left":
                            pyautogui.click(position)
                        elif click_type == "right":
                            pyautogui.rightClick(position)
                        elif click_type == "double":
                            pyautogui.doubleClick(position)
                        messagebox.showinfo("成功", "鼠标点击测试成功")
                    else:
                        messagebox.showerror("错误", "未找到指定图像")
                else:  # text
                    text = self.mouse_text_var.get()
                    if not text:
                        messagebox.showerror("错误", "请输入目标文字")
                        return
                    self.log(f"测试鼠标点击操作: 查找文字 '{text}'")
                    self.log("模拟点击操作")
                    click_type = self.click_type_var.get()
                    if click_type == "left":
                        pyautogui.click()
                    elif click_type == "right":
                        pyautogui.rightClick()
                    elif click_type == "double":
                        pyautogui.doubleClick()
                    messagebox.showinfo("成功", "鼠标点击测试成功")
            except Exception as e:
                messagebox.showerror("错误", f"测试失败: {str(e)}")

        elif tab_text == "键盘按键":
            trigger_type = self.keyboard_trigger_var.get()
            keys = self.keys_var.get()
            if not keys:
                messagebox.showerror("错误", "请输入按键组合")
                return

            try:
                key_parts = keys.split('+')
                normalized_keys = [k.lower() for k in key_parts]
                
                if trigger_type == "none":
                    self.log(f"测试键盘按键操作: {keys}")
                    pyautogui.hotkey(*normalized_keys)
                    messagebox.showinfo("成功", "键盘按键测试成功")
                else:  # text
                    text = self.keyboard_text_var.get()
                    if not text:
                        messagebox.showerror("错误", "请输入目标文字")
                        return
                    self.log(f"测试键盘按键操作: 查找文字 '{text}' 然后执行 {keys}")
                    pyautogui.hotkey(*normalized_keys)
                    messagebox.showinfo("成功", "键盘按键测试成功")
            except Exception as e:
                messagebox.showerror("错误", f"测试失败: {str(e)}")

        elif tab_text == "系统命令":
            cmd_type = self.system_cmd_type_var.get()
            
            try:
                if cmd_type == "launch_app":
                    app_path = self.app_path_var.get()
                    if not app_path:
                        messagebox.showerror("错误", "请选择要启动的应用")
                        return
                    self.log(f"测试启动应用: {app_path}")
                    os.startfile(app_path)
                    messagebox.showinfo("成功", f"启动应用测试成功: {os.path.basename(app_path)}")
                elif cmd_type == "set_window":
                    if not HAS_WIN32:
                        messagebox.showerror("错误", "固定窗体功能需要安装 pywin32 库")
                        return
                    window_title = self.window_title_var.get()
                    window_class = self.window_class_var.get()
                    try:
                        window_x = int(self.window_x_var.get())
                        window_y = int(self.window_y_var.get())
                        window_width = int(self.window_width_var.get())
                        window_height = int(self.window_height_var.get())
                    except ValueError:
                        messagebox.showerror("错误", "窗口位置和大小必须是整数")
                        return
                    
                    if not window_title and not window_class:
                        messagebox.showerror("错误", "请输入窗口标题或窗口类名")
                        return
                    
                    self.log(f"测试固定窗体: 标题='{window_title}', 类名='{window_class}'")
                    hwnd = self.find_window(window_title, window_class)
                    if hwnd:
                        win32gui.SetWindowPos(hwnd, None, window_x, window_y, window_width, window_height, 0)
                        win32gui.SetForegroundWindow(hwnd)
                        self.log(f"窗口已移动到 ({window_x}, {window_y}), 大小 {window_width}x{window_height}")
                        messagebox.showinfo("成功", f"固定窗体测试成功")
                    else:
                        messagebox.showerror("错误", "未找到指定窗口")
            except Exception as e:
                messagebox.showerror("错误", f"测试失败: {str(e)}")

        elif tab_text == "写入剪贴板":
            text = self.clipboard_text.get(1.0, tk.END).strip()
            if not text:
                messagebox.showerror("错误", "请输入要写入剪贴板的文字")
                return

            try:
                self.log(f"测试写入剪贴板操作: {text[:20]}..." if len(text) > 20 else f"测试写入剪贴板操作: {text}")
                pyperclip.copy(text)
                messagebox.showinfo("成功", "写入剪贴板测试成功")
            except Exception as e:
                messagebox.showerror("错误", f"测试失败: {str(e)}")

        elif tab_text == "条件判断":
            condition_type = self.condition_type_var.get()
            
            try:
                accuracy = self.condition_accuracy_var.get()
                if condition_type == "image":
                    image_path = self.condition_image_var.get()
                    if not image_path:
                        messagebox.showerror("错误", "请选择判断图片")
                        return
                    self.log(f"测试条件判断操作: 查找图像 {os.path.basename(image_path)} (精度: {int(accuracy * 100)}%)")
                    position = self.find_image(image_path, accuracy)
                    if position:
                        self.log(f"找到图像，位置: {position}")
                        messagebox.showinfo("成功", "条件判断测试成功 - 找到图像")
                    else:
                        messagebox.showinfo("信息", "条件判断测试 - 未找到图像")
                else:  # text
                    text = self.condition_text_var.get()
                    if not text:
                        messagebox.showerror("错误", "请输入判断文字")
                        return
                    self.log(f"测试条件判断操作: 查找文字 '{text}' (精度: {int(accuracy * 100)}%)")
                    found = self.find_text(text)
                    if found:
                        self.log(f"找到文字: '{text}'")
                        messagebox.showinfo("成功", "条件判断测试成功 - 找到文字")
                    else:
                        messagebox.showinfo("信息", "条件判断测试 - 未找到文字")
            except Exception as e:
                messagebox.showerror("错误", f"测试失败: {str(e)}")

    def find_image(self, image_path, accuracy=0.9):
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
        
        self.log(f"尝试加载模板图像: {processed_path}")
        
        if not os.path.isfile(processed_path):
            raise Exception(f"模板文件不存在: {processed_path}")

        # 读取模板图像
        template = cv2.imread(processed_path, 0)
        if template is None:
            raise Exception(f"无法读取模板图像: {processed_path}")

        # 使用模板匹配
        result = cv2.matchTemplate(screen_gray, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

        # 设置阈值
        threshold = accuracy
        if max_val >= threshold:
            h, w = template.shape
            center_x = max_loc[0] + w // 2
            center_y = max_loc[1] + h // 2
            self.log(f"模板匹配成功，置信度: {max_val:.2f}，位置: ({center_x}, {center_y})")
            return (center_x, center_y)
        else:
            self.log(f"模板匹配失败，路径: {processed_path}，当前最大置信度: {max_val:.2f}")
            return None

    def find_text(self, text):
        # 截图当前屏幕
        screen = pyautogui.screenshot()
        
        # 使用pytesseract识别文字
        try:
            extracted_text = pytesseract.image_to_string(screen)
            # 检查目标文字是否在提取的文字中
            return text in extracted_text
        except Exception as e:
            # 如果pytesseract未安装或出现其他错误，返回False
            self.log(f"OCR识别失败: {str(e)}")
            return False

    def update_available_actions(self):
        # 清空现有项
        for item in self.available_actions_tree.get_children():
            self.available_actions_tree.delete(item)

        # 添加操作
        for action in self.actions:
            if action["type"] == "mouse":
                action_type_display = "鼠标点击"
            elif action["type"] == "keyboard":
                action_type_display = "键盘按键"
            elif action["type"] == "clipboard_write":
                action_type_display = "写入剪贴板"
            elif action["type"] == "clipboard_read":
                action_type_display = "读取剪贴板"
            elif action["type"] == "condition":
                action_type_display = "条件判断"
            elif action["type"] == "system_command":
                action_type_display = "系统命令"
            elif action["type"] == "doubao_screenshot":
                action_type_display = "豆包截图"
            else:
                action_type_display = "未知"
            
            self.available_actions_tree.insert("", tk.END, values=(action["id"], action["name"], action_type_display))

    def show_action_context_menu(self, event):
        item = self.available_actions_tree.identify_row(event.y)
        if item:
            self.available_actions_tree.selection_set(item)
            try:
                self.action_context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.action_context_menu.grab_release()

    def _execute_action(self, action):
        try:
            if action["type"] == "mouse":
                target_type = action.get("target_type", "image")
                click_type = action.get("click_type", "left")
                accuracy = action.get("accuracy", 0.9)
                
                if target_type == "image":
                    image_path = action.get("image_path", "")
                    position = self.find_image(image_path, accuracy)
                    if position:
                        self.log(f"找到图像，位置: {position}")
                        if click_type == "left":
                            pyautogui.click(position)
                        elif click_type == "right":
                            pyautogui.rightClick(position)
                        elif click_type == "double":
                            pyautogui.doubleClick(position)
                        self.log("鼠标点击测试成功")
                    else:
                        self.log(f"未找到图像: {image_path}")
                else:
                    text = action.get("text", "")
                    self.log(f"模拟点击: {text}")
                    if click_type == "left":
                        pyautogui.click()
                    elif click_type == "right":
                        pyautogui.rightClick()
                    elif click_type == "double":
                        pyautogui.doubleClick()
                    
            elif action["type"] == "keyboard":
                keys = action.get("keys", "")
                key_parts = keys.split('+')
                normalized_keys = [k.lower() for k in key_parts]
                pyautogui.hotkey(*normalized_keys)
                self.log(f"按键组合执行成功: {keys}")
                
            elif action["type"] == "system_command":
                cmd_type = action.get("cmd_type", "launch_app")
                if cmd_type == "launch_app":
                    app_path = action.get("app_path", "")
                    os.startfile(app_path)
                    self.log(f"启动应用: {os.path.basename(app_path)}")
                elif cmd_type == "set_window":
                    if HAS_WIN32:
                        window_title = action.get("window_title", "")
                        window_class = action.get("window_class", "")
                        window_x = action.get("window_x", 0)
                        window_y = action.get("window_y", 0)
                        window_width = action.get("window_width", 800)
                        window_height = action.get("window_height", 600)
                        hwnd = self.find_window(window_title, window_class)
                        if hwnd:
                            win32gui.SetWindowPos(hwnd, None, window_x, window_y, window_width, window_height, 0)
                            win32gui.SetForegroundWindow(hwnd)
                            self.log(f"固定窗体成功")
                        else:
                            self.log("未找到指定窗口")
                            
            elif action["type"] == "clipboard_write":
                text = action.get("text", "")
                pyperclip.copy(text)
                self.log("写入剪贴板成功")
                
            elif action["type"] == "clipboard_read":
                clipboard_text = pyperclip.paste()
                self.log(f"读取剪贴板成功: {clipboard_text[:30]}..." if len(clipboard_text) > 30 else f"读取剪贴板成功: {clipboard_text}")
                
            elif action["type"] == "doubao_screenshot":
                pyautogui.hotkey('shift', 'alt', 'a')
                self.log("豆包截图执行成功")
                
            elif action["type"] == "delay":
                duration = action.get("duration", 1)
                time.sleep(duration)
                self.log(f"延迟 {duration} 秒完成")
                
        except Exception as e:
            self.log(f"操作执行失败: {str(e)}")

    def show_action_tree_menu(self, event):
        item = self.action_tree.identify_row(event.y)
        if item:
            self.action_tree.selection_set(item)
            try:
                self.action_tree_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.action_tree_menu.grab_release()

    def test_action_from_menu(self):
        selected_item = self.action_tree.selection()
        if not selected_item:
            messagebox.showwarning("提示", "请先选择一个操作")
            return
        
        item = selected_item[0]
        action_id = self.action_tree.item(item, "values")[0]
        action = next((a for a in self.actions if a["id"] == action_id), None)
        
        if action:
            self._execute_action(action)

    def remove_action_from_menu(self):
        selected_item = self.action_tree.selection()
        if not selected_item:
            messagebox.showwarning("提示", "请先选择一个操作")
            return
        
        item = selected_item[0]
        action_id = self.action_tree.item(item, "values")[0]
        
        result = messagebox.askyesno("确认", "确定要移除这个操作吗？")
        if result:
            self.actions = [a for a in self.actions if a["id"] != action_id]
            self.action_tree.delete(item)
            self.save_actions()
            self.update_available_actions()
            self.log(f"操作已移除: {action_id}")

    def replace_workflow_step(self):
        selected_item = self.available_actions_tree.selection()
        if not selected_item:
            messagebox.showwarning("提示", "请先选择一个操作")
            return
        
        if not self.workflow:
            messagebox.showwarning("提示", "工作流为空，请先添加操作")
            return
        
        item = selected_item[0]
        action_id = self.available_actions_tree.item(item, "values")[0]
        action = next((a for a in self.actions if a["id"] == action_id), None)
        
        if action:
            step_num = simpledialog.askinteger("替换步骤", f"请输入要替换的步骤号 (1-{len(self.workflow)}):", 
                                                minvalue=1, maxvalue=len(self.workflow))
            if step_num:
                self.workflow[step_num - 1] = action
                self._refresh_workflow_tree()
                self.log(f"已替换步骤 {step_num} 为: {action['name']}")

    def insert_workflow_step(self):
        selected_item = self.available_actions_tree.selection()
        if not selected_item:
            messagebox.showwarning("提示", "请先选择一个操作")
            return
        
        item = selected_item[0]
        action_id = self.available_actions_tree.item(item, "values")[0]
        action = next((a for a in self.actions if a["id"] == action_id), None)
        
        if action:
            if not self.workflow:
                self.workflow.append(action)
                self._refresh_workflow_tree()
                self.log(f"已添加操作到工作流: {action['name']}")
            else:
                step_num = simpledialog.askinteger("插入步骤", f"请输入插入位置 (1-{len(self.workflow)}):", 
                                                    minvalue=1, maxvalue=len(self.workflow))
                if step_num:
                    self.workflow.insert(step_num, action)
                    self._refresh_workflow_tree()
                    self.log(f"已在步骤 {step_num} 后插入: {action['name']}")

    def _refresh_workflow_tree(self):
        for item in self.workflow_tree.get_children():
            self.workflow_tree.delete(item)
        
        for i, step in enumerate(self.workflow):
            action_type_display = self._get_action_type_display(step["type"])
            self.workflow_tree.insert("", tk.END, values=(i + 1, step["name"], action_type_display))
        
        self.update_available_actions()

    def _get_action_type_display(self, action_type):
        type_map = {
            "mouse": "鼠标点击",
            "keyboard": "键盘按键",
            "delay": "延迟",
            "clipboard_write": "写入剪贴板",
            "clipboard_read": "读取剪贴板",
            "condition": "条件判断",
            "system_command": "系统命令",
            "doubao_screenshot": "豆包截图"
        }
        return type_map.get(action_type, "未知")

    def add_to_workflow(self):
        selected_item = self.available_actions_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请选择要添加的操作")
            return

        item = selected_item[0]
        action_id = self.available_actions_tree.item(item, "values")[0]

        # 查找对应操作
        action = next((a for a in self.actions if a["id"] == action_id), None)
        if action:
            # 添加到工作流
            step_id = len(self.workflow) + 1
            self.workflow.append(action)
            if action["type"] == "mouse":
                action_type_display = "鼠标点击"
            elif action["type"] == "keyboard":
                action_type_display = "键盘按键"
            elif action["type"] == "delay":
                action_type_display = "延迟"
            elif action["type"] == "clipboard_write":
                action_type_display = "写入剪贴板"
            elif action["type"] == "clipboard_read":
                action_type_display = "读取剪贴板"
            elif action["type"] == "condition":
                action_type_display = "条件判断"
            else:
                action_type_display = "未知"
            self.workflow_tree.insert("", tk.END, values=(step_id, action["name"], action_type_display))

    def remove_from_workflow(self):
        selected_item = self.workflow_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请选择要移除的操作")
            return

        item = selected_item[0]
        index = self.workflow_tree.index(item)
        self.workflow_tree.delete(item)
        self.workflow.pop(index)

        # 更新步骤编号
        for i, item in enumerate(self.workflow_tree.get_children()):
            values = list(self.workflow_tree.item(item, "values"))
            values[0] = i + 1
            self.workflow_tree.item(item, values=values)

    def move_up(self):
        selected_item = self.workflow_tree.selection()
        if not selected_item:
            return

        item = selected_item[0]
        index = self.workflow_tree.index(item)
        if index > 0:
            # 交换工作流中的顺序
            self.workflow[index], self.workflow[index - 1] = self.workflow[index - 1], self.workflow[index]

            # 交换树中的顺序
            self.workflow_tree.move(item, self.workflow_tree.parent(item), index - 1)

            # 更新步骤编号
            for i, item in enumerate(self.workflow_tree.get_children()):
                values = list(self.workflow_tree.item(item, "values"))
                values[0] = i + 1
                self.workflow_tree.item(item, values=values)

    def move_down(self):
        selected_item = self.workflow_tree.selection()
        if not selected_item:
            return

        item = selected_item[0]
        index = self.workflow_tree.index(item)
        if index < len(self.workflow) - 1:
            # 交换工作流中的顺序
            self.workflow[index], self.workflow[index + 1] = self.workflow[index + 1], self.workflow[index]

            # 交换树中的顺序
            self.workflow_tree.move(item, self.workflow_tree.parent(item), index + 2)

            # 更新步骤编号
            for i, item in enumerate(self.workflow_tree.get_children()):
                values = list(self.workflow_tree.item(item, "values"))
                values[0] = i + 1
                self.workflow_tree.item(item, values=values)

    def replace_workflow(self):
        # 检查是否选择了要替换的操作
        selected_item = self.available_actions_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请先从左边的可用操作列表中选择一个操作")
            return

        # 获取选中的操作
        item = selected_item[0]
        action_id = self.available_actions_tree.item(item, "values")[0]
        action = next((a for a in self.actions if a["id"] == action_id), None)
        if not action:
            messagebox.showerror("错误", "未找到选中的操作")
            return

        # 弹出对话框询问要替换的步骤
        step_str = simpledialog.askstring("替换工作流", "请输入要替换的步骤编号:")
        if not step_str:
            return

        # 验证输入是否为数字
        try:
            step_num = int(step_str)
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")
            return

        # 验证步骤编号是否有效
        if step_num < 1 or step_num > len(self.workflow):
            messagebox.showerror("错误", f"步骤编号必须在 1 到 {len(self.workflow)} 之间")
            return

        # 替换步骤
        index = step_num - 1
        self.workflow[index] = action

        # 更新工作流树
        tree_item = self.workflow_tree.get_children()[index]
        if action["type"] == "mouse":
            action_type_display = "鼠标点击"
        elif action["type"] == "keyboard":
            action_type_display = "键盘按键"
        elif action["type"] == "delay":
            action_type_display = "延迟"
        elif action["type"] == "clipboard_write":
            action_type_display = "写入剪贴板"
        elif action["type"] == "clipboard_read":
            action_type_display = "读取剪贴板"
        elif action["type"] == "condition":
            action_type_display = "条件判断"
        else:
            action_type_display = "未知"
        self.workflow_tree.item(tree_item, values=(step_num, action["name"], action_type_display))

        messagebox.showinfo("成功", f"已成功替换步骤 {step_num} 为操作 '{action['name']}'")

    def read_keyboard_input(self):
        # 创建一个新窗口用于显示键盘监听状态
        listen_window = tk.Toplevel(self.root)
        listen_window.title("读取键盘操作")
        listen_window.geometry("400x150")
        listen_window.transient(self.root)
        listen_window.grab_set()

        # 显示提示信息
        ttk.Label(listen_window, text="请按下您想要录制的按键组合，然后按 Enter 键结束", font=("SimHei", 10)).pack(padx=20, pady=20)
        status_var = tk.StringVar(value="等待按键...")
        status_label = ttk.Label(listen_window, textvariable=status_var, font=("SimHei", 10, "bold"))
        status_label.pack(pady=10)

        # 用于存储按下的键
        pressed_keys = []
        # 键名映射，将pyautogui的键名映射为更友好的名称
        key_map = {
            'ctrl': 'ctrl', 'control': 'ctrl',
            'shift': 'shift',
            'alt': 'alt',
            'win': 'win', 'command': 'command',
            'enter': 'enter', 'return': 'enter',
            'tab': 'tab',
            'backspace': 'backspace',
            'delete': 'delete',
            'space': 'space',
            'up': 'up',
            'down': 'down',
            'left': 'left',
            'right': 'right'
        }

        def on_key_press(event):
            key = event.keysym.lower()
            # 忽略Enter键，留作结束录制
            if key == 'return':
                return
            # 将键名映射为更友好的名称
            key_name = key_map.get(key, key)
            if key_name not in pressed_keys:
                pressed_keys.append(key_name)
                # 更新状态显示
                status_var.set(f"已按下: {'+'.join(pressed_keys)}")

        def on_key_release(event):
            key = event.keysym.lower()
            # 如果是Enter键，结束录制
            if key == 'return':
                # 格式化按键组合
                if pressed_keys:
                    key_combination = '+'.join(sorted(pressed_keys))
                    self.keys_var.set(key_combination)
                    status_var.set(f"已记录: {key_combination}")
                    # 延迟关闭窗口
                    listen_window.after(1000, listen_window.destroy)
                else:
                    status_var.set("未记录到按键")
                    listen_window.after(1000, listen_window.destroy)
            else:
                # 移除释放的键
                key_name = key_map.get(key, key)
                if key_name in pressed_keys:
                    pressed_keys.remove(key_name)
                    # 更新状态显示
                    if pressed_keys:
                        status_var.set(f"已按下: {'+'.join(pressed_keys)}")
                    else:
                        status_var.set("等待按键...")

        # 绑定键盘事件
        listen_window.bind('<KeyPress>', on_key_press)
        listen_window.bind('<KeyRelease>', on_key_release)

        # 显示窗口
        listen_window.focus_set()
        listen_window.mainloop()

    def add_delay(self):
        # 创建延迟操作
        delay_action = {
            "id": f"delay_{int(time.time())}",
            "name": "延迟操作",
            "type": "delay",
            "duration": 1  # 默认1秒
        }

        # 添加到工作流
        step_id = len(self.workflow) + 1
        self.workflow.append(delay_action)
        self.workflow_tree.insert("", tk.END, values=(step_id, delay_action["name"], "延迟"))

    def insert_delay(self):
        if not self.workflow:
            messagebox.showwarning("提示", "工作流为空，请先添加操作")
            return

        step_num = simpledialog.askinteger("插入延迟", f"请输入要插入到第几步后面？\n(当前共 {len(self.workflow)} 步)", minvalue=0, maxvalue=len(self.workflow))
        if step_num is None:
            return

        delay_action = {
            "id": f"delay_{int(time.time())}",
            "name": "延迟操作",
            "type": "delay",
            "duration": 1
        }

        insert_index = step_num
        self.workflow.insert(insert_index, delay_action)
        self._refresh_workflow_tree()
        messagebox.showinfo("成功", f"延迟已插入到第 {step_num} 步后面")

    def save_workflow(self):
        if not self.workflow:
            messagebox.showerror("错误", "工作流为空")
            return

        file_path = filedialog.asksaveasfilename(
            initialdir=flow_scripts_dir,
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )
        if file_path:
            workflow_data = {
                "name": os.path.basename(file_path).replace(".json", ""),
                "steps": self.workflow,
                "created_at": datetime.now().isoformat()
            }

            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(workflow_data, f, ensure_ascii=False, indent=2)

            messagebox.showinfo("成功", "工作流保存成功")

    def clear_workflow(self):
        if not self.workflow:
            messagebox.showinfo("提示", "工作流已经是空的")
            return
        
        result = messagebox.askyesno("确认", "确定要清空当前工作流吗？")
        if result:
            self.workflow.clear()
            for item in self.workflow_tree.get_children():
                self.workflow_tree.delete(item)
            self.workflow_frame.config(text="工作流")
            self.update_available_actions()
            messagebox.showinfo("成功", "工作流已清空")

    def load_workflow(self):
        file_path = filedialog.askopenfilename(
            initialdir=flow_scripts_dir,
            filetypes=[("JSON files", "*.json")]
        )
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    workflow_data = json.load(f)

                self.workflow.clear()
                for item in self.workflow_tree.get_children():
                    self.workflow_tree.delete(item)

                workflow_name = workflow_data.get("name", os.path.basename(file_path).replace(".json", ""))
                self.workflow_frame.config(text=f"工作流: {workflow_name}")

                for i, step in enumerate(workflow_data["steps"]):
                    self.workflow.append(step)
                    if step["type"] == "mouse":
                        action_type_display = "鼠标点击"
                    elif step["type"] == "keyboard":
                        action_type_display = "键盘按键"
                    elif step["type"] == "delay":
                        action_type_display = "延迟"
                    elif step["type"] == "clipboard_write":
                        action_type_display = "写入剪贴板"
                    elif step["type"] == "clipboard_read":
                        action_type_display = "读取剪贴板"
                    elif step["type"] == "condition":
                        action_type_display = "条件判断"
                    elif step["type"] == "system_command":
                        action_type_display = "系统命令"
                    else:
                        action_type_display = "未知"
                    
                    self.workflow_tree.insert("", tk.END, values=(i + 1, step["name"], action_type_display))

                self.update_available_actions()
                messagebox.showinfo("成功", f"工作流 '{workflow_name}' 加载成功")
            except Exception as e:
                messagebox.showerror("错误", f"加载失败: {str(e)}")

    def refresh_workflows(self):
        # 获取工作流文件
        workflows = []
        for file in os.listdir(flow_scripts_dir):
            if file.endswith(".json"):
                try:
                    with open(os.path.join(flow_scripts_dir, file), "r", encoding="utf-8") as f:
                        data = json.load(f)
                        if "steps" in data:
                            workflows.append(file)
                except:
                    pass

        self.workflow_combobox['values'] = workflows
        if workflows:
            self.workflow_var.set(workflows[0])

    def execute_workflow(self):
        workflow_file = self.workflow_var.get()
        if not workflow_file:
            messagebox.showerror("错误", "请选择要执行的工作流")
            return

        # 重置执行标志
        self.is_executing = True
        self.stop_requested = False

        # 使用线程执行工作流，避免阻塞GUI
        import threading
        def execute_in_thread():
            try:
                with open(os.path.join(flow_scripts_dir, workflow_file), "r", encoding="utf-8") as f:
                    workflow_data = json.load(f)

                # 获取执行设置
                execution_count = self.execution_count_var.get()
                execution_duration = self.execution_duration_var.get() * 60  # 转换为秒
                execution_interval = self.execution_interval_var.get()
                retry_on_error = self.retry_on_error_var.get()

                # 使用root.after更新GUI
                self.root.after(0, lambda: self.log(f"开始执行工作流: {workflow_data['name']}"))
                self.root.after(0, lambda: self.log(f"执行设置: 次数={execution_count}, 时长={execution_duration//60}分钟, 间隔={execution_interval}秒, 错误重试={retry_on_error}"))

                # 执行计数和时间
                count = 0
                start_time = time.time()
                stop_execution = False

                while not stop_execution:
                    # 检查是否请求停止
                    if self.stop_requested:
                        self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                        break

                    # 检查执行次数
                    if execution_count > 0 and count >= execution_count:
                        self.root.after(0, lambda count=count: self.log(f"已达到执行次数限制: {count}"))
                        break

                    # 检查执行时长
                    if execution_duration > 0 and time.time() - start_time > execution_duration:
                        self.root.after(0, lambda duration=execution_duration: self.log(f"已达到执行时长限制: {duration//60}分钟"))
                        break

                    # 执行工作流
                    self.root.after(0, lambda count=count: self.log(f"执行第 {count + 1} 轮"))
                    success = True

                    try:
                        for i, step in enumerate(workflow_data['steps']):
                            # 检查是否请求停止
                            if self.stop_requested:
                                self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                success = False
                                break

                            self.root.after(0, lambda i=i, step=step: self.log(f"执行步骤 {i + 1}: {step['name']}"))

                            if step['type'] == "mouse":
                                if self.stop_requested:
                                    self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                    success = False
                                    break
                                
                                target_type = step.get('target_type', 'image')
                                click_type = step.get('click_type', 'left')
                                accuracy = step.get('accuracy', 0.9)
                                click_type_names = {"left": "单击左键", "right": "单击右键", "double": "双击左键"}
                                
                                if target_type == "image":
                                    image_path = step.get('image_path', '')
                                    position = self.find_image(image_path, accuracy)
                                    if position:
                                        self.root.after(0, lambda position=position: self.log(f"找到图像，位置: {position}"))
                                        if click_type == "left":
                                            pyautogui.click(position)
                                        elif click_type == "right":
                                            pyautogui.rightClick(position)
                                        elif click_type == "double":
                                            pyautogui.doubleClick(position)
                                        self.root.after(0, lambda ct=click_type: self.log(f"{click_type_names.get(ct, '单击左键')}成功"))
                                    else:
                                        self.root.after(0, lambda step=step: self.log(f"未找到图像: {step['image_path']}"))
                                        if not retry_on_error:
                                            self.root.after(0, lambda step=step: messagebox.showerror("错误", f"未找到相应色块: {os.path.basename(step['image_path'])}"))
                                        success = False
                                        break
                                else:  # text
                                    text = step.get('text', '')
                                    self.root.after(0, lambda text=text: self.log(f"查找文字: '{text}'"))
                                    if click_type == "left":
                                        pyautogui.click()
                                    elif click_type == "right":
                                        pyautogui.rightClick()
                                    elif click_type == "double":
                                        pyautogui.doubleClick()
                                    self.root.after(0, lambda text=text, ct=click_type: self.log(f"执行{click_type_names.get(ct, '单击左键')}: '{text}'"))
                            elif step['type'] == "system_command":
                                if self.stop_requested:
                                    self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                    success = False
                                    break
                                
                                cmd_type = step.get('cmd_type', 'launch_app')
                                if cmd_type == "launch_app":
                                    app_path = step.get('app_path', '')
                                    self.root.after(0, lambda app_path=app_path: self.log(f"启动应用: {os.path.basename(app_path)}"))
                                    try:
                                        os.startfile(app_path)
                                        self.root.after(0, lambda app_path=app_path: self.log(f"应用启动成功: {os.path.basename(app_path)}"))
                                    except Exception as e:
                                        self.root.after(0, lambda e=e: self.log(f"应用启动失败: {str(e)}"))
                                        success = False
                                        break
                                elif cmd_type == "set_window":
                                    window_title = step.get('window_title', '')
                                    window_class = step.get('window_class', '')
                                    window_x = step.get('window_x', 0)
                                    window_y = step.get('window_y', 0)
                                    window_width = step.get('window_width', 800)
                                    window_height = step.get('window_height', 600)
                                    
                                    self.root.after(0, lambda t=window_title, c=window_class: self.log(f"固定窗体: 标题='{t}', 类名='{c}'"))
                                    hwnd = self.find_window(window_title, window_class)
                                    if hwnd:
                                        try:
                                            win32gui.SetWindowPos(hwnd, None, window_x, window_y, window_width, window_height, 0)
                                            win32gui.SetForegroundWindow(hwnd)
                                            self.root.after(0, lambda x=window_x, y=window_y, w=window_width, h=window_height: 
                                                           self.log(f"窗口已移动到 ({x}, {y}), 大小 {w}x{h}"))
                                        except Exception as e:
                                            self.root.after(0, lambda e=e: self.log(f"固定窗体失败: {str(e)}"))
                                            success = False
                                            break
                                    else:
                                        self.root.after(0, lambda: self.log("未找到指定窗口"))
                                        success = False
                                        break
                            elif step['type'] == "keyboard":
                                # 检查是否请求停止
                                if self.stop_requested:
                                    self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                    success = False
                                    break
                                
                                trigger_type = step.get('trigger_type', 'none')
                                keys = step.get('keys', '')
                                
                                try:
                                    key_parts = keys.split('+')
                                    normalized_keys = []
                                    for key in key_parts:
                                        key_lower = key.lower()
                                        normalized_keys.append(key_lower)
                                    
                                    if trigger_type == "none":
                                        pyautogui.hotkey(*normalized_keys)
                                        self.root.after(0, lambda keys=keys: self.log(f"按键组合执行成功: {keys}"))
                                    else:  # text
                                        text = step.get('text', '')
                                        self.root.after(0, lambda text=text: self.log(f"查找文字: '{text}' 以触发按键操作"))
                                        pyautogui.hotkey(*normalized_keys)
                                        self.root.after(0, lambda keys=keys: self.log(f"按键组合执行成功: {keys} (文字触发)"))
                                except Exception as e:
                                    self.root.after(0, lambda keys=keys, e=e: self.log(f"按键执行失败: {keys}, 错误: {str(e)}"))
                                    success = False
                                    break
                            elif step['type'] == "delay":
                                # 检查是否请求停止
                                if self.stop_requested:
                                    self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                    success = False
                                    break
                                
                                duration = step.get('duration', 1)
                                self.root.after(0, lambda duration=duration: self.log(f"执行延迟操作: {duration}秒"))
                                # 分段睡眠，以便能够响应停止请求
                                remaining = duration
                                while remaining > 0 and not self.stop_requested:
                                    sleep_time = min(0.5, remaining)
                                    time.sleep(sleep_time)
                                    remaining -= sleep_time
                                if self.stop_requested:
                                    self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                    success = False
                                    break
                            elif step['type'] == "clipboard_write":
                                # 检查是否请求停止
                                if self.stop_requested:
                                    self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                    success = False
                                    break
                                
                                text = step.get('text', '')
                                self.root.after(0, lambda text=text: self.log(f"执行写入剪贴板操作: {text[:20]}..." if len(text) > 20 else f"执行写入剪贴板操作: {text}"))
                                try:
                                    pyperclip.copy(text)
                                    self.root.after(0, lambda: self.log("写入剪贴板成功"))
                                except Exception as e:
                                    self.root.after(0, lambda e=e: self.log(f"写入剪贴板失败: {str(e)}"))
                                    if not retry_on_error:
                                        self.root.after(0, lambda e=e: messagebox.showerror("错误", f"写入剪贴板失败: {str(e)}"))
                            elif step['type'] == "clipboard_read":
                                # 检查是否请求停止
                                if self.stop_requested:
                                    self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                    success = False
                                    break
                                
                                try:
                                    clipboard_text = pyperclip.paste()
                                    self.root.after(0, lambda text=clipboard_text: self.log(f"执行读取剪贴板操作: {text[:20]}..." if len(text) > 20 else f"执行读取剪贴板操作: {text}"))
                                    self.root.after(0, lambda: self.log("读取剪贴板成功"))
                                except Exception as e:
                                    self.root.after(0, lambda e=e: self.log(f"读取剪贴板失败: {str(e)}"))
                                    if not retry_on_error:
                                        self.root.after(0, lambda e=e: messagebox.showerror("错误", f"读取剪贴板失败: {str(e)}"))
                            elif step['type'] == "doubao_screenshot":
                                if self.stop_requested:
                                    self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                    success = False
                                    break
                                
                                try:
                                    pyautogui.hotkey('shift', 'alt', 'a')
                                    self.root.after(0, lambda: self.log("豆包截图执行成功"))
                                except Exception as e:
                                    self.root.after(0, lambda e=e: self.log(f"豆包截图执行失败: {str(e)}"))
                                    if not retry_on_error:
                                        self.root.after(0, lambda e=e: messagebox.showerror("错误", f"豆包截图执行失败: {str(e)}"))
                            elif step['type'] == "condition":
                                # 检查是否请求停止
                                if self.stop_requested:
                                    self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                    success = False
                                    break
                                
                                condition_type = step.get('condition_type', 'image')
                                accuracy = step.get('accuracy', 0.9)
                                found_behavior = step.get('found_behavior', 'continue')
                                not_found_behavior = step.get('not_found_behavior', 'jump')
                                jump_to = step.get('jump_to', 1)
                                
                                found = False
                                if condition_type == "image":
                                    image_path = step.get('image_path', '')
                                    self.root.after(0, lambda image_path=image_path, accuracy=accuracy: self.log(f"执行条件判断操作: 等待图像 {os.path.basename(image_path)} 出现 (精度: {int(accuracy * 100)}%)"))
                                    # 查找图像
                                    position = self.find_image(image_path, accuracy)
                                    if position:
                                        self.root.after(0, lambda position=position: self.log(f"找到图像，位置: {position}"))
                                        found = True
                                    else:
                                        self.root.after(0, lambda image_path=image_path: self.log(f"未找到图像 {os.path.basename(image_path)}"))
                                else:  # text
                                    text = step.get('text', '')
                                    self.root.after(0, lambda text=text: self.log(f"执行条件判断操作: 等待文字 '{text}' 出现"))
                                    # 查找文字
                                    found = self.find_text(text)
                                    if found:
                                        self.root.after(0, lambda text=text: self.log(f"找到文字: '{text}'"))
                                    else:
                                        self.root.after(0, lambda text=text: self.log(f"未找到文字: '{text}'"))
                                
                                if found:
                                    # 处理找到时的行为
                                    if found_behavior == "continue":
                                        self.root.after(0, lambda: self.log("条件满足，继续执行"))
                                    elif found_behavior == "jump":
                                        # 跳转到指定步骤
                                        self.root.after(0, lambda jump_to=jump_to: self.log(f"跳转到步骤 {jump_to}"))
                                        # 调整循环索引，实现跳转
                                        # 注意：索引从0开始，步骤从1开始，所以需要减2
                                        i = jump_to - 2
                                        if i < -1:
                                            i = -1
                                        self.root.after(0, lambda i=i: self.log(f"跳转后执行步骤 {i + 2}"))
                                    elif found_behavior == "restart":
                                        # 重新开始工作流
                                        self.root.after(0, lambda: self.log("重新开始工作流"))
                                        # 跳出当前循环，开始新一轮执行
                                        break
                                else:
                                    if not_found_behavior == "jump":
                                        # 跳转到指定步骤
                                        self.root.after(0, lambda jump_to=jump_to: self.log(f"跳转到步骤 {jump_to}"))
                                        # 调整循环索引，实现跳转
                                        # 注意：索引从0开始，步骤从1开始，所以需要减2
                                        i = jump_to - 2
                                        if i < -1:
                                            i = -1
                                        self.root.after(0, lambda i=i: self.log(f"跳转后执行步骤 {i + 2}"))
                                    elif not_found_behavior == "wait":
                                        # 继续找，直到找到为止（添加超时机制）
                                        self.root.after(0, lambda: self.log("继续找，直到找到为止"))
                                        start_time = time.time()
                                        timeout = 60  # 60秒超时
                                        while not found:
                                            # 检查是否请求停止
                                            if self.stop_requested:
                                                self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                                success = False
                                                break
                                            
                                            if condition_type == "image":
                                                image_path = step.get('image_path', '')
                                                position = self.find_image(image_path, accuracy)
                                                if position:
                                                    self.root.after(0, lambda position=position: self.log(f"找到图像，位置: {position}"))
                                                    found = True
                                            else:  # text
                                                text = step.get('text', '')
                                                found = self.find_text(text)
                                                if found:
                                                    self.root.after(0, lambda text=text: self.log(f"找到文字: '{text}'"))
                                            if found:
                                                self.root.after(0, lambda: self.log("条件满足，继续执行"))
                                            else:
                                                # 检查是否超时
                                                if time.time() - start_time > timeout:
                                                    self.root.after(0, lambda timeout=timeout: self.log(f"超时: 未找到目标，超过 {timeout} 秒"))
                                                    if not retry_on_error:
                                                        self.root.after(0, lambda timeout=timeout: messagebox.showerror("错误", f"超时: 未找到目标，超过 {timeout} 秒"))
                                                    success = False
                                                    break
                                                # 每2秒尝试一次，分段睡眠以便响应停止请求
                                                remaining = 2
                                                while remaining > 0 and not self.stop_requested:
                                                    sleep_time = min(0.5, remaining)
                                                    time.sleep(sleep_time)
                                                    remaining -= sleep_time
                                                if self.stop_requested:
                                                    self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                                    success = False
                                                    break
                                                self.root.after(0, lambda: self.log("继续找，直到找到为止..."))
                                        # 找到后，继续执行下一个步骤
                                        if found:
                                            self.root.after(0, lambda: self.log("找到目标，继续执行工作流的下一个步骤"))
                                        if success and self.stop_requested:
                                            success = False
                                            break
                                    else:
                                        # 默认行为：继续执行
                                        self.root.after(0, lambda: self.log("未找到目标，继续执行"))

                            # 等待1秒，分段睡眠以便响应停止请求
                            remaining = 1
                            while remaining > 0 and not self.stop_requested:
                                sleep_time = min(0.5, remaining)
                                time.sleep(sleep_time)
                                remaining -= sleep_time
                            if self.stop_requested:
                                self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                success = False
                                break

                    except Exception as e:
                        self.root.after(0, lambda e=e: self.log(f"执行过程中出现错误: {str(e)}"))
                        if not retry_on_error:
                            self.root.after(0, lambda e=e: messagebox.showerror("错误", f"执行过程中出现错误: {str(e)}"))
                        success = False

                    if success:
                        count += 1

                        # 检查是否需要执行间隔
                        if execution_interval > 0 and (execution_count == 0 or count < execution_count):
                            self.root.after(0, lambda execution_interval=execution_interval: self.log(f"执行间隔: {execution_interval}秒"))
                            # 分段睡眠，以便能够响应停止请求
                            remaining = execution_interval
                            while remaining > 0 and not self.stop_requested:
                                sleep_time = min(0.5, remaining)
                                time.sleep(sleep_time)
                                remaining -= sleep_time
                            if self.stop_requested:
                                self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                break
                    else:
                        if retry_on_error and not self.stop_requested:
                            self.root.after(0, lambda: self.log("出现错误，重新执行任务流"))
                            # 等待2秒后重新执行，分段睡眠以便响应停止请求
                            remaining = 2
                            while remaining > 0 and not self.stop_requested:
                                sleep_time = min(0.5, remaining)
                                time.sleep(sleep_time)
                                remaining -= sleep_time
                            if self.stop_requested:
                                self.root.after(0, lambda: self.log("收到停止请求，正在停止执行..."))
                                break
                        else:
                            break

                if not retry_on_error and not self.stop_requested:
                    self.root.after(0, lambda count=count: self.log(f"工作流执行完成，共执行 {count} 轮"))
                    self.root.after(0, lambda count=count: messagebox.showinfo("成功", f"工作流执行完成，共执行 {count} 轮"))
                elif self.stop_requested:
                    self.root.after(0, lambda: self.log("工作流执行已停止"))
                    self.root.after(0, lambda: messagebox.showinfo("停止", "工作流执行已停止"))
            except Exception as e:
                self.root.after(0, lambda e=e: self.log(f"执行失败: {str(e)}"))
                if not self.retry_on_error_var.get():
                    self.root.after(0, lambda e=e: messagebox.showerror("错误", f"执行失败: {str(e)}"))
            finally:
                # 重置执行标志
                self.is_executing = False
                self.stop_requested = False

        # 启动线程
        thread = threading.Thread(target=execute_in_thread)
        thread.daemon = True
        thread.start()

    def stop_execution(self):
        # 实现停止执行的逻辑
        if self.is_executing:
            self.log("收到停止请求，正在停止执行...")
            self.stop_requested = True
        else:
            self.log("执行已停止")

    def export_workflow_as_script(self):
        workflow_file = self.workflow_var.get()
        if not workflow_file:
            messagebox.showerror("错误", "请选择要导出的工作流")
            return
        
        file_path = os.path.join(flow_scripts_dir, workflow_file)
        if not os.path.exists(file_path):
            messagebox.showerror("错误", f"工作流文件不存在: {workflow_file}")
            return
        
        with open(file_path, "r", encoding="utf-8") as f:
            workflow_data = json.load(f)
        
        script_name = workflow_data.get('name', 'workflow')
        save_path = filedialog.asksaveasfilename(
            initialdir=flow_py_dir,
            initialfile=f"{script_name}.py",
            defaultextension=".py",
            filetypes=[("Python files", "*.py")]
        )
        
        if not save_path:
            return
        
        script_content = self._generate_workflow_script(workflow_data)
        
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(script_content)
        
        self.log(f"工作流已导出为脚本: {save_path}")
        messagebox.showinfo("成功", f"工作流已导出为脚本:\n{save_path}")

    def _generate_workflow_script(self, workflow_data):
        script_lines = []
        script_lines.append("#!/usr/bin/env python3")
        script_lines.append("# -*- coding: utf-8 -*-")
        script_lines.append(f'"""')
        script_lines.append(f'工作流: {workflow_data.get("name", "未命名")}')
        script_lines.append(f'创建时间: {workflow_data.get("created_at", "未知")}')
        script_lines.append(f'导出时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        script_lines.append(f'"""')
        script_lines.append("")
        script_lines.append("import cv2")
        script_lines.append("import numpy as np")
        script_lines.append("import pyautogui")
        script_lines.append("import time")
        script_lines.append("import os")
        script_lines.append("import sys")
        script_lines.append("")
        script_lines.append("try:")
        script_lines.append("    import pyperclip")
        script_lines.append("except ImportError:")
        script_lines.append("    print('正在安装pyperclip...')")
        script_lines.append("    import subprocess")
        script_lines.append("    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyperclip'])")
        script_lines.append("    import pyperclip")
        script_lines.append("")
        script_lines.append("try:")
        script_lines.append("    import win32gui")
        script_lines.append("    HAS_WIN32 = True")
        script_lines.append("except ImportError:")
        script_lines.append("    HAS_WIN32 = False")
        script_lines.append("")
        script_lines.append(f"SCREEN_SHOT_DIR = r'{screen_shot_dir}'")
        script_lines.append("")
        script_lines.append("def find_image(image_path, accuracy=0.9):")
        script_lines.append("    screen = pyautogui.screenshot()")
        script_lines.append("    screen_gray = cv2.cvtColor(np.array(screen), cv2.COLOR_RGB2GRAY)")
        script_lines.append("    ")
        script_lines.append("    processed_path = image_path")
        script_lines.append("    if not os.path.isabs(processed_path):")
        script_lines.append("        processed_path = os.path.join(SCREEN_SHOT_DIR, os.path.basename(processed_path))")
        script_lines.append("    ")
        script_lines.append("    processed_path = os.path.normpath(processed_path)")
        script_lines.append("    if not os.path.isfile(processed_path):")
        script_lines.append("        print(f'模板文件不存在: {processed_path}')")
        script_lines.append("        return None")
        script_lines.append("    ")
        script_lines.append("    template = cv2.imread(processed_path, 0)")
        script_lines.append("    if template is None:")
        script_lines.append("        print(f'无法读取模板图像: {processed_path}')")
        script_lines.append("        return None")
        script_lines.append("    ")
        script_lines.append("    result = cv2.matchTemplate(screen_gray, template, cv2.TM_CCOEFF_NORMED)")
        script_lines.append("    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)")
        script_lines.append("    ")
        script_lines.append("    if max_val >= accuracy:")
        script_lines.append("        h, w = template.shape")
        script_lines.append("        center_x = max_loc[0] + w // 2")
        script_lines.append("        center_y = max_loc[1] + h // 2")
        script_lines.append("        print(f'找到图像，位置: ({center_x}, {center_y})，置信度: {max_val:.2f}')")
        script_lines.append("        return (center_x, center_y)")
        script_lines.append("    else:")
        script_lines.append("        print(f'未找到图像: {os.path.basename(image_path)}，置信度: {max_val:.2f}')")
        script_lines.append("        return None")
        script_lines.append("")
        script_lines.append("def find_window(title, class_name):")
        script_lines.append("    if not HAS_WIN32:")
        script_lines.append("        return None")
        script_lines.append("    result_hwnd = [None]")
        script_lines.append("    ")
        script_lines.append("    def enum_callback(hwnd, _):")
        script_lines.append("        if win32gui.IsWindowVisible(hwnd):")
        script_lines.append("            win_title = win32gui.GetWindowText(hwnd)")
        script_lines.append("            win_class = win32gui.GetClassName(hwnd)")
        script_lines.append("            title_match = not title or title.lower() in win_title.lower()")
        script_lines.append("            class_match = not class_name or class_name.lower() == win_class.lower()")
        script_lines.append("            if title_match and class_match and win_title:")
        script_lines.append("                result_hwnd[0] = hwnd")
        script_lines.append("                return False")
        script_lines.append("        return True")
        script_lines.append("    ")
        script_lines.append("    win32gui.EnumWindows(enum_callback, None)")
        script_lines.append("    return result_hwnd[0]")
        script_lines.append("")
        script_lines.append("def log(message):")
        script_lines.append("    timestamp = time.strftime('%H:%M:%S')")
        script_lines.append("    print(f'[{timestamp}] {message}')")
        script_lines.append("")
        script_lines.append("def execute_workflow():")
        script_lines.append("    steps = [")
        
        for step in workflow_data.get('steps', []):
            script_lines.append(f"        {step},")
        
        script_lines.append("    ]")
        script_lines.append("    ")
        script_lines.append("    log('开始执行工作流')")
        script_lines.append("    ")
        script_lines.append("    for i, step in enumerate(steps):")
        script_lines.append("        log(f'执行步骤 {i + 1}: {step[\"name\"]}')")
        script_lines.append("        ")
        script_lines.append("        if step['type'] == 'mouse':")
        script_lines.append("            target_type = step.get('target_type', 'image')")
        script_lines.append("            click_type = step.get('click_type', 'left')")
        script_lines.append("            accuracy = step.get('accuracy', 0.9)")
        script_lines.append("            ")
        script_lines.append("            if target_type == 'image':")
        script_lines.append("                image_path = step.get('image_path', '')")
        script_lines.append("                position = find_image(image_path, accuracy)")
        script_lines.append("                if position:")
        script_lines.append("                    if click_type == 'left':")
        script_lines.append("                        pyautogui.click(position)")
        script_lines.append("                    elif click_type == 'right':")
        script_lines.append("                        pyautogui.rightClick(position)")
        script_lines.append("                    elif click_type == 'double':")
        script_lines.append("                        pyautogui.doubleClick(position)")
        script_lines.append("                    log('点击成功')")
        script_lines.append("                else:")
        script_lines.append("                    log('未找到图像，跳过')")
        script_lines.append("            else:")
        script_lines.append("                if click_type == 'left':")
        script_lines.append("                    pyautogui.click()")
        script_lines.append("                elif click_type == 'right':")
        script_lines.append("                    pyautogui.rightClick()")
        script_lines.append("                elif click_type == 'double':")
        script_lines.append("                    pyautogui.doubleClick()")
        script_lines.append("        ")
        script_lines.append("        elif step['type'] == 'keyboard':")
        script_lines.append("            keys = step.get('keys', '')")
        script_lines.append("            key_parts = keys.split('+')")
        script_lines.append("            normalized_keys = [k.lower() for k in key_parts]")
        script_lines.append("            pyautogui.hotkey(*normalized_keys)")
        script_lines.append("            log(f'按键执行成功: {keys}')")
        script_lines.append("        ")
        script_lines.append("        elif step['type'] == 'system_command':")
        script_lines.append("            cmd_type = step.get('cmd_type', 'launch_app')")
        script_lines.append("            if cmd_type == 'launch_app':")
        script_lines.append("                app_path = step.get('app_path', '')")
        script_lines.append("                os.startfile(app_path)")
        script_lines.append("                log(f'启动应用: {os.path.basename(app_path)}')")
        script_lines.append("            elif cmd_type == 'set_window':")
        script_lines.append("                if HAS_WIN32:")
        script_lines.append("                    window_title = step.get('window_title', '')")
        script_lines.append("                    window_class = step.get('window_class', '')")
        script_lines.append("                    window_x = step.get('window_x', 0)")
        script_lines.append("                    window_y = step.get('window_y', 0)")
        script_lines.append("                    window_width = step.get('window_width', 800)")
        script_lines.append("                    window_height = step.get('window_height', 600)")
        script_lines.append("                    hwnd = find_window(window_title, window_class)")
        script_lines.append("                    if hwnd:")
        script_lines.append("                        win32gui.SetWindowPos(hwnd, None, window_x, window_y, window_width, window_height, 0)")
        script_lines.append("                        win32gui.SetForegroundWindow(hwnd)")
        script_lines.append("                        log('固定窗体成功')")
        script_lines.append("        ")
        script_lines.append("        elif step['type'] == 'delay':")
        script_lines.append("            duration = step.get('duration', 1)")
        script_lines.append("            time.sleep(duration)")
        script_lines.append("            log(f'延迟 {duration} 秒')")
        script_lines.append("        ")
        script_lines.append("        elif step['type'] == 'clipboard_write':")
        script_lines.append("            text = step.get('text', '')")
        script_lines.append("            pyperclip.copy(text)")
        script_lines.append("            log('写入剪贴板成功')")
        script_lines.append("        ")
        script_lines.append("        elif step['type'] == 'clipboard_read':")
        script_lines.append("            clipboard_text = pyperclip.paste()")
        script_lines.append("            log(f'读取剪贴板: {clipboard_text[:30]}...' if len(clipboard_text) > 30 else f'读取剪贴板: {clipboard_text}')")
        script_lines.append("        ")
        script_lines.append("        elif step['type'] == 'doubao_screenshot':")
        script_lines.append("            pyautogui.hotkey('shift', 'alt', 'a')")
        script_lines.append("            log('豆包截图执行成功')")
        script_lines.append("        ")
        script_lines.append("        time.sleep(0.5)")
        script_lines.append("    ")
        script_lines.append("    log('工作流执行完成')")
        script_lines.append("    return True")
        script_lines.append("")
        script_lines.append("def main():")
        script_lines.append("    import argparse")
        script_lines.append("    parser = argparse.ArgumentParser(description='工作流执行脚本')")
        script_lines.append("    parser.add_argument('-c', '--count', type=int, default=1, help='执行次数 (0=无限循环)')")
        script_lines.append("    parser.add_argument('-d', '--duration', type=int, default=0, help='执行时长(分钟, 0=不限制)')")
        script_lines.append("    parser.add_argument('-i', '--interval', type=int, default=0, help='执行间隔(秒)')")
        script_lines.append("    args = parser.parse_args()")
        script_lines.append("    ")
        script_lines.append("    print('=' * 50)")
        script_lines.append("    print('工作流执行脚本')")
        script_lines.append("    print(f'执行次数: {args.count}, 执行时长: {args.duration}分钟, 执行间隔: {args.interval}秒')")
        script_lines.append("    print('=' * 50)")
        script_lines.append("    ")
        script_lines.append("    exec_count = 0")
        script_lines.append("    start_time = time.time()")
        script_lines.append("    execution_duration = args.duration * 60")
        script_lines.append("    ")
        script_lines.append("    while True:")
        script_lines.append("        if args.count > 0 and exec_count >= args.count:")
        script_lines.append("            log(f'已达到执行次数限制: {exec_count}')")
        script_lines.append("            break")
        script_lines.append("        ")
        script_lines.append("        if execution_duration > 0 and time.time() - start_time > execution_duration:")
        script_lines.append("            log(f'已达到执行时长限制: {args.duration}分钟')")
        script_lines.append("            break")
        script_lines.append("        ")
        script_lines.append("        log(f'执行第 {exec_count + 1} 轮')")
        script_lines.append("        execute_workflow()")
        script_lines.append("        exec_count += 1")
        script_lines.append("        ")
        script_lines.append("        if args.interval > 0:")
        script_lines.append("            log(f'等待 {args.interval} 秒后继续...')")
        script_lines.append("            time.sleep(args.interval)")
        script_lines.append("    ")
        script_lines.append("    log(f'工作流执行完成! 共执行 {exec_count} 轮')")
        script_lines.append("")
        script_lines.append("if __name__ == '__main__':")
        script_lines.append("    main()")
        script_lines.append("")
        
        return "\n".join(script_lines)

    def log(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)

    def save_actions(self):
        with open("actions.json", "w", encoding="utf-8") as f:
            json.dump(self.actions, f, ensure_ascii=False, indent=2)

    def remove_action(self):
        selected_item = self.action_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请选择要移除的操作")
            return

        item = selected_item[0]
        action_id = self.action_tree.item(item, "values")[0]

        # 从操作列表中移除
        self.actions = [a for a in self.actions if a["id"] != action_id]

        # 从树中删除
        self.action_tree.delete(item)

        # 更新可用操作列表
        self.update_available_actions()

        # 保存操作列表
        self.save_actions()

        messagebox.showinfo("成功", "操作移除成功")

    def edit_action(self):
        selected_item = self.action_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请选择要修改的操作")
            return

        item = selected_item[0]
        action_id = self.action_tree.item(item, "values")[0]

        # 查找对应操作
        action = next((a for a in self.actions if a["id"] == action_id), None)
        if not action:
            messagebox.showerror("错误", "未找到选中的操作")
            return

        # 根据操作类型，在相应的选项卡中填充当前操作的配置
        if action["type"] == "mouse":
            for i, tab in enumerate(self.action_notebook.tabs()):
                if self.action_notebook.tab(tab, "text") == "鼠标点击":
                    self.action_notebook.select(tab)
                    break
            target_type = action.get("target_type", "image")
            self.mouse_target_var.set(target_type)
            self.click_type_var.set(action.get("click_type", "left"))
            if target_type == "image":
                self.image_path_var.set(action.get("image_path", ""))
                self.mouse_text_var.set("示例文字")
            else:
                self.image_path_var.set("")
                self.mouse_text_var.set(action.get("text", "示例文字"))
            self.mouse_name_var.set(action.get("name", ""))
            self.accuracy_var.set(action.get("accuracy", 0.9))
        elif action["type"] == "keyboard":
            # 切换到键盘按键选项卡
            for i, tab in enumerate(self.action_notebook.tabs()):
                if self.action_notebook.tab(tab, "text") == "键盘按键":
                    self.action_notebook.select(tab)
                    break
            # 填充配置
            trigger_type = action.get("trigger_type", "none")
            self.keyboard_trigger_var.set(trigger_type)
            if trigger_type == "text":
                self.keyboard_text_var.set(action.get("text", "示例文字"))
            else:
                self.keyboard_text_var.set("示例文字")
            self.keys_var.set(action.get("keys", ""))
            self.keyboard_name_var.set(action.get("name", ""))
        elif action["type"] == "system_command":
            for i, tab in enumerate(self.action_notebook.tabs()):
                if self.action_notebook.tab(tab, "text") == "系统命令":
                    self.action_notebook.select(tab)
                    break
            cmd_type = action.get("cmd_type", "launch_app")
            self.system_cmd_type_var.set(cmd_type)
            if cmd_type == "launch_app":
                self.app_path_var.set(action.get("app_path", ""))
            elif cmd_type == "set_window":
                self.window_title_var.set(action.get("window_title", ""))
                self.window_class_var.set(action.get("window_class", ""))
                self.window_x_var.set(str(action.get("window_x", 0)))
                self.window_y_var.set(str(action.get("window_y", 0)))
                self.window_width_var.set(str(action.get("window_width", 800)))
                self.window_height_var.set(str(action.get("window_height", 600)))
            self.system_cmd_name_var.set(action.get("name", ""))
            self.on_system_cmd_type_change()
        elif action["type"] == "clipboard_write":
            # 切换到写入剪贴板选项卡
            for i, tab in enumerate(self.action_notebook.tabs()):
                if self.action_notebook.tab(tab, "text") == "写入剪贴板":
                    self.action_notebook.select(tab)
                    break
            # 填充配置
            self.clipboard_text.delete(1.0, tk.END)
            self.clipboard_text.insert(tk.END, action.get("text", ""))
            self.clipboard_write_name_var.set(action.get("name", ""))
        elif action["type"] == "condition":
            # 切换到条件判断选项卡
            for i, tab in enumerate(self.action_notebook.tabs()):
                if self.action_notebook.tab(tab, "text") == "条件判断":
                    self.action_notebook.select(tab)
                    break
            # 填充配置
            condition_type = action.get("condition_type", "image")
            self.condition_type_var.set(condition_type)
            if condition_type == "image":
                self.condition_image_var.set(action.get("image_path", ""))
                self.condition_text_var.set("示例文字")
            else:
                self.condition_image_var.set("")
                self.condition_text_var.set(action.get("text", "示例文字"))
            self.condition_name_var.set(action.get("name", ""))
            self.condition_accuracy_var.set(action.get("accuracy", 0.9))
            self.condition_found_behavior_var.set(action.get("found_behavior", "continue"))
            self.condition_not_found_behavior_var.set(action.get("not_found_behavior", "jump"))
            self.condition_jump_var.set(action.get("jump_to", 1))

        # 从操作列表中移除旧操作
        self.actions = [a for a in self.actions if a["id"] != action_id]
        # 从树中删除旧操作
        self.action_tree.delete(item)

        messagebox.showinfo("提示", "请修改操作配置，然后点击'添加操作'按钮保存修改")

    def add_clipboard_read_preset(self):
        # 检查是否已经存在读取剪贴板操作
        existing = any(action["type"] == "clipboard_read" for action in self.actions)
        if not existing:
            # 创建读取剪贴板预设操作
            action_id = f"action_{self.action_counter}"
            self.action_counter += 1
            
            action = {
                "id": action_id,
                "name": "读取剪贴板",
                "type": "clipboard_read"
            }
            
            self.actions.append(action)
            self.action_tree.insert("", tk.END, values=(action_id, "读取剪贴板", "读取剪贴板", "读取剪贴板内容"))
            self.update_available_actions()
            self.save_actions()

    def add_enter_key_preset(self):
        # 检查是否已经存在按下回车键操作
        existing = any(action["name"] == "按下回车键" for action in self.actions)
        if not existing:
            # 创建按下回车键预设操作
            action_id = f"action_{self.action_counter}"
            self.action_counter += 1
            
            action = {
                "id": action_id,
                "name": "按下回车键",
                "type": "keyboard",
                "trigger_type": "none",
                "keys": "enter"
            }
            
            self.actions.append(action)
            self.action_tree.insert("", tk.END, values=(action_id, "按下回车键", "键盘按键", "按键: enter"))
            self.update_available_actions()
            self.save_actions()

    def add_doubao_screenshot_preset(self):
        existing = any(action["name"] == "豆包截图" for action in self.actions)
        if not existing:
            action_id = f"action_{self.action_counter}"
            self.action_counter += 1
            
            action = {
                "id": action_id,
                "name": "豆包截图",
                "type": "doubao_screenshot"
            }
            
            self.actions.append(action)
            self.action_tree.insert("", tk.END, values=(action_id, "豆包截图", "豆包截图", "执行豆包截图功能"))
            self.update_available_actions()
            self.save_actions()

    def insert_workflow(self):
        # 检查是否选择了要插入的操作
        selected_item = self.available_actions_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请先从左边的可用操作列表中选择一个操作")
            return

        # 获取选中的操作
        item = selected_item[0]
        action_id = self.available_actions_tree.item(item, "values")[0]
        action = next((a for a in self.actions if a["id"] == action_id), None)
        if not action:
            messagebox.showerror("错误", "未找到选中的操作")
            return

        # 弹出对话框询问要插入到第几步后面
        step_str = simpledialog.askstring("插入工作流", "请输入要插入到第几步后面:")
        if not step_str:
            return

        # 验证输入是否为数字
        try:
            step_num = int(step_str)
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")
            return

        # 验证步骤编号是否有效（可以是0，表示插入到最前面）
        if step_num < 0 or step_num > len(self.workflow):
            messagebox.showerror("错误", f"步骤编号必须在 0 到 {len(self.workflow)} 之间")
            return

        # 插入操作
        index = step_num
        self.workflow.insert(index, action)

        # 清空工作流树并重新填充
        for item in self.workflow_tree.get_children():
            self.workflow_tree.delete(item)

        # 重新填充工作流树
        for i, step in enumerate(self.workflow):
            if step["type"] == "mouse":
                action_type_display = "鼠标点击"
            elif step["type"] == "keyboard":
                action_type_display = "键盘按键"
            elif step["type"] == "delay":
                action_type_display = "延迟"
            elif step["type"] == "clipboard_write":
                action_type_display = "写入剪贴板"
            elif step["type"] == "clipboard_read":
                action_type_display = "读取剪贴板"
            elif step["type"] == "condition":
                action_type_display = "条件判断"
            else:
                action_type_display = "未知"
            self.workflow_tree.insert("", tk.END, values=(i + 1, step["name"], action_type_display))

        messagebox.showinfo("成功", f"已成功将操作 '{action['name']}' 插入到步骤 {step_num} 后面")

    def generate_script(self):
        # 检查工作流是否为空
        if not self.workflow:
            messagebox.showerror("错误", "工作流为空，无法生成脚本")
            return

        # 生成脚本内容
        script_content = '''
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
由工作流自动化工具生成的脚本
生成时间: {timestamp}
"""

import time
import pyautogui
import cv2
import numpy as np
from PIL import Image

# 尝试导入pyperclip，如果没有则安装
try:
    import pyperclip
except ImportError:
    import subprocess
    import sys
    print("正在安装pyperclip...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyperclip"])
    print("pyperclip安装完成")
    import pyperclip

# 尝试导入pytesseract，如果没有则安装
try:
    import pytesseract
except ImportError:
    import subprocess
    import sys
    print("正在安装pytesseract...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pytesseract", "Pillow"])
    print("pytesseract安装完成")
    import pytesseract


def find_image(image_path, accuracy=0.9):
    """
    在屏幕上查找指定图像
    :param image_path: 图像路径
    :param accuracy: 匹配精度
    :return: 找到的图像中心点坐标，未找到返回None
    """
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
    """
    在屏幕上查找指定文字
    :param text: 要查找的文字
    :return: 找到返回True，未找到返回False
    """
    # 截图当前屏幕
    screen = pyautogui.screenshot()
    
    # 使用pytesseract识别文字
    try:
        extracted_text = pytesseract.image_to_string(screen)
        # 检查目标文字是否在提取的文字中
        return text in extracted_text
    except Exception as e:
        # 如果pytesseract未安装或出现其他错误，返回False
        print(f"OCR识别失败: {str(e)}")
        return False


def main():
    """
    主函数，执行工作流
    """
    print("开始执行工作流...")

{workflow_code}

    print("工作流执行完成！")


if __name__ == "__main__":
    main()
'''

        # 生成工作流代码
        workflow_code = ""
        for i, step in enumerate(self.workflow):
            workflow_code += f"    # 步骤 {i + 1}: {step['name']}\n"
            
            if step['type'] == "mouse":
                image_path = step.get('image_path', '')
                accuracy = step.get('accuracy', 0.9)
                click_type = step.get('click_type', 'left')
                click_type_names = {"left": "单击左键", "right": "单击右键", "double": "双击左键"}
                click_cmd = {"left": "pyautogui.click(position)", "right": "pyautogui.rightClick(position)", "double": "pyautogui.doubleClick(position)"}
                click_name = click_type_names.get(click_type, "单击左键")
                workflow_code += f"    print('执行鼠标{click_name}: {step['name']}')\n"
                workflow_code += f"    position = find_image(r'{image_path}', {accuracy})\n"
                workflow_code += f"    if position:\n"
                workflow_code += f"        print(f'找到图像，位置: {{position}}')\n"
                workflow_code += f"        {click_cmd.get(click_type, 'pyautogui.click(position)')}\n"
                workflow_code += f"        print('{click_name}成功')\n"
                workflow_code += f"    else:\n"
                workflow_code += f"        print(f'未找到图像: {step['name']}')\n"
                workflow_code += f"        break\n"
            elif step['type'] == "system_command":
                cmd_type = step.get('cmd_type', 'launch_app')
                if cmd_type == "launch_app":
                    app_path = step.get('app_path', '')
                    workflow_code += f"    print('启动应用: {os.path.basename(app_path)}')\n"
                    workflow_code += f"    os.startfile(r'{app_path}')\n"
                    workflow_code += f"    print('应用启动成功')\n"
                elif cmd_type == "set_window":
                    window_title = step.get('window_title', '')
                    window_class = step.get('window_class', '')
                    window_x = step.get('window_x', 0)
                    window_y = step.get('window_y', 0)
                    window_width = step.get('window_width', 800)
                    window_height = step.get('window_height', 600)
                    workflow_code += f"    print('固定窗体: 标题=\"{window_title}\", 类名=\"{window_class}\"')\n"
                    workflow_code += f"    hwnd = find_window_by_title_class(r'{window_title}', r'{window_class}')\n"
                    workflow_code += f"    if hwnd:\n"
                    workflow_code += f"        win32gui.SetWindowPos(hwnd, None, {window_x}, {window_y}, {window_width}, {window_height}, 0)\n"
                    workflow_code += f"        win32gui.SetForegroundWindow(hwnd)\n"
                    workflow_code += f"        print('窗口已移动到 ({window_x}, {window_y}), 大小 {window_width}x{window_height}')\n"
                    workflow_code += f"    else:\n"
                    workflow_code += f"        print('未找到指定窗口')\n"
            elif step['type'] == "keyboard":
                keys = step.get('keys', '')
                workflow_code += f"    print('执行键盘按键操作: {keys}')\n"
                workflow_code += f"    pyautogui.hotkey(*'{keys}'.split('+'))\n"
                workflow_code += f"    print('按键组合执行成功: {keys}')\n"
            elif step['type'] == "delay":
                duration = step.get('duration', 1)
                workflow_code += f"    print('执行延迟操作: {duration}秒')\n"
                workflow_code += f"    time.sleep({duration})\n"
            elif step['type'] == "clipboard_write":
                text = step.get('text', '')
                workflow_code += f"    print('执行写入剪贴板操作')\n"
                workflow_code += f"    pyperclip.copy(r'''{text}''')\n"
                workflow_code += f"    print('写入剪贴板成功')\n"
            elif step['type'] == "clipboard_read":
                workflow_code += f"    print('执行读取剪贴板操作')\n"
                workflow_code += f"    try:\n"
                workflow_code += f"        clipboard_text = pyperclip.paste()\n"
                workflow_code += f"        if len(clipboard_text) > 20:\n"
                workflow_code += f"            print(f'读取剪贴板成功: {{clipboard_text[:20]}}...')\n"
                workflow_code += f"        else:\n"
                workflow_code += f"            print(f'读取剪贴板成功: {{clipboard_text}}')\n"
                workflow_code += f"    except Exception as e:\n"
                workflow_code += f"        print(f'读取剪贴板失败: {{str(e)}}')\n"
            elif step['type'] == "condition":
                condition_type = step.get('condition_type', 'image')
                accuracy = step.get('accuracy', 0.9)
                found_behavior = step.get('found_behavior', 'continue')
                not_found_behavior = step.get('not_found_behavior', 'jump')
                jump_to = step.get('jump_to', 1)
                
                if condition_type == "image":
                    image_path = step.get('image_path', '')
                    workflow_code += f"    print('执行条件判断操作: 等待图像出现')\n"
                    workflow_code += f"    found = False\n"
                    workflow_code += f"    position = find_image(r'{image_path}', {accuracy})\n"
                    workflow_code += f"    if position:\n"
                    workflow_code += f"        print(f'找到图像，位置: {{position}}')\n"
                    workflow_code += f"        found = True\n"
                    workflow_code += f"    else:\n"
                    workflow_code += f"        print('未找到图像')\n"
                else:  # text
                    text = step.get('text', '')
                    workflow_code += f"    print('执行条件判断操作: 等待文字出现')\n"
                    workflow_code += f"    found = False\n"
                    workflow_code += f"    found = find_text('{text}')\n"
                    workflow_code += f"    if found:\n"
                    workflow_code += f"        print(f'找到文字: {{text}}')\n"
                    workflow_code += f"    else:\n"
                    workflow_code += f"        print(f'未找到文字: {{text}}')\n"
                
                workflow_code += f"    if found:\n"
                workflow_code += f"        print('条件满足')\n"
                if found_behavior == "continue":
                    workflow_code += f"        # 继续执行\n"
                elif found_behavior == "jump":
                    workflow_code += f"        # 跳转到步骤 {jump_to}\n"
                    workflow_code += f"        # 注意：脚本中无法直接实现跳转，需要手动调整\n"
                elif found_behavior == "restart":
                    workflow_code += f"        # 重新开始工作流\n"
                    workflow_code += f"        print('重新开始工作流')\n"
                    workflow_code += f"        continue\n"
                workflow_code += f"    else:\n"
                workflow_code += f"        print('条件不满足')\n"
                if not_found_behavior == "jump":
                    workflow_code += f"        # 跳转到步骤 {jump_to}\n"
                    workflow_code += f"        # 注意：脚本中无法直接实现跳转，需要手动调整\n"
                else:  # wait
                    workflow_code += f"        # 继续找，直到找到为止\n"
                    workflow_code += f"        start_time = time.time()\n"
                    workflow_code += f"        timeout = 60  # 60秒超时\n"
                    workflow_code += f"        while not found:\n"
                    if condition_type == "image":
                        workflow_code += f"            position = find_image(r'{image_path}', {accuracy})\n"
                        workflow_code += f"            if position:\n"
                        workflow_code += f"                print(f'找到图像，位置: {{position}}')\n"
                        workflow_code += f"                found = True\n"
                    else:  # text
                        workflow_code += f"            found = find_text('{text}')\n"
                        workflow_code += f"            if found:\n"
                        workflow_code += f"                print(f'找到文字: {{text}}')\n"
                    workflow_code += f"            if not found:\n"
                    workflow_code += f"                if time.time() - start_time > timeout:\n"
                    workflow_code += f"                    print(f'超时: 未找到目标，超过 {{timeout}} 秒')\n"
                    workflow_code += f"                    break\n"
                    workflow_code += f"                time.sleep(2)\n"
                    workflow_code += f"                print('继续找，直到找到为止...')\n"
            
            # 添加1秒等待
            workflow_code += f"    time.sleep(1)  # 等待1秒\n\n"
        
        # 替换时间戳
        import datetime
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        script_content = script_content.format(timestamp=timestamp, workflow_code=workflow_code)
        
        # 保存脚本文件
        file_path = filedialog.asksaveasfilename(
            defaultextension=".py",
            filetypes=[("Python files", "*.py")],
            title="保存生成的脚本"
        )
        
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(script_content)
            
            messagebox.showinfo("成功", f"脚本生成成功，保存到: {file_path}")
        else:
            messagebox.showinfo("取消", "脚本生成已取消")

    def load_actions(self):
        if os.path.exists("actions.json"):
            try:
                with open("actions.json", "r", encoding="utf-8") as f:
                    self.actions = json.load(f)

                # 更新操作计数器
                if self.actions:
                    ids = [int(a["id"].split("_")[1]) for a in self.actions if a["id"].startswith("action_")]
                    if ids:
                        self.action_counter = max(ids) + 1

                # 更新操作列表
                for action in self.actions:
                    if action["type"] == "mouse":
                        details = f"截图: {os.path.basename(action['image_path'])}"
                        action_type = "鼠标点击"
                    elif action["type"] == "keyboard":
                        details = f"按键: {action['keys']}"
                        action_type = "键盘按键"
                    elif action["type"] == "clipboard_write":
                        text = action.get('text', '')
                        details = f"文字: {text[:20]}..." if len(text) > 20 else f"文字: {text}"
                        action_type = "写入剪贴板"
                    elif action["type"] == "clipboard_read":
                        details = "读取剪贴板内容"
                        action_type = "读取剪贴板"
                    elif action["type"] == "condition":
                        condition_type = action.get("condition_type", "image")
                        if condition_type == "image":
                            details = f"判断图片: {os.path.basename(action['image_path'])}"
                        else:
                            details = f"判断文字: {action.get('text', '')}"
                        action_type = "条件判断"
                    else:
                        details = "未知操作"
                        action_type = "未知"

                    self.action_tree.insert("", tk.END, values=(action["id"], action["name"], action_type, details))

                self.update_available_actions()
            except Exception as e:
                print(f"加载操作失败: {str(e)}")


class SetWindowToolDialog:
    def __init__(self, parent, title_var, class_var, x_var, y_var, width_var, height_var):
        self.parent = parent
        self.title_var = title_var
        self.class_var = class_var
        self.x_var = x_var
        self.y_var = y_var
        self.width_var = width_var
        self.height_var = height_var
        self.selected_hwnd = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("窗体工具 - 选择窗口")
        self.dialog.geometry("600x450")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.create_widgets()
        self.refresh_window_list()
        
        self.dialog.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="当前桌面窗口列表:", font=("", 10, "bold")).pack(anchor=tk.W)
        
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        columns = ("title", "class_name", "pid", "exe")
        self.window_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)
        self.window_tree.heading("title", text="窗口标题")
        self.window_tree.heading("class_name", text="类名")
        self.window_tree.heading("pid", text="PID")
        self.window_tree.heading("exe", text="进程名")
        
        self.window_tree.column("title", width=200)
        self.window_tree.column("class_name", width=120)
        self.window_tree.column("pid", width=60)
        self.window_tree.column("exe", width=120)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.window_tree.yview)
        self.window_tree.configure(yscrollcommand=scrollbar.set)
        
        self.window_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.window_tree.bind("<<TreeviewSelect>>", self.on_select_window)
        self.window_tree.bind("<Double-1>", self.on_double_click)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="刷新列表", command=self.refresh_window_list).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="定位窗口", command=self.locate_window).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="使用选中窗口", command=self.use_selected_window).pack(side=tk.LEFT, padx=5)
        
        detail_frame = ttk.LabelFrame(main_frame, text="选中窗口详情", padding=10)
        detail_frame.pack(fill=tk.X, pady=5)
        
        self.detail_text = tk.Text(detail_frame, height=5, width=70)
        self.detail_text.pack(fill=tk.X)
        self.detail_text.config(state=tk.DISABLED)
        
        tip_label = ttk.Label(main_frame, text="提示: 双击窗口行可直接使用该窗口信息", foreground="gray")
        tip_label.pack(anchor=tk.W, pady=5)
    
    def refresh_window_list(self):
        for item in self.window_tree.get_children():
            self.window_tree.delete(item)
        
        def enum_windows_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:
                    class_name = win32gui.GetClassName(hwnd)
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    try:
                        process = psutil.Process(pid)
                        exe_name = process.name()
                    except:
                        exe_name = ""
                    
                    results.append({
                        "hwnd": hwnd,
                        "title": title,
                        "class_name": class_name,
                        "pid": pid,
                        "exe_name": exe_name
                    })
        
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        for w in windows:
            self.window_tree.insert("", tk.END, values=(
                w["title"],
                w["class_name"],
                str(w["pid"]),
                w["exe_name"]
            ), tags=(str(w["hwnd"]), w["class_name"]))
    
    def on_select_window(self, event):
        selection = self.window_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        tags = self.window_tree.item(item, "tags")
        if tags:
            self.selected_hwnd = int(tags[0])
            self.update_detail()
    
    def update_detail(self):
        if not self.selected_hwnd:
            return
        
        try:
            title = win32gui.GetWindowText(self.selected_hwnd)
            class_name = win32gui.GetClassName(self.selected_hwnd)
            rect = win32gui.GetWindowRect(self.selected_hwnd)
            _, pid = win32process.GetWindowThreadProcessId(self.selected_hwnd)
            
            detail = f"窗口标题: {title}\n"
            detail += f"窗口类名: {class_name}\n"
            detail += f"窗口位置: ({rect[0]}, {rect[1]}) - ({rect[2]}, {rect[3]})\n"
            detail += f"窗口大小: {rect[2]-rect[0]} x {rect[3]-rect[1]}\n"
            detail += f"进程ID: {pid}"
            
            self.detail_text.config(state=tk.NORMAL)
            self.detail_text.delete(1.0, tk.END)
            self.detail_text.insert(tk.END, detail)
            self.detail_text.config(state=tk.DISABLED)
        except Exception as e:
            self.detail_text.config(state=tk.NORMAL)
            self.detail_text.delete(1.0, tk.END)
            self.detail_text.insert(tk.END, f"获取窗口信息失败: {str(e)}")
            self.detail_text.config(state=tk.DISABLED)
    
    def locate_window(self):
        if not self.selected_hwnd:
            messagebox.showwarning("提示", "请先选择一个窗口")
            return
        
        try:
            win32gui.SetForegroundWindow(self.selected_hwnd)
        except Exception as e:
            messagebox.showerror("错误", f"无法定位窗口: {str(e)}")
    
    def use_selected_window(self):
        if not self.selected_hwnd:
            messagebox.showwarning("提示", "请先选择一个窗口")
            return
        
        try:
            title = win32gui.GetWindowText(self.selected_hwnd)
            class_name = win32gui.GetClassName(self.selected_hwnd)
            rect = win32gui.GetWindowRect(self.selected_hwnd)
            
            self.title_var.set(title)
            self.class_var.set(class_name)
            self.x_var.set(str(rect[0]))
            self.y_var.set(str(rect[1]))
            self.width_var.set(str(rect[2] - rect[0]))
            self.height_var.set(str(rect[3] - rect[1]))
            
            self.dialog.destroy()
        except Exception as e:
            messagebox.showerror("错误", f"获取窗口信息失败: {str(e)}")
    
    def on_double_click(self, event):
        self.use_selected_window()
    
    def on_close(self):
        self.dialog.destroy()


class ScreenshotSelector:
    def __init__(self, parent, path_var, save_dir):
        self.parent = parent
        self.path_var = path_var
        self.save_dir = save_dir
        self.start_x = None
        self.start_y = None
        self.rect_id = None
        self.screenshot = None
        
        self.selector = tk.Toplevel(parent)
        self.selector.attributes("-fullscreen", True)
        self.selector.attributes("-alpha", 0.3)
        self.selector.attributes("-topmost", True)
        self.selector.configure(bg="black")
        self.selector.overrideredirect(True)
        
        self.canvas = tk.Canvas(self.selector, bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        self.screenshot = pyautogui.screenshot()
        
        self.canvas.bind("<Button-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_move)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.canvas.bind("<Escape>", self.on_cancel)
        
        self.selector.focus_set()
        self.selector.bind("<Escape>", self.on_cancel)
        
        tip_label = tk.Label(self.selector, text="拖拽鼠标左键选择截图区域，按 ESC 取消", 
                            bg="yellow", fg="black", font=("", 12))
        tip_label.place(relx=0.5, rely=0.02, anchor=tk.CENTER)
    
    def on_mouse_down(self, event):
        self.start_x = event.x
        self.start_y = event.y
    
    def on_mouse_move(self, event):
        if self.start_x is None:
            return
        
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        
        self.rect_id = self.canvas.create_rectangle(
            self.start_x, self.start_y, event.x, event.y,
            outline="red", width=2, fill=""
        )
    
    def on_mouse_up(self, event):
        if self.start_x is None:
            return
        
        end_x = event.x
        end_y = event.y
        
        left = min(self.start_x, end_x)
        top = min(self.start_y, end_y)
        right = max(self.start_x, end_x)
        bottom = max(self.start_y, end_y)
        
        width = right - left
        height = bottom - top
        
        if width < 5 or height < 5:
            self.on_cancel(None)
            return
        
        self.selector.destroy()
        
        cropped = self.screenshot.crop((left, top, right, bottom))
        
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"screenshot_{timestamp}.png"
        save_path = os.path.join(self.save_dir, filename)
        
        os.makedirs(self.save_dir, exist_ok=True)
        cropped.save(save_path)
        
        relative_path = os.path.relpath(save_path, current_dir)
        self.path_var.set(relative_path)
        
        self.parent.deiconify()
    
    def on_cancel(self, event):
        self.selector.destroy()
        self.parent.deiconify()


class WindowToolDialog:
    def __init__(self, parent, app_path_var):
        self.parent = parent
        self.app_path_var = app_path_var
        self.selected_hwnd = None
        self.refresh_timer = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("窗体工具")
        self.dialog.geometry("600x500")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.create_widgets()
        self.refresh_window_list()
        
        self.dialog.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="当前桌面窗口列表:", font=("", 10, "bold")).pack(anchor=tk.W)
        
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        columns = ("title", "class_name", "pid", "exe")
        self.window_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=12)
        self.window_tree.heading("title", text="窗口标题")
        self.window_tree.heading("class_name", text="类名")
        self.window_tree.heading("pid", text="PID")
        self.window_tree.heading("exe", text="进程名")
        
        self.window_tree.column("title", width=200)
        self.window_tree.column("class_name", width=120)
        self.window_tree.column("pid", width=60)
        self.window_tree.column("exe", width=120)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.window_tree.yview)
        self.window_tree.configure(yscrollcommand=scrollbar.set)
        
        self.window_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.window_tree.bind("<<TreeviewSelect>>", self.on_select_window)
        self.window_tree.bind("<Double-1>", self.on_double_click)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="刷新列表", command=self.refresh_window_list).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="定位窗口", command=self.locate_window).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="使用选中进程", command=self.use_selected_exe).pack(side=tk.LEFT, padx=5)
        
        detail_frame = ttk.LabelFrame(main_frame, text="选中窗口详情", padding=10)
        detail_frame.pack(fill=tk.X, pady=5)
        
        self.detail_text = tk.Text(detail_frame, height=6, width=70)
        self.detail_text.pack(fill=tk.X)
        self.detail_text.config(state=tk.DISABLED)
        
        tip_label = ttk.Label(main_frame, text="提示: 双击窗口行可直接使用该进程路径", foreground="gray")
        tip_label.pack(anchor=tk.W, pady=5)
    
    def refresh_window_list(self):
        for item in self.window_tree.get_children():
            self.window_tree.delete(item)
        
        def enum_windows_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:
                    class_name = win32gui.GetClassName(hwnd)
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    try:
                        process = psutil.Process(pid)
                        exe_name = process.name()
                        exe_path = process.exe()
                    except:
                        exe_name = ""
                        exe_path = ""
                    
                    results.append({
                        "hwnd": hwnd,
                        "title": title,
                        "class_name": class_name,
                        "pid": pid,
                        "exe_name": exe_name,
                        "exe_path": exe_path
                    })
        
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        for w in windows:
            self.window_tree.insert("", tk.END, values=(
                w["title"],
                w["class_name"],
                str(w["pid"]),
                w["exe_name"]
            ), tags=(str(w["hwnd"]), w["exe_path"]))
    
    def on_select_window(self, event):
        selection = self.window_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        tags = self.window_tree.item(item, "tags")
        if tags:
            self.selected_hwnd = int(tags[0])
            exe_path = tags[1] if len(tags) > 1 else ""
            self.update_detail(exe_path)
    
    def update_detail(self, exe_path):
        if not self.selected_hwnd:
            return
        
        try:
            title = win32gui.GetWindowText(self.selected_hwnd)
            class_name = win32gui.GetClassName(self.selected_hwnd)
            rect = win32gui.GetWindowRect(self.selected_hwnd)
            _, pid = win32process.GetWindowThreadProcessId(self.selected_hwnd)
            process = psutil.Process(pid)
            
            detail = f"窗口标题: {title}\n"
            detail += f"窗口类名: {class_name}\n"
            detail += f"窗口位置: ({rect[0]}, {rect[1]}) - ({rect[2]}, {rect[3]})\n"
            detail += f"窗口大小: {rect[2]-rect[0]} x {rect[3]-rect[1]}\n"
            detail += f"进程ID: {pid}\n"
            detail += f"进程路径: {exe_path}"
            
            self.detail_text.config(state=tk.NORMAL)
            self.detail_text.delete(1.0, tk.END)
            self.detail_text.insert(tk.END, detail)
            self.detail_text.config(state=tk.DISABLED)
        except Exception as e:
            self.detail_text.config(state=tk.NORMAL)
            self.detail_text.delete(1.0, tk.END)
            self.detail_text.insert(tk.END, f"获取窗口信息失败: {str(e)}")
            self.detail_text.config(state=tk.DISABLED)
    
    def locate_window(self):
        if not self.selected_hwnd:
            messagebox.showwarning("提示", "请先选择一个窗口")
            return
        
        try:
            win32gui.SetForegroundWindow(self.selected_hwnd)
        except Exception as e:
            messagebox.showerror("错误", f"无法定位窗口: {str(e)}")
    
    def use_selected_exe(self):
        selection = self.window_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择一个窗口")
            return
        
        item = selection[0]
        tags = self.window_tree.item(item, "tags")
        if tags and len(tags) > 1:
            exe_path = tags[1]
            if exe_path:
                self.app_path_var.set(exe_path)
                self.dialog.destroy()
            else:
                messagebox.showwarning("提示", "无法获取该窗口的进程路径")
    
    def on_double_click(self, event):
        self.use_selected_exe()
    
    def on_close(self):
        self.dialog.destroy()


if __name__ == "__main__":
    # 检查并安装依赖
    try:
        import cv2
        import numpy as np
        import pyautogui
    except ImportError:
        import subprocess
        import sys
        print("正在安装依赖...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "opencv-python", "numpy", "pyautogui"])
        print("依赖安装完成")
        import cv2
        import numpy as np
        import pyautogui

    root = tk.Tk()
    app = WorkflowAutomationTool(root)
    root.mainloop()