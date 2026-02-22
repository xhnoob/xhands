#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xHands - 工作流自动化工具
统一入口：自动检测运行模式
- 终端运行 -> CLI 模式
- 双击运行 -> GUI 模式
"""

import sys
import os

def is_running_from_explorer():
    """
    检测是否从资源管理器双击启动
    """
    if len(sys.argv) > 1:
        return False
    
    if hasattr(sys, 'frozen'):
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            
            console_hwnd = kernel32.GetConsoleWindow()
            if console_hwnd == 0:
                return True
            
            import psutil
            current_pid = os.getpid()
            try:
                current_process = psutil.Process(current_pid)
                parent = current_process.parent()
                if parent:
                    parent_name = parent.name().lower()
                    if parent_name == 'explorer.exe':
                        return True
                    if parent_name in ['cmd.exe', 'powershell.exe', 'pwsh.exe', 'windowsterminal.exe']:
                        return False
            except:
                pass
            
            process_list = (ctypes.c_ulong * 256)()
            process_count = kernel32.GetConsoleProcessList(process_list, 256)
            
            if process_count == 1:
                return True
            
        except Exception:
            pass
    
    if hasattr(sys, 'frozen'):
        return True
    
    if sys.stdout.isatty():
        return False
    
    return True

def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    if hasattr(sys, 'frozen'):
        current_dir = os.path.dirname(sys.executable)
    
    os.chdir(current_dir)
    
    if is_running_from_explorer():
        import tkinter as tk
        from xHands import WorkflowAutomationTool
        root = tk.Tk()
        app = WorkflowAutomationTool(root)
        root.mainloop()
    else:
        from xHands_cli import main as cli_main
        cli_main()

if __name__ == "__main__":
    main()
