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

def is_running_from_terminal():
    if len(sys.argv) > 1:
        return True
    
    if sys.stdout.isatty():
        return True
    
    if hasattr(sys, 'frozen'):
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            console_hwnd = kernel32.GetConsoleWindow()
            if console_hwnd == 0:
                return False
            
            parent_pid = ctypes.windll.kernel32.GetProcessId(
                ctypes.windll.kernel32.GetParentProcess(
                    ctypes.windll.kernel32.GetCurrentProcess()
                )
            )
            
            if parent_pid > 0:
                return True
        except:
            pass
    
    return False

def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    if hasattr(sys, 'frozen'):
        current_dir = os.path.dirname(sys.executable)
    
    if is_running_from_terminal():
        from xHands_cli import main as cli_main
        cli_main()
    else:
        import tkinter as tk
        from xHands import WorkflowAutomationTool
        root = tk.Tk()
        app = WorkflowAutomationTool(root)
        root.mainloop()

if __name__ == "__main__":
    main()
