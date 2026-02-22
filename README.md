# xHands - 工作流自动化工具

![Python](https://img.shields.io/badge/python-3.8%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Platform](https://img.shields.io/badge/platform-Windows-blue)

xHands 是一款功能强大的工作流自动化工具，支持图像识别、OCR文字识别，可与飞书协同工作，实现智能化的自动化操作。

## 功能特性

- **智能识别** - 图像识别 + OCR文字识别，精准定位目标
- **多维度操作** - 鼠标点击、键盘输入、剪贴板操作、系统命令
- **条件分支** - 基于识别结果的智能判断、跳转和重试
- **脚本导出** - 一键生成可独立运行的 Python 脚本
- **双模式运行** - GUI 图形界面 + CLI 命令行
- **飞书协同** - 与飞书机器人集成，实现消息推送

## 快速开始

### 方式一：直接使用 EXE（推荐）

下载 `dist` 目录，双击 `xhands.exe` 即可启动 GUI 界面。

```
dist/
├── xhands.exe        # 主程序
├── config.json       # 配置文件
├── screenShot/       # 截图资源
├── flowScripts/      # 工作流定义
└── flowPY/           # 导出脚本
```

### 方式二：从源码运行(直接使用EXE不用看这个)

```bash
# 克隆项目
git clone https://github.com/your-repo/xHands.git
cd xHands

# 安装依赖
pip install opencv-python numpy pyautogui pyperclip pytesseract Pillow psutil pywin32

# 运行
python xHands.py
```

### 可选：安装 Tesseract OCR(直接使用EXE不用看这个)

从 [GitHub](https://github.com/UB-Mannheim/tesseract/wiki) 下载安装，用于文字识别功能。

## 使用方式

### GUI 模式

双击 `xhands.exe` 或执行：

```bash
xhands.exe gui
```

### CLI 模式

在终端中运行：

```bash
xhands.exe                    # 进入交互模式
xhands.exe list               # 列出所有工作流
xhands.exe show <name>        # 显示工作流详情
xhands.exe run <name>         # 执行工作流
xhands.exe run <name> -c 5    # 执行5次
xhands.exe run <name> -d 10   # 执行10分钟
xhands.exe run <name> -r      # 错误时重试
xhands.exe exec <script.py>   # 执行脚本
```

### CLI 参数说明

| 参数 | 说明 |
|------|------|
| `-c, --count` | 执行次数（0=无限循环） |
| `-d, --duration` | 执行时长（分钟，0=不限制） |
| `-i, --interval` | 执行间隔（秒） |
| `-r, --retry` | 出现错误时重试 |
| `-q, --quiet` | 静默模式 |

## 操作指南

### 基本流程

1. **创建操作** - 在「操作管理」中创建鼠标、键盘、延迟等操作
2. **构建工作流** - 在「工作流编辑」中组合操作，添加条件和循环
3. **执行任务** - 在「工作流执行」中选择工作流运行

### 操作类型

| 类型 | 说明 |
|------|------|
| 鼠标点击 | 图像/文字识别定位后点击 |
| 键盘按键 | 模拟键盘按键组合 |
| 延迟等待 | 固定延时 |
| 写入剪贴板 | 复制文本到剪贴板 |
| 读取剪贴板 | 读取剪贴板内容 |
| 条件判断 | 基于图像/文字识别的分支跳转 |
| 系统命令 | 启动应用、固定窗体位置 |

### 快捷键

- `F8` - 开始执行工作流
- `F9` - 停止执行

## 飞书协同

xHands 支持与飞书机器人集成，通过 OpenClaw 实现消息推送：

1. **安装 OpenClaw** - 确保 OpenClaw CLI 已安装
2. **获取飞书ID** - 向飞书机器人发送：「我的飞书ID多少」
3. **配置连接** - 在设置页面填入飞书ID并保存
4. **测试连接** - 点击「连接测试」验证

```bash
# 发送消息示例
openclaw message send --channel feishu --account main --target "your_id" --message "任务完成！"
```

## 自行打包

如需自行打包，执行：

```bash
# 安装 PyInstaller
pip install pyinstaller

# 执行打包
build.bat

# 或手动执行
python -m PyInstaller --clean xhands.spec
```

## 项目结构

```
xHands/
├── xHands.py           # GUI 主程序
├── xHands_cli.py       # CLI 命令行程序
├── xhands_main.py      # 统一入口
├── xhands.spec         # PyInstaller 配置
├── build.bat           # 打包脚本
├── config.json         # 配置文件
├── flowScripts/        # 工作流定义
│   └── *.json
├── flowPY/             # 导出的脚本
│   └── *.py
├── screenShot/         # 截图资源
│   └── *.png
└── dist/               # 打包输出
    └── xhands.exe
```

## 许可证

本项目采用 MIT 许可证。

---

**作者**: 抖音 @黑化的超级奶爸
