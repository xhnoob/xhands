@echo off
chcp 65001 >nul
echo ========================================
echo   xHands 打包脚本
echo ========================================
echo.

echo 检查 PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo 正在安装 PyInstaller...
    pip install pyinstaller
)

echo.
echo 开始打包...
echo.

pyinstaller --clean xhands.spec

if errorlevel 1 (
    echo.
    echo ========================================
    echo   打包失败！
    echo ========================================
    pause
    exit /b 1
)

echo.
echo ========================================
echo   打包完成！
echo ========================================
echo.
echo 输出文件: dist\xhands.exe
echo.

if exist "dist\xhands.exe" (
    echo 正在复制必要文件到 dist 目录...
    if not exist "dist\screenShot" mkdir "dist\screenShot"
    if not exist "dist\flowScripts" mkdir "dist\flowScripts"
    if not exist "dist\flowPY" mkdir "dist\flowPY"
    xcopy /Y /E /I screenShot dist\screenShot >nul 2>&1
    xcopy /Y /E /I flowScripts dist\flowScripts >nul 2>&1
    xcopy /Y /E /I flowPY dist\flowPY >nul 2>&1
    copy /Y config.json dist\ >nul 2>&1
    echo 文件复制完成！
)

echo.
pause
