@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   审计底稿复核系统 - 启动中...
echo ========================================

:: 1. 停掉旧进程（占用8000端口的）
echo [1/3] 检查并停止旧进程...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8000.*LISTENING" 2^>nul') do (
    echo   停止进程 %%a
    taskkill /f /pid %%a >nul 2>&1
)

:: 2. 清理缓存，确保加载最新代码
echo [2/3] 清理缓存...
for /d /r "%~dp0" %%d in (__pycache__) do @if exist "%%d" rd /s /q "%%d" 2>nul
del /s /q "%~dp0*.pyc" 2>nul
del /s /q "%~dp0.cache\*.json" 2>nul
echo   缓存已清理（含Python缓存和AI缓存）

:: 3. 启动服务（每次检查时自动重载最新代码）
echo [3/3] 启动Web服务...
echo ========================================
echo   打开浏览器访问: http://localhost:8000
echo   每次上传底稿检查时自动重载最新代码
echo   按 Ctrl+C 可停止服务
echo ========================================
echo.

start "" http://localhost:8000
python web\app.py
pause
