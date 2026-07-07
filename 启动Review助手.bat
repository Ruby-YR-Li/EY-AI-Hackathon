@echo off
chcp 65001 >nul
cd /d "E:\AI现场比赛"
echo.
echo ========================================
echo   费用底稿 AI Review 助手
echo ========================================
echo.
echo 正在启动 Streamlit 服务...
echo 浏览器打开后即可上传底稿开始 Review
echo 关闭此窗口即可停止服务
echo.
D:\Python312\python.exe -m streamlit run src/app.py
pause
