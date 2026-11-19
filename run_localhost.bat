@echo off
chcp 65001 > nul
echo =======================================================================
echo   AI301 - ĐỀ DỰ ÁN 09: AI AGENT PHÂN TÍCH TIẾN ĐỘ & DỰ KIẾN DOANH SỐ
echo   Công ty Bình Minh (B2B Tech Services)
echo =======================================================================
echo.
echo Đang khởi chạy AI Agent trên localhost: http://localhost:8000 ...
echo Nhấn Ctrl+C để dừng server.
echo.
start http://localhost:8000
python -m uvicorn app:app --host 127.0.0.1 --port 8000 --reload
pause

