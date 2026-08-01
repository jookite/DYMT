@echo off
chcp 65001 >nul
title 한글 분할 인쇄 서버
cd /d "%~dp0"
echo 인쇄 서버를 시작합니다. 이 창을 닫으면 서버가 꺼집니다.
python HWP_print_web.py
pause
