@echo off
chcp 65001 >nul 2>&1
set PATH=C:\Users\xiaow\AppData\Local\Programs\Python\Python311;%PATH%
cd /d "%~dp0"
echo ========================================
echo   合同审核工具 v3 — 后端服务
echo ========================================
echo.
echo 后端: http://127.0.0.1:5577
echo 前端: index.html (双击打开即可)
echo.
echo L2 模式: 后端在线 → PDF/DOC/DOCX/OCR
echo L0 模式: 后端离线 → 仅 DOCX + AI
echo.
echo 端点:
echo   POST /audit     — 完整审核(规则+条款)
echo   POST /extract   — 仅提取文本
echo   GET  /rules     — 获取规则配置
echo   POST /rules     — 保存规则配置
echo   GET  /health    — 健康检查
echo.
echo 按 Ctrl+C 停止服务
echo ========================================
echo.
python contract_audit_service.py
pause
