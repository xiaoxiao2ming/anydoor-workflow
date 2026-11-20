@echo off
rem ── PDF转Word 节点引擎一键安装（百度开源 PaddleOCR）──────────────
rem 本脚本装的是本机 Python 环境（不进软件安装包）。
rem 要求：本机已装 Python 3.9~3.12（python.org 下载，勾选 Add to PATH）。
chcp 65001 >nul
echo [1/3] 检查 Python...
python --version || (echo 未找到 python——请先安装 Python 3.9-3.12 并勾选 Add to PATH & pause & exit /b 1)
echo [2/3] 安装 paddlepaddle + paddleocr（约 600MB，耐心等待）...
python -m pip install -U pip
python -m pip install paddlepaddle paddleocr
if errorlevel 1 (echo 安装失败——检查网络/代理后重试 & pause & exit /b 1)
echo [3/3] 验证引擎...
python -c "from paddleocr import PPStructureV3; print('engine ok')"
if errorlevel 1 (echo 验证失败 & pause & exit /b 1)
echo.
echo 完成！回到软件 PDF转Word 节点点「W」执行转换。
echo 注意：首次转换会自动下载模型（约 1-2GB，只下一次，之后离线可用）。
pause
