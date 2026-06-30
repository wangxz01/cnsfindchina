@echo off
REM 一键环境准备脚本（Windows cmd / PowerShell）
REM 用法：双击或 setup.bat
setlocal enabledelayedexpansion

echo =========================================
echo  CNS 期刊爬虫 · 环境准备
echo =========================================

REM 1. Python 检查
where python >nul 2>&1
if errorlevel 1 (
    echo [error] 未找到 python，请先装 Python 3.10+（建议勾选 Add to PATH）
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"') do set PY_VER=%%v
echo [1/4] Python 版本: !PY_VER!

REM 2. 安装依赖
echo.
echo [2/4] 安装 Python 依赖 ^(requirements.txt^)...
pip install -r requirements.txt
if errorlevel 1 (
    echo [error] pip 安装失败
    pause
    exit /b 1
)

REM 3. 安装 Playwright Chromium
echo.
echo [3/4] 安装 Playwright Chromium 浏览器（约 150MB，首次较慢）...
python -m playwright install chromium
if errorlevel 1 (
    echo [error] Playwright 浏览器安装失败
    pause
    exit /b 1
)

REM 4. 创建必要目录
echo.
echo [4/4] 创建工作目录...
if not exist browser_profile mkdir browser_profile
if not exist browser_profile_nature mkdir browser_profile_nature
if not exist browser_profile_science mkdir browser_profile_science
if not exist cache mkdir cache
if not exist cache_nature mkdir cache_nature
if not exist cache_science mkdir cache_science
if not exist static mkdir static
echo   √ browser_profile*\ cache*\

echo.
echo =========================================
echo  完成！
echo =========================================
echo.
echo 下一步：
echo   python unified_web.py        REM 统一服务 http://127.0.0.1:8000/
echo   python scraper.py            REM Cell CLI
echo   python nature_scraper.py     REM Nature CLI
echo   python science_scraper.py    REM Science CLI
echo.
pause
