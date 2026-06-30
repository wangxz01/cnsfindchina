#!/usr/bin/env bash
# 一键环境准备脚本（macOS / Linux / Git Bash on Windows）
# 用法：bash setup.sh
set -e

echo "========================================="
echo " CNS 期刊爬虫 · 环境准备"
echo "========================================="

# 1. Python 版本检查
if ! command -v python &> /dev/null; then
    echo "[error] 未找到 python，请先装 Python 3.10+"
    exit 1
fi
PY_VER=$(python -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
echo "[1/4] Python 版本: $PY_VER"

# 2. 安装依赖
echo ""
echo "[2/4] 安装 Python 依赖 (requirements.txt)..."
pip install -r requirements.txt

# 3. 安装 Playwright Chromium
echo ""
echo "[3/4] 安装 Playwright Chromium 浏览器（约 150MB，首次较慢）..."
python -m playwright install chromium

# 4. 创建必要目录（运行时也会自动建，这里显式建方便排查）
echo ""
echo "[4/4] 创建工作目录..."
mkdir -p browser_profile browser_profile_nature browser_profile_science
mkdir -p cache cache_nature cache_science
mkdir -p static
echo "  ✓ browser_profile*/ cache*/"

echo ""
echo "========================================="
echo " 完成！"
echo "========================================="
echo ""
echo "下一步："
echo "  python unified_web.py        # 统一服务 http://127.0.0.1:8000/"
echo "  python scraper.py            # Cell CLI"
echo "  python nature_scraper.py     # Nature CLI"
echo "  python science_scraper.py    # Science CLI"
echo ""
echo "若从 GitHub 迁移：git clone https://github.com/wangxz01/cnsfindchina.git"
