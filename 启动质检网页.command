#!/bin/zsh
set -e

cd "$(dirname "$0")"

echo "========================================"
echo "  审计底稿复核系统 - 启动中..."
echo "========================================"

if [ ! -x ".venv/bin/python" ]; then
  echo "未找到项目虚拟环境：.venv"
  echo "请先在项目根目录执行：python3.12 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt"
  read -r "?按回车退出..."
  exit 1
fi

export QUALITY_AGENT_DATA_DIR="${QUALITY_AGENT_DATA_DIR:-$HOME/QualityAgent-dev}"

echo "[1/3] 检查并停止旧服务..."
lsof -ti tcp:8000 | xargs kill -TERM 2>/dev/null || true

echo "[2/3] 清理 Python 缓存..."
find . -name "__pycache__" -type d -prune -exec rm -rf {} + 2>/dev/null || true
find . -name "*.pyc" -type f -delete 2>/dev/null || true
echo "  缓存已清理"

echo "[3/3] 启动 Web 服务..."
echo "========================================"
echo "  打开浏览器访问: http://127.0.0.1:8000"
echo "  按 Ctrl+C 可停止服务"
echo "========================================"
echo

(sleep 2; open "http://127.0.0.1:8000") &
".venv/bin/python" run_server.py
