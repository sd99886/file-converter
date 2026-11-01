#!/usr/bin/env bash
# 启动图形界面（无显示则提示使用 CLI）
set -euo pipefail
cd "$(dirname "$0")"

if [[ ! -d .venv ]]; then
  echo "正在创建虚拟环境并安装依赖…"
  python3 -m venv .venv
  # shellcheck disable=SC1091
  source .venv/bin/activate
  pip install -U pip
  pip install -r requirements.txt
else
  # shellcheck disable=SC1091
  source .venv/bin/activate
  # 补装可能新增的依赖
  pip install -q -r requirements.txt
fi

if [[ -n "${DISPLAY:-}" ]] || [[ "$(uname -s)" == "Darwin" ]]; then
  if python -c "import tkinter" 2>/dev/null; then
    exec python gui.py
  else
    echo "缺少 tkinter。Linux 请安装: sudo apt install python3-tk"
    echo "或使用命令行，例如:"
    echo "  python convert.py a.png b.jpg --to pdf --merge -o ./out --name merged.pdf"
    exit 1
  fi
else
  echo "未检测到图形界面。请使用命令行，例如："
  echo "  source .venv/bin/activate"
  echo "  python convert.py a.png b.jpg --to pdf --merge -o ./out --name merged.pdf"
  echo "  python convert.py file.pdf --to png -o ./pages/"
  exec bash
fi
