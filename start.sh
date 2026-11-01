#!/usr/bin/env bash
# 准备虚拟环境并启动图形界面。没有桌面时提示改用命令行。
set -euo pipefail
cd "$(dirname "$0")"

if command -v python3 >/dev/null 2>&1; then
  PY=python3
elif command -v python >/dev/null 2>&1; then
  PY=python
else
  echo "未找到 Python。请安装 Python 3.10 或更高版本。"
  exit 1
fi

if [[ ! -x .venv/bin/python ]]; then
  echo "正在创建虚拟环境并安装依赖…"
  "$PY" -m venv .venv
  .venv/bin/python -m pip install -U pip
  .venv/bin/python -m pip install -r requirements.txt
elif ! .venv/bin/python -c "import fitz, PIL, pptx, natsort, pdf2docx" >/dev/null 2>&1; then
  echo "正在安装缺失依赖…"
  .venv/bin/python -m pip install -U pip
  .venv/bin/python -m pip install -r requirements.txt
fi

if [[ -n "${DISPLAY:-}" || -n "${WAYLAND_DISPLAY:-}" || "$(uname -s)" == "Darwin" ]]; then
  if .venv/bin/python -c "import tkinter" >/dev/null 2>&1; then
    exec .venv/bin/python gui.py
  fi
  echo "缺少 tkinter。Linux 请安装: sudo apt install python3-tk"
  echo "或使用命令行，例如:"
  echo "  .venv/bin/python convert.py a.png b.jpg --to pdf --merge -o ./out --name merged.pdf"
  exit 1
fi

echo "未检测到图形界面。请使用命令行，例如："
echo "  .venv/bin/python convert.py a.png b.jpg --to pdf --merge -o ./out --name merged.pdf"
echo "  .venv/bin/python convert.py file.pdf --to png -o ./pages/"
exit 1
