#!/usr/bin/env bash
set -euo pipefail

PYTHON="/Users/mac/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"

if [ ! -x "$PYTHON" ]; then
  echo "未找到 Codex 内置 Python：$PYTHON"
  echo "请改用安装了 openpyxl 的 Python 运行：python3 app.py"
  exit 1
fi

cd "$(dirname "$0")"
exec "$PYTHON" app.py "$@"
