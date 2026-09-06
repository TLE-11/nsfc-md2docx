#!/usr/bin/env bash
# macOS / Linux 入口。真正的逻辑全在 md2docx.py 里（跨平台），
# 这里只是个薄壳，避免两套实现分叉。
#
#   ./md2docx.sh 本子.md
#   ./md2docx.sh 本子.md -o 本子.docx --number-style chapter
#   ./md2docx.sh 本子.md --math-mode latex
#
# 完整选项见 python3 md2docx.py --help
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
exec python3 "$HERE/md2docx.py" "$@"
