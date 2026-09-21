#!/usr/bin/env bash
# 回归测试：把 example/sample.md 跑遍全部参数组合，靠 verify.py 断言。
# 用法:  ./example/run_tests.sh  [额外的 md 文件 ...]
#
# sample.md 刻意覆盖了每一类已知陷阱，任何一处回归都会让 verify.py 非零退出。
set -u
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$HERE"

DOCS=("example/sample.md" "$@")
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

# BOM + CRLF 变体：Windows 编辑器产物的典型形态。BOM 曾让首个标题匹配不上
# 行首正则（chapter 编号静默错乱），必须有回归保护。
CRLF="$TMP/sample_crlf_bom.md"
python3 - "$CRLF" <<'EOF'
import sys
raw = open('example/sample.md', encoding='utf-8').read().replace('\n', '\r\n')
open(sys.argv[1], 'w', encoding='utf-8-sig', newline='').write(raw)
EOF
DOCS+=("$CRLF")

ok=0; fail=0
for doc in "${DOCS[@]}"; do
  name="$(basename "$doc")"
  for mm in omml latex; do
    for ns in plain chapter; do
      for nb in all tag none; do
        label="$name $mm/$ns/$nb"
        r=$(python3 md2docx.py "$doc" -o "$TMP/out.docx" \
              --math-mode "$mm" --number-style "$ns" --number "$nb" 2>&1)
        if echo "$r" | grep -q "全部检查通过"; then
          printf "  PASS  %-34s %s\n" "$label" "$(echo "$r" | grep '^\[verify\] 块')"
          ok=$((ok + 1))
        else
          printf "  FAIL  %-34s\n" "$label"
          echo "$r" | grep -E "失败|Traceback|Error" | sed 's/^/          /'
          fail=$((fail + 1))
        fi
      done
    done
  done
done

echo
echo "通过 $ok / 失败 $fail"
[[ $fail -eq 0 ]]
