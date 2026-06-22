#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"
PYTHON_BIN="${PYTHON_BIN:-/Users/weichang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3}"

for file in data/HR台账模板_员工花名册.xlsx data/HR台账模板_薪资数据.xlsx data/HR台账模板_绩效数据.xlsx; do
  if [[ ! -f "$file" ]]; then
    echo "缺少文件：$file"
    exit 1
  fi
done

echo "1/2 正在读取 Excel 并生成网页..."
"$PYTHON_BIN" scripts/generate_dashboard.py
echo "2/2 正在发布到 Cloudflare Pages..."
./deploy.sh
echo "完成：https://hr-mark-ledger-dashboard.pages.dev"
