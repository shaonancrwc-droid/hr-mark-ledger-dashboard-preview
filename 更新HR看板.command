#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"
clear
echo "HR全景看板正在更新..."
echo
echo "请确认 data 文件夹里的三张 Excel 已经是最新版本。"
echo "网页地址固定不变：https://hr-mark-ledger-dashboard.pages.dev"
echo

if ./update_dashboard.sh; then
  echo
  echo "更新完成。已经发布到固定网页地址："
  echo "https://hr-mark-ledger-dashboard.pages.dev"
  open "https://hr-mark-ledger-dashboard.pages.dev" >/dev/null 2>&1 || true
else
  echo
  echo "更新失败。请把这个窗口里的报错截图发给维护同事。"
fi

echo
echo "按回车键关闭窗口。"
read -r _
