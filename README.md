# HR 全景看板维护包

固定网页：

https://hr-mark-ledger-dashboard.pages.dev

## 给 HR 同事的日常流程

1. 打开 `data` 文件夹。
2. 用最新数据覆盖三张 Excel，文件名不要改。
3. 双击 `更新HR看板.command`。

更新成功后，飞书里的网页标签会自动显示最新内容。

## 仓库说明

- `scripts/generate_dashboard.py`：读取 Excel 并生成网页。
- `templates/dashboard-template.html`：看板模板。
- `public/index.html`：Cloudflare Pages 实际发布内容。
- `deploy.sh` / `update_dashboard.sh` / `更新HR看板.command`：更新发布入口。
- `data/*.xlsx`：本地敏感数据，不提交到 Git。
