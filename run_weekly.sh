#!/bin/bash
# launchd 每週排程進入點:先爬取,成功後才分析產報告
set -euo pipefail
cd "$(dirname "$0")"

echo "===== $(date '+%Y-%m-%d %H:%M:%S') 開始爬取 ====="
.venv/bin/python crawler.py  # 爬取天數由 config.json 的 days 控制

echo "===== $(date '+%Y-%m-%d %H:%M:%S') 開始分析 ====="
.venv/bin/python analyze.py

echo "===== $(date '+%Y-%m-%d %H:%M:%S') 完成 ====="
