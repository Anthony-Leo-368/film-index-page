#!/bin/bash
cd "$(dirname "$0")" || exit 1
echo "==============================================================="
echo "  重新生成 assets/snapshot.json（含各卷图片清单）"
echo "  只在「文件模式」（直接双击 index.html）下才需要；"
echo "  用「010_启动Mac.command」的服务模式时是自动更新的。"
echo "==============================================================="
echo
/usr/bin/python3 assets/worker.py --build
echo
read -n 1 -s -r -p "按任意键关闭此窗口…"
