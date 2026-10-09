#!/bin/bash
# 把硬盘上现用版的代码同步到本仓库文件夹（不含任何个人数据）
set -e
SRC='/Volumes/Port_HDD/010_MediaStorage/FILM/000_胶卷索引信息'
DST="$(cd "$(dirname "$0")" && pwd)"
echo '========================================'
echo ' 同步代码到仓库：'
echo "   来源: $SRC"
echo "   仓库: $DST"
echo '========================================'
if [ ! -d "$SRC" ]; then echo '❌ 找不到硬盘上的软件目录（硬盘没插？）'; exit 1; fi
for f in 010_启动Mac.command 011_启动Win.bat 030_手动更新索引.command 031_手动更新索引.bat 使用说明.txt; do
  [ -f "$SRC/$f" ] && cp -p "$SRC/$f" "$DST/$f"
done
for f in worker.py build_html.py index.html template.xlsx index_template.docx; do
  [ -f "$SRC/assets/$f" ] && cp -p "$SRC/assets/$f" "$DST/assets/$f"
done
for d in film_icons film_icons_hd; do
  [ -d "$SRC/assets/$d" ] && rsync -a --delete "$SRC/assets/$d/" "$DST/assets/$d/"
done
echo
echo '✅ 同步完成。个人数据已被 .gitignore 排除，不会进仓库。'
echo
echo '下一步：打开 GitHub Desktop → 填 Summary → Commit to main → Push origin'
echo
read -n 1 -s -r -p '按任意键关闭…'
