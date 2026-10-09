# -*- coding: utf-8 -*-
"""胶卷索引 · 本地工具（纯标准库）

  python3 assets/worker.py --serve    启动本地服务（备份 + 日志 + 增删改）
  python3 assets/worker.py --build    重新生成 assets/snapshot.json

目录：
  000_胶卷索引信息/
    ├─ 010_启动Mac.command / 011_启动Win.bat / 030_手动更新索引.command
    └─ assets/
         ├─ index.html        页面
         ├─ films.xlsx        数据源（当数据库用）
         ├─ template.xlsx     模板
         ├─ settings.json     设置       snapshot.json 快照
         ├─ worker.py build_html.py
         ├─ backup/           备份（运行时 10 个 + 定时 30 个）
         └─ log/              日志（按天，最多 30 个）
"""
import os, re, sys, json, csv, shutil, socket, subprocess, zipfile, threading, time, hashlib
import unicodedata
import queue, heapq
import struct, tempfile
import datetime as dt
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs, unquote, quote
import xml.etree.ElementTree as ET

ASSETS = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
IDXDIR = os.path.dirname(ASSETS)
ROOT = os.path.dirname(IDXDIR)
BACKUP = os.path.join(ASSETS, "backup")
LOGDIR = os.path.join(ASSETS, "log")
ICON_DIR = os.path.join(ASSETS, "film_icons")      # 型号图标（列表用，96×77 小图）
ICONHD_DIR = os.path.join(ASSETS, "film_icons_hd") # 同款高清图（悬停放大预览用）
ICONMAP = os.path.join(ASSETS, "icon_map.json") # 手动指定的「型号 → 图标文件」

def cache_root():
    """系统缓存目录（不放机械盘；各系统惯例位置）"""
    home = os.path.expanduser("~")
    if sys.platform == "darwin":
        base = os.path.join(home, "Library", "Caches")
    elif os.name == "nt":
        base = (os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP")
                or os.path.join(home, "AppData", "Local"))
    else:
        base = os.environ.get("XDG_CACHE_HOME") or os.path.join(home, ".cache")
    return os.path.join(base, "FilmIndex")

THUMB_OLD = os.path.join(ASSETS, "thumb")        # 旧版缓存位置（用于迁移）
THUMB = os.path.join(cache_root(), "thumb")       # 预览缩略图缓存（系统缓存目录）
SNAP = os.path.join(ASSETS, "snapshot.json")
SETT = os.path.join(ASSETS, "settings.json")
TPL = os.path.join(ASSETS, "template.xlsx")
UNDO = os.path.join(BACKUP, "_undo.xlsx")
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
MISS = "-"
MAXROWS = 300
VERSION = "0.5.2"          # 公开发布版（文件夹名自动同步 + 一键创建文件夹）

DEFAULT_SET = {"excel": "films.xlsx", "auto": True, "theme": "#4472c4", "throttle": 60,
               "imgOrder": "scan", "dateFmt": "iso", "thumb": True, "thumbQ": 1,
               "thumbMaxMB": 200,
               "ipDir": "000_INDEX PAPER", "ipEdge": 1000, "ipQ": 76,
               "ipRot": "cw", "ipSkip": True, "onePre": True, "folderSync": True, "showIcon": True,
               "colOrder": ["seq", "code", "model", "size", "cam", "lens", "d1", "d2", "dev",
                            "devnote", "da", "re", "na", "ia", "note", "name"],
               "colHide": ["lens", "dev", "devnote", "da", "re", "na", "ia", "note"],
               "stColors": {"拍摄中": "#d94f4f", "待冲洗": "#e0a92e", "冲洗中": "#e0a92e",
                            "待归档": "#9aa0a6"}}
# 拍摄状态：可选值 + 默认值（已归档 = 不高亮）
STATUSES = ["拍摄中", "待冲洗", "冲洗中", "待归档", "已归档"]
STATUS_DEFAULT = "已归档"
# 预览缩略图档位：1=性能优先 … 5=质量优先（最大边 px, JPEG 质量）
THUMB_LEVELS = {1: (360, 55), 2: (540, 65), 3: (720, 72), 4: (1080, 82), 5: (1600, 90)}
COLS = ["序号", "胶卷编号", "胶卷型号", "画幅", "相机", "镜头", "拍摄开始", "拍摄结束",
        "冲扫", "冲扫备注", "数码归档", "翻拍", "底片归档", "索引归档", "备注", "拍摄状态"]
# 选项表三列：key -> (中文名, 上限条数)
OPT_KEYS = (("cam", "相机"), ("size", "画幅"), ("model", "胶卷型号"))
OPT_MAX = {"cam": 200, "size": 50, "model": 400}

for d in (BACKUP, LOGDIR):
    os.makedirs(d, exist_ok=True)
try:
    os.makedirs(THUMB, exist_ok=True)      # 系统缓存目录；不可写时的回退见 ensure_thumb_dir()
except OSError as e:
    sys.stderr.write("系统缓存目录不可用（%s）：%s\n" % (THUMB, e))

# ==================== 日志 ====================
_log_lock = threading.Lock()

def log(level, msg):
    """level: INFO / WARNING / ERROR；按天分文件，最多保留 30 个"""
    level = (level or "INFO").upper()
    now = dt.datetime.now()
    line = "[%s] [%s] %s\n" % (now.strftime("%Y-%m-%d %H:%M:%S"), level, msg)
    try:
        with _log_lock:
            fp = os.path.join(LOGDIR, now.strftime("%Y-%m-%d") + ".txt")
            with open(fp, "a", encoding="utf-8") as fh:
                fh.write(line)
            files = sorted(f for f in os.listdir(LOGDIR) if f.endswith(".txt"))
            while len(files) > 30:
                try:
                    os.remove(os.path.join(LOGDIR, files.pop(0)))
                except OSError:
                    break
    except Exception as e:
        sys.stderr.write("日志写入失败: %s\n" % e)
    try:
        print(line.rstrip())
    except Exception:
        pass

# ==================== 设置 ====================
def load_settings():
    try:
        with open(SETT, encoding="utf-8") as fh:
            o = json.load(fh)
        return {**DEFAULT_SET, **o}
    except Exception:
        log("WARNING", "settings.json 读取失败，使用默认设置")
        return dict(DEFAULT_SET)

def save_settings(o):
    try:
        old = load_settings()
        changed = [k for k in set(old) | set(o) if old.get(k) != o.get(k)]
        with open(SETT, "w", encoding="utf-8") as fh:
            json.dump(o, fh, ensure_ascii=False, indent=2)
        log("INFO", "设置已保存：%s" % json.dumps(o, ensure_ascii=False))
        if changed:
            log("INFO", "设置变更项：%s" % "\u3001".join(sorted(changed)))
        return True
    except Exception as e:
        log("ERROR", "设置保存失败：%s" % e)
        return False

def ensure_excel(path=None):
    """首次运行：数据文件不存在时，自动用 template.xlsx 复制出一份空白数据文件"""
    path = path or os.path.join(ASSETS, "films.xlsx")
    if os.path.exists(path):
        return True
    try:
        if os.path.exists(TPL):
            shutil.copy2(TPL, path)
            log("INFO", "首次运行：数据文件不存在，已用 template.xlsx 生成 %s" % os.path.basename(path))
            return True
        log("ERROR", "首次运行：缺少 template.xlsx，无法生成数据文件")
    except Exception as e:
        log("ERROR", "首次运行生成数据文件失败：%s" % e)
    return False

def excel_path():
    name = (load_settings().get("excel") or "").strip() or "films.xlsx"
    p = os.path.join(ASSETS, name)
    if os.path.exists(p):
        return p
    p2 = os.path.join(ASSETS, "films.xlsx")
    if not os.path.exists(p2):
        ensure_excel(p2)
    if p != p2:
        log("WARNING", "设置里的 Excel「%s」不存在，回退到 films.xlsx" % name)
    return p2

# 目录列举缓存：同一卷短时间内重复列举（重开预览、统计扫描）不再读盘
_LIST_CACHE = {}
LIST_TTL = 60

def safe_under(base, rel):
    p = os.path.abspath(os.path.join(base, rel))
    return p if p == base or p.startswith(base + os.sep) else None

def excel_read_error(p=None):
    """Excel 能不能读：不能读时返回人类可读的说明（可用就直接返回 None）"""
    s = load_settings()
    cfg = (s.get("excel") or "").strip() or "films.xlsx"
    p = p or excel_path()
    name = os.path.basename(p)
    if not os.path.exists(p):
        extra = ("（当前设置里的文件名是「%s」）" % cfg) if cfg != name else ""
        return {"ok": False, "err": "missing", "excel": name, "configured": cfg,
                "msg": "找不到 Excel 文件：assets/%s%s" % (name, extra),
                "hint": "请在「设置 → 数据源」里确认 Excel 文件名，或把数据文件放进 assets/ 目录"}
    if not zipfile.is_zipfile(p):
        return {"ok": False, "err": "notzip", "excel": name, "configured": cfg,
                "msg": "assets/%s 不是有效的 xlsx 文件（可能已损坏或根本不是 Excel）" % name,
                "hint": "请用 Excel 重新另存为 .xlsx 后放回 assets/ 目录"}
    return None

def excel_fallback_warn():
    """设置里的文件名不存在、已静默回退到 films.xlsx 时给出提醒"""
    s = load_settings()
    cfg = (s.get("excel") or "").strip() or "films.xlsx"
    if cfg != "films.xlsx" and not os.path.exists(os.path.join(ASSETS, cfg)) \
            and os.path.exists(os.path.join(ASSETS, "films.xlsx")):
        return ("设置里的 Excel「%s」不存在，已暂时回退到 films.xlsx"
                "（请在「设置 → 数据源」里改正文件名）" % cfg)
    return None

# ==================== 备份 ====================
def _pool(pattern_has_daily):
    fs = []
    for f in os.listdir(BACKUP):
        if not f.lower().endswith(".xlsx") or f.startswith("_"):
            continue
        if ("_daily" in f) == pattern_has_daily:
            fs.append(f)
    return sorted(fs)

def backup(kind):
    """kind: startup / edit / daily / undo-restore"""
    src = excel_path()
    if not os.path.exists(src):
        log("ERROR", "备份失败：找不到 Excel %s" % src)
        return None
    base = os.path.splitext(os.path.basename(src))[0]
    ts = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    dst = os.path.join(BACKUP, "%s_%s_%s.xlsx" % (base, ts, kind))
    try:
        shutil.copy2(src, dst)
    except Exception as e:
        log("ERROR", "备份失败：%s" % e)
        return None
    daily = (kind == "daily")
    keep = 30 if daily else 10
    fs = _pool(daily)
    while len(fs) > keep:
        old = fs.pop(0)
        try:
            os.remove(os.path.join(BACKUP, old))
            log("INFO", "备份轮换：删除旧备份 %s" % old)
        except OSError as e:
            log("WARNING", "删除旧备份失败 %s：%s" % (old, e))
            break
    log("INFO", "已创建备份 %s" % os.path.basename(dst))
    return dst

def last_backup_time(daily=None):
    fs = _pool(bool(daily)) if daily is not None else _pool(False)
    best = None
    for f in fs:
        m = re.search(r"_(\d{8}-\d{6})_", f)
        if not m:
            continue
        try:
            t = dt.datetime.strptime(m.group(1), "%Y%m%d-%H%M%S")
        except ValueError:
            continue
        if best is None or t > best:
            best = t
    return best

def throttle_sec():
    """增删改备份节流秒数，可在设置里改（默认 60 秒；0 表示不节流）"""
    try:
        v = int(float(load_settings().get("throttle", 60)))
    except (TypeError, ValueError):
        v = 60
    return max(0, min(v, 86400))

def backup_on_edit():
    """增删改后备份：距上次增删改备份不足（设置里的节流秒数）则跳过"""
    sec = throttle_sec()
    last = last_backup_time(False)
    if sec and last and (dt.datetime.now() - last).total_seconds() < sec:
        log("INFO", "备份节流：距上次增删改备份不足 %d 秒（设置 throttle=%d），本次跳过备份"
            % (throttle_sec(), sec))
        return None
    return backup("edit")

def backup_on_start():
    """启动/定时备份：距上次定时备份超过 1 天（或从未有过）则做一次定时备份"""
    p = backup("startup")
    last = last_backup_time(True)
    if last is None or (dt.datetime.now() - last).total_seconds() > 86400:
        p2 = backup("daily")
        log("INFO", "已执行定时备份（daily）")
        return p2
    return p
# ==================== xlsx 写入（复用原生成器） ====================

def esc(s):
    if s is None: s = ""
    s = str(s)
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    s = "".join(ch for ch in s if ch in "\t\n" or ord(ch) >= 32)
    return s

def col_letter(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/></numFmts>
<fonts count="4">
<font><sz val="11"/><name val="PingFang SC"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="PingFang SC"/></font>
<font><sz val="11"/><color rgb="FF9C0006"/><name val="PingFang SC"/></font>
<font><sz val="11"/><color rgb="FF808080"/><name val="PingFang SC"/></font>
</fonts>
<fills count="5">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="8">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

def cell_xml(ci, ri, cell):
    """cell: str / (text, style) / ('=FORMULA', style)"""
    if cell is None: return ""
    style = 2
    val = cell
    if isinstance(cell, tuple):
        val, style = cell[0], cell[1]
    if val is None: val = ""
    if isinstance(val, (int, float)):
        return '<c r="%s%d" s="%d"><v>%s</v></c>' % (col_letter(ci), ri, style, val)
    val = str(val)
    if val == "" and not isinstance(cell, tuple):
        return ""
    ref = "%s%d" % (col_letter(ci), ri)
    if val.startswith("="):
        return '<c r="%s" s="%d"><f>%s</f></c>' % (ref, style, esc(val[1:]))
    if val == "":
        return '<c r="%s" s="%d"/>' % (ref, style)
    return '<c r="%s" t="inlineStr" s="%d"><is><t xml:space="preserve">%s</t></is></c>' % (ref, style, esc(val))

def sheet_xml(rows, widths, freeze=None, autofilter=False, validations=None):
    ncol = max(len(r) for r in rows) if rows else 1
    nrow = len(rows)
    p = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
    p.append('<dimension ref="A1:%s%d"/>' % (col_letter(ncol-1), max(nrow,1)))
    p.append('<sheetViews><sheetView workbookViewId="0">')
    if freeze:
        p.append('<pane ySplit="%d" xSplit="0" topLeftCell="A%d" activePane="bottomLeft" state="frozen"/>' % (freeze, freeze+1))
    p.append('</sheetView></sheetViews>')
    p.append('<sheetFormatPr defaultRowHeight="15"/>')
    if widths:
        p.append('<cols>')
        for i, w in enumerate(widths):
            p.append('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i+1, i+1, w))
        p.append('</cols>')
    p.append('<sheetData>')
    for ri, row in enumerate(rows, 1):
        p.append('<row r="%d">' % ri)
        for ci, cell in enumerate(row):
            if ci == 0 and ri == 1:
                st = 1
                val = cell[0] if isinstance(cell, tuple) else cell
                p.append(cell_xml(ci, ri, (val, 1)))
            else:
                p.append(cell_xml(ci, ri, cell))
        p.append('</row>')
    p.append('</sheetData>')
    if autofilter and nrow >= 1:
        p.append('<autoFilter ref="A1:%s%d"/>' % (col_letter(ncol-1), nrow))
    if validations:
        p.append('<dataValidations count="%d">' % len(validations))
        for sqref, f1 in validations:
            p.append('<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="无效输入" '
                     'error="请从下拉列表中选择，或到右侧的选项列中添加新值" sqref="%s"><formula1>%s</formula1></dataValidation>'
                     % (sqref, esc(f1)))
        p.append('</dataValidations>')
    p.append('</worksheet>')
    return "".join(p)

def write_xlsx(path, sheets, defined_names=None):
    n = len(sheets)
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    for i in range(n):
        ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i+1))
    ct.append("</Types>")
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')
    wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
          '<workbookPr/><bookViews><workbookView/></bookViews><sheets>']
    for i, s in enumerate(sheets):
        wb.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (esc(s[0]), i+1, i+1))
    wb.append('</sheets>')
    if defined_names:
        wb.append('<definedNames>')
        for nm, ref in defined_names:
            wb.append('<definedName name="%s">%s</definedName>' % (esc(nm), esc(ref)))
        wb.append('</definedNames>')
    wb.append('<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    wbrels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i in range(n):
        wbrels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i+1, i+1))
    wbrels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (n+1))
    wbrels.append("</Relationships>")

    z = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    z.writestr("[Content_Types].xml", "".join(ct))
    z.writestr("_rels/.rels", rels)
    z.writestr("xl/workbook.xml", "".join(wb))
    z.writestr("xl/_rels/workbook.xml.rels", "".join(wbrels))
    z.writestr("xl/styles.xml", STYLES)
    for i, s in enumerate(sheets):
        name, rows, widths, freeze, af = (list(s) + [None, None, None])[:5]
        val = s[5] if len(s) > 5 else None
        z.writestr("xl/worksheets/sheet%d.xml" % (i+1), sheet_xml(rows, widths, freeze, af, val))
    z.close()

# ==================== 读取 Excel ====================
HDR = ["序号", "胶卷编号", "型号", "画幅", "相机", "镜头", "拍摄开始", "拍摄结束",
       "冲扫", "冲扫备注", "数码归档", "翻拍", "底片归档", "索引归档", "备注",
       "拍摄状态", "文件夹名称（自动生成）"]
W = [6, 10, 24, 6, 24, 30, 12, 12, 12, 12, 12, 12, 12, 12, 28, 14, 52]
OPT_HDR = ["相机选项（在下方继续添加新型号即可）", "画幅选项", "胶卷型号选项（在下方继续添加新型号即可）"]
DATE_COLS = (6, 7, 8, 10, 11, 12, 13)
NAMES = [("相机列表", "选项!$A$2:$A$200"), ("画幅列表", "选项!$B$2:$B$50"),
         ("胶卷型号列表", "选项!$C$2:$C$400")]
VAL = [("C2:C400", "OFFSET(胶卷型号列表,0,0,MAX(1,COUNTA(胶卷型号列表)),1)"),
       ("D2:D400", "OFFSET(画幅列表,0,0,MAX(1,COUNTA(画幅列表)),1)"),
       ("E2:E400", "OFFSET(相机列表,0,0,MAX(1,COUNTA(相机列表)),1)")]

def colnum(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1

def as_date(v):
    v = (v or "").strip()
    if not v:
        return ""
    if re.fullmatch(r"\d+(\.\d+)?", v):
        n = int(float(v))
        if 20000 < n < 80000:
            return (dt.date(1899, 12, 30) + dt.timedelta(days=n)).strftime("%Y-%m-%d")
        return ""
    m = re.match(r"(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", v)
    if m and 1990 <= int(m.group(1)) <= 2100:
        return "%04d-%02d-%02d" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.fullmatch(r"(\d{2})(\d{2})(\d{2})", v)
    if m and 1 <= int(m.group(2)) <= 12:
        return "20%s-%s-%s" % m.groups()
    return v

def iso_to_serial(s):
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", (s or "").strip())
    if not m:
        return None
    try:
        d = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None
    return (d - dt.date(1899, 12, 30)).days

def read_grid(path=None):
    """返回 (rows, options)；rows 为 16 列字符串列表"""
    path = path or excel_path()
    z = zipfile.ZipFile(path)
    sst = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(NS + "si"):
            sst.append("".join(t.text or "" for t in si.iter(NS + "t")))

    def sh(name):
        d = {}
        root = ET.fromstring(z.read(name))
        for row in root.iter(NS + "row"):
            try:
                ri = int(row.get("r"))
            except (TypeError, ValueError):
                continue
            cells = {}
            for c in row.iter(NS + "c"):
                m = re.match(r"([A-Z]+)", c.get("r") or "")
                if not m:
                    continue
                ci = colnum(m.group(1))
                t = c.get("t")
                if t == "inlineStr":
                    isx = c.find(NS + "is")
                    v = "".join(x.text or "" for x in isx.iter(NS + "t")) if isx is not None else ""
                else:
                    ve = c.find(NS + "v")
                    v = (ve.text or "") if ve is not None else ""
                    if t == "s":
                        try:
                            v = sst[int(v)]
                        except (ValueError, IndexError):
                            v = ""
                cells[ci] = v
            d[ri] = cells
        return d

    s1 = sh("xl/worksheets/sheet1.xml")
    s2 = sh("xl/worksheets/sheet2.xml")
    rows = []
    for ri in range(2, MAXROWS + 1):
        cells = s1.get(ri, {})
        vals = []
        for ci in range(17):
            v = str(cells.get(ci, "")).strip()
            if ci == 0 and v:
                try:
                    if v.isdigit() and 3 <= len(v) <= 4:
                        pass                       # 三位 / 四位序号原样保留
                    else:
                        v = "%03d" % int(float(v))
                except ValueError:
                    pass
            elif ci in DATE_COLS:
                v = as_date(v)
            if v == MISS:
                v = ""
            vals.append(v)
        rows.append(vals)
    while rows and not any(rows[-1][:15]):
        rows.pop()
    opts = {}
    for k, ci in (("cam", 0), ("size", 1), ("model", 2)):
        lst = []
        for ri in range(2, MAXROWS + 1):
            v = str(s2.get(ri, {}).get(ci, "")).strip()
            if v:
                lst.append(v)
        opts[k] = lst
    return rows, opts

# ==================== 写出 Excel ====================
def p_formula(ri):
    return ('=IF(OR($A{r}="",$B{r}="",$C{r}="",$G{r}=""),"",'
            '$A{r}&"_index "&$B{r}&"_"&$C{r}&"_"&'
            'TEXT($G{r},"yy")&TEXT($G{r},"mm")&TEXT($G{r},"dd")&'
            'IF(OR($H{r}="",$H{r}=$G{r}),"","-"&TEXT($H{r},"yy")&TEXT($H{r},"mm")&TEXT($H{r},"dd")))').format(r=ri)

def write_workbook(rows, opts, path=None):
    path = path or excel_path()
    sheet1 = [HDR]
    for i in range(MAXROWS - 1):
        ri = i + 2
        vals = rows[i] if i < len(rows) else [""] * 17
        has_data = any(str(vals[k] if k < len(vals) else "").strip() for k in range(15))
        line = []
        for ci in range(17):
            if ci == 16:
                line.append((p_formula(ri), 2))
                continue
            v = str(vals[ci] if ci < len(vals) else "").strip()
            if ci == 14:
                line.append((v, 3))
            elif ci == 15:
                line.append((v if v in STATUSES else (STATUS_DEFAULT if has_data else ""), 2))
            elif ci in DATE_COLS:
                s = iso_to_serial(v)
                line.append((s, 6) if s is not None else (v, 6))
            else:
                line.append((v, 2))
        sheet1.append(line)
    rows_opt = [OPT_HDR]
    cams, sizes, models = opts.get("cam", []), opts.get("size", []), opts.get("model", [])
    for i in range(max(len(cams), len(sizes), len(models), 1)):
        rows_opt.append([cams[i] if i < len(cams) else "",
                         sizes[i] if i < len(sizes) else "",
                         models[i] if i < len(models) else ""])
    write_xlsx(path, [("胶卷信息总表", sheet1, W, 1, True, VAL),
                      ("选项", rows_opt, [42, 16, 46], 1, False)], defined_names=NAMES)

# ==================== 快照 ====================
def _split_model(s):
    """把型号拆成 (基础名, 格式)。
    括号里如果是 110/120/135 就是【格式标记】要保留；
    其它括号（EI 1600 / PUSH 400 / REVERSAL 160 等）是【校准·冲洗标记】，一律忽略。"""
    s = unicodedata.normalize("NFC", str(s or ""))
    m = re.search(r"[\(（]\s*(110|120|135)\s*[\)）]", s)
    fmt = m.group(1) if m else ""
    s = re.sub(r"[\(（].*?[\)）]", " ", s)
    s = s.lower()
    s = re.sub(r"[^0-9a-z\u4e00-\u9fff]+", " ", s)
    return " ".join(s.split()), fmt


def load_icon_map():
    """读手动映射表；文件不存在或损坏都当作空表"""
    try:
        with open(ICONMAP, encoding="utf-8") as fh:
            d = json.load(fh)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def save_icon_map(d):
    try:
        with open(ICONMAP, "w", encoding="utf-8") as fh:
            json.dump(d, fh, ensure_ascii=False, indent=1, sort_keys=True)
        return True
    except Exception as e:
        log("WARNING", "写 icon_map.json 失败：%s" % e)
        return False


def pick_icon(mp, idx, model):
    """优先用手动映射（且文件确实存在）→ 否则自动按名称匹配"""
    f = (mp or {}).get(str(model or "").strip())
    if f and os.path.isfile(os.path.join(ICON_DIR, f)):
        return f
    return _icon_pick(idx, model)


def icon_all_files():
    try:
        return sorted(f for f in os.listdir(ICON_DIR) if f.lower().endswith(".png"))
    except OSError:
        return []


def icon_report():
    """给设置面板用：每个型号的当前图标 + 来源"""
    mp, idx = load_icon_map(), _icon_files()
    rows, opts = read_grid()
    names = []
    for n in (opts.get("model") or []):
        n = str(n).strip()
        if n and n not in names:
            names.append(n)
    for r in rows:
        if len(r) > 2:
            n = r[2].strip()
            if n and n not in names:
                names.append(n)
    out = []
    for n in names:
        man = mp.get(n)
        auto = _icon_pick(idx, n) or ""
        cur = man if (man and os.path.isfile(os.path.join(ICON_DIR, man))) else auto
        out.append({"model": n, "icon": cur or "", "manual": bool(man and man == cur), "auto": auto})
    return {"ok": True, "list": out, "files": icon_all_files(),
            "missing": len([x for x in out if not x["icon"]])}


def _icon_files():
    """assets/film_icons/*.png → {基础名: {格式: 文件名}}（格式 "" 表示不限画幅）"""
    idx = {}
    try:
        for f in sorted(os.listdir(ICON_DIR)):
            if not f.lower().endswith(".png"):
                continue
            base, fmt = _split_model(os.path.splitext(f)[0])
            if base:
                idx.setdefault(base, {})[fmt] = f
    except OSError:
        pass
    return idx


# 型号别名：查不到自己的图时，借用同款的图
ICON_ALIAS = {
    "kodak 5219 ahu": "kodak 500t 5219",   # AHU = 新版（无碳层）5219，先用普通 5219 的图
    "lucky c200 t": "lucky c200",          # T = 测试卷，与 C200 同款
}


def _icon_pick(idx, model):
    """按型号挑图标：优先同画幅 → 再退到不限画幅 → 再退到任意一张"""
    base, fmt = _split_model(model)
    cands = idx.get(base)
    if not cands:
        alt = ICON_ALIAS.get(base)
        if alt:
            cands = idx.get(alt)
    if not cands:
        return None
    if fmt and fmt in cands:
        return cands[fmt]
    if "" in cands:
        return cands[""]
    return sorted(cands.values())[0]


def _model_icons(lst):
    """给「选项库里的型号」逐个配上图标文件名（供编辑表单下拉框用）"""
    idx, mp = _icon_files(), load_icon_map()
    out = {}
    for m in (lst or []):
        f = pick_icon(mp, idx, m)
        if f:
            out[m] = f
    return out


def build_snapshot():
    rows, _ = read_grid()
    data = []
    issues = []
    icons = _icon_files()
    _mp = load_icon_map()
    for r in rows:
        seq = r[0].strip()
        if not re.fullmatch(r"\d{3,4}", seq):
            continue
        if not r[6]:
            continue
        if r[7] and r[7] < r[6]:            # 日期合理性检查：结束不能早于开始
            issues.append({"seq": seq, "d1": r[6], "d2": r[7]})
        name = "%s_index %s_%s_%s" % (seq, r[1], r[2], r[6][2:4] + r[6][5:7] + r[6][8:10])
        if r[7] and r[7] != r[6]:
            name += "-" + r[7][2:4] + r[7][5:7] + r[7][8:10]
        st = (r[15] if len(r) > 15 else "") or ""
        data.append(dict(seq=seq, code=r[1], model=r[2], size=r[3], cam=r[4] or MISS,
                         lens=r[5], d1=r[6], d2=r[7], dev=r[8], devnote=r[9],
                         da=r[10], re=r[11], na=r[12], ia=r[13],
                         name=name, note=r[14],
                         st=(st if st in STATUSES else STATUS_DEFAULT),
                         icon=pick_icon(_mp, icons, r[2])))
    if issues:
        log("WARNING", "数据检查：%d 条记录的「拍摄结束」早于「拍摄开始」（序号 %s）"
            % (len(issues), "、".join(i["seq"] for i in issues[:20])))
    return {"generated": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "excel": os.path.basename(excel_path()), "count": len(data),
            "issues": issues, "data": data}

def write_snapshot():
    snap = build_snapshot()
    with open(SNAP, "w", encoding="utf-8") as fh:
        json.dump(snap, fh, ensure_ascii=False)
    return snap

# ==================== 增 / 删 / 改 / 撤回 ====================
LAST = {"desc": None, "ts": None}
_lock = threading.Lock()

def _clean(cells):
    out = []
    for i in range(16):
        v = cells[i] if i < len(cells) else ""
        v = str(v if v is not None else "").strip()
        if v == MISS:          # 表格里用 - 表示缺省，写回时按空处理
            v = ""
        out.append(v)
    return out

def validate(cells, rows, key=None):
    c = _clean(cells)
    if c[15] not in STATUSES:          # 拍摄状态：非法/空 → 默认已归档
        c[15] = STATUS_DEFAULT
    if not c[2]:
        return None, "胶卷型号为必填项"
    if not c[6]:
        return None, "拍摄开始日期为必填项"
    for i, v in enumerate(c):
        if len(v) > 200:
            return None, "「%s」超过 200 字上限" % COLS[i]
    for ci, label in ((6, "拍摄开始"), (7, "拍摄结束"), (8, "冲扫"),
                      (10, "数码归档"), (11, "翻拍"), (12, "底片归档"), (13, "索引归档")):
        if c[ci] and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", c[ci]):
            return None, "「%s」日期格式应为 YYYY-MM-DD" % label
    if c[6] and c[7] and c[7] < c[6]:
        return None, "「拍摄结束」(%s) 不能早于「拍摄开始」(%s)" % (c[7], c[6])
    if key is None:
        if not re.fullmatch(r"\d{3,4}", c[0]):
            return None, "序号必须是三位或四位数字（如 117 / 1175）"
        if any(r[0].strip() == c[0] for r in rows):
            return None, "序号 %s 已存在，请换一个" % c[0]
    else:
        c[0] = key
    return c, None

def next_seq(rows):
    mx = 0
    for r in rows:
        if re.fullmatch(r"\d{3,4}", r[0].strip()):
            mx = max(mx, int(r[0]))
    n = mx + 1
    return "%03d" % n if n <= 999 else "%04d" % n

def excel_lock_file():
    """Excel / LibreOffice 打开数据文件时会生成锁文件；存在则返回其路径"""
    p = excel_path()
    d, b = os.path.dirname(p), os.path.basename(p)
    for cand in (os.path.join(d, "~$" + b), os.path.join(d, ".~lock." + b + "#")):
        if os.path.exists(cand):
            return cand
    return None

def excel_lock_warning():
    """写入后提示：Excel 还开着的话，我们刚写的内容可能被它保存时覆盖"""
    lock = excel_lock_file()
    if lock:
        log("WARNING", "检测到 Excel 正打开该文件（%s），写入可能被其保存覆盖" % os.path.basename(lock))
        return "⚠ Excel 正打开着该文件，你的修改可能在 Excel 保存时被覆盖，建议先关闭 Excel 再改"
    return None

def guard_lock(force, opname):
    """写入前的前置检查：Excel 开着则拒绝（除非用户显式确认强制继续）"""
    lock = excel_lock_file()
    if not lock:
        return None
    xlsx = os.path.basename(excel_path())
    if force:
        log("WARNING", "%s：Excel 正打开 %s（检测到锁文件 %s），用户选择强制继续写入"
            % (opname, xlsx, os.path.basename(lock)))
        return None
    log("WARNING", "%s 被拒绝：Excel 正打开 %s（检测到锁文件 %s），请先关闭 Excel"
        % (opname, xlsx, os.path.basename(lock)))
    return {"ok": False, "locked": True, "lockfile": os.path.basename(lock), "excel": xlsx,
            "msg": "检测到 Excel 正打开 %s，为避免修改被覆盖，请先关闭 Excel 再操作" % xlsx}

def _prepare_undo():
    src = excel_path()
    try:
        shutil.copy2(src, UNDO)
    except Exception as e:
        log("WARNING", "撤回点写入失败：%s" % e)

def _after_write(tag):
    try:
        backup_on_edit()
    except Exception as e:
        log("ERROR", "操作备份失败：%s" % e)
    try:
        write_snapshot()
    except Exception as e:
        log("ERROR", "快照刷新失败：%s" % e)

def op_add(cells, force=False):
    with _lock:
        bad = guard_lock(force, "新增")
        if bad:
            return bad
        rows, opts = read_grid()
        c, err = validate(cells, rows)
        if err:
            log("WARNING", "新增被拒绝：%s" % err)
            return {"ok": False, "msg": err}
        rows.append(c)
        rows.sort(key=lambda r: int(r[0]) if re.fullmatch(r"\d{3,4}", r[0]) else 9999)
        _prepare_undo()
        try:
            write_workbook(rows, opts)
        except Exception as e:
            log("ERROR", "新增写入失败：%s" % e)
            return {"ok": False, "msg": "写入 Excel 失败：%s" % e}
        chk, _ = read_grid()
        ok = any(r[0].strip() == c[0] and r[2].strip() == c[2] for r in chk)
        name = "%s_index %s_%s_%s" % (c[0], c[1], c[2], c[6].replace("-", "")[2:] if c[6] else "")
        desc = "新增条目 %s（%s）" % (c[0], name)
        LAST["desc"], LAST["ts"] = desc, dt.datetime.now()
        log("INFO", "%s；写入校验：%s" % (desc, "成功" if ok else "失败"))
        if not ok:
            log("ERROR", "新增后校验未通过（序号 %s）" % c[0])
        _after_write("add")
        fr = None
        try:
            fr = folder_info(c[0], roll_name(c))
        except Exception as e:
            log("WARNING", "文件夹检查出错：%s: %s" % (type(e).__name__, e))
        return {"ok": ok, "verified": ok, "msg": "已新增" if ok else "已写入但校验未通过",
                "desc": desc, "seq": c[0], "next": next_seq(chk), "warn": excel_lock_warning(),
                "folder": fr}

def op_update(key, cells, force=False):
    with _lock:
        bad = guard_lock(force, "修改")
        if bad:
            return bad
        rows, opts = read_grid()
        idx = next((i for i, r in enumerate(rows) if r[0].strip() == key), None)
        if idx is None:
            log("WARNING", "修改失败：找不到序号 %s" % key)
            return {"ok": False, "msg": "找不到序号 %s" % key}
        c, err = validate(cells, rows, key=key)
        if err:
            log("WARNING", "修改被拒绝：%s" % err)
            return {"ok": False, "msg": err}
        old = rows[idx][:]
        changes = [COLS[i] for i in range(16) if old[i].strip() != c[i].strip()]
        if not changes:
            return {"ok": False, "msg": "没有检测到任何修改"}
        rows[idx] = c
        _prepare_undo()
        try:
            write_workbook(rows, opts)
        except Exception as e:
            log("ERROR", "修改写入失败：%s" % e)
            return {"ok": False, "msg": "写入 Excel 失败：%s" % e}
        chk, _ = read_grid()
        now = next((r for r in chk if r[0].strip() == key), None)
        ok = bool(now) and all(now[i].strip() == c[i].strip() for i in range(16))
        desc = "修改条目 %s（%s）的：%s" % (key, c[2], "、".join(changes))
        LAST["desc"], LAST["ts"] = desc, dt.datetime.now()
        log("INFO", "%s；写入校验：%s" % (desc, "成功" if ok else "失败"))
        if not ok:
            log("ERROR", "修改后校验未通过（序号 %s）" % key)
        _after_write("update")
        fr = None
        try:
            if load_settings().get("folderSync", True):
                fr = folder_sync(key, c, old)      # 直接同步改名，不再让用户二次确认
        except Exception as e:
            log("WARNING", "文件夹同步出错：%s: %s" % (type(e).__name__, e))
        return {"ok": ok, "verified": ok, "msg": "已修改" if ok else "已写入但校验未通过",
                "desc": desc, "seq": key, "changes": changes, "warn": excel_lock_warning(),
                "folder": fr}

def op_delete(key, force=False, del_folder=False):
    with _lock:
        bad = guard_lock(force, "删除")
        if bad:
            return bad
        rows, opts = read_grid()
        idx = next((i for i, r in enumerate(rows) if r[0].strip() == key), None)
        if idx is None:
            log("WARNING", "删除失败：找不到序号 %s" % key)
            return {"ok": False, "msg": "找不到序号 %s" % key}
        gone = rows.pop(idx)
        _prepare_undo()
        try:
            write_workbook(rows, opts)
        except Exception as e:
            log("ERROR", "删除写入失败：%s" % e)
            return {"ok": False, "msg": "写入 Excel 失败：%s" % e}
        chk, _ = read_grid()
        ok = not any(r[0].strip() == key for r in chk)
        name = "%s_index %s_%s" % (gone[0], gone[1], gone[2])
        desc = "删除条目 %s（%s）" % (key, name)
        LAST["desc"], LAST["ts"] = desc, dt.datetime.now()
        log("WARNING", "%s；写入校验：%s" % (desc, "成功" if ok else "失败"))
        if not ok:
            log("ERROR", "删除后校验未通过（序号 %s）" % key)
        _after_write("delete")
        fr = None
        if ok and del_folder:
            try:
                fr = folder_delete(key, roll_name(gone))
            except Exception as e:
                log("WARNING", "删除文件夹出错：%s: %s" % (type(e).__name__, e))
        return {"ok": ok, "verified": ok, "msg": "已删除" if ok else "已写入但校验未通过",
                "desc": desc, "seq": key, "warn": excel_lock_warning(), "folder": fr}

def op_undo(force=False):
    with _lock:
        if not os.path.exists(UNDO):
            log("WARNING", "撤回失败：没有可用的撤回点")
            return {"ok": False, "msg": "没有可撤回的操作"}
        if LAST["ts"] and (dt.datetime.now() - LAST["ts"]).total_seconds() > 180:
            log("WARNING", "撤回失败：超出撤回时限")
            return {"ok": False, "msg": "已超出撤回时限"}
        bad = guard_lock(force, "撤回")
        if bad:
            return bad
        try:
            shutil.copy2(UNDO, excel_path())
        except Exception as e:
            log("ERROR", "撤回写入失败：%s" % e)
            return {"ok": False, "msg": "撤回失败：%s" % e}
        desc = LAST["desc"] or "上一次操作"
        log("WARNING", "已撤回：%s" % desc)
        try:
            write_snapshot()
        except Exception as e:
            log("ERROR", "撤回后快照刷新失败：%s" % e)
        LAST["desc"], LAST["ts"] = None, None
        return {"ok": True, "msg": "已撤回：%s" % desc, "desc": desc}

# ==================== 选项表（相机 / 画幅 / 胶卷型号） ====================
def _norm_opts(src):
    """去空行 / 去重 / 截断超长，按上限裁列"""
    out = {}
    for k, _label in OPT_KEYS:
        lst, seen = [], set()
        for v in (src.get(k) or []):
            v = str(v if v is not None else "").strip()
            if not v or v in seen:
                continue
            if len(v) > 100:
                v = v[:100]
            seen.add(v)
            lst.append(v)
        if len(lst) > OPT_MAX[k]:
            log("WARNING", "选项「%s」超出上限 %d，已截断" % (k, OPT_MAX[k]))
            lst = lst[:OPT_MAX[k]]
        out[k] = lst
    return out

def op_save_options(src, force=False):
    """把设置里编辑的选项写回 Excel 的「选项」工作表"""
    with _lock:
        bad = guard_lock(force, "保存选项")
        if bad:
            return bad
        rows, old = read_grid()
        new = _norm_opts(src or {})
        oldn = {k: list(old.get(k) or []) for k, _ in OPT_KEYS}
        if new == oldn:
            return {"ok": False, "msg": "选项没有任何变化"}
        _prepare_undo()
        try:
            write_workbook(rows, new)
        except Exception as e:
            log("ERROR", "选项写入失败：%s" % e)
            return {"ok": False, "msg": "写入 Excel 失败：%s" % e}
        chk = read_grid()[1]
        ok = all(list(chk.get(k) or []) == new[k] for k, _ in OPT_KEYS)
        summary = "、".join("%s %d→%d 项" % (lb, len(oldn[k]), len(new[k])) for k, lb in OPT_KEYS)
        desc = "修改选项表（%s）" % summary
        LAST["desc"], LAST["ts"] = desc, dt.datetime.now()
        log("INFO", "%s；写入校验：%s" % (desc, "成功" if ok else "失败"))
        if not ok:
            log("ERROR", "选项保存后校验未通过")
        _after_write("options")
        return {"ok": ok, "verified": ok, "msg": "选项已保存" if ok else "已写入但校验未通过",
                "desc": desc, "options": new,
                "counts": {k: len(new[k]) for k, _ in OPT_KEYS},
                "warn": excel_lock_warning()}

def ensure_thumb_dir():
    """确认缩略图缓存目录可写；系统缓存目录不可用时自动回退到 assets/thumb"""
    global THUMB
    for d in (THUMB, os.path.join(ASSETS, "thumb")):
        try:
            os.makedirs(d, exist_ok=True)
            probe = os.path.join(d, ".write_test")
            with open(probe, "w") as fh:
                fh.write("1")
            os.remove(probe)
            if d != THUMB:
                log("WARNING", "系统缓存目录不可写，缩略图缓存回退到：%s" % d)
            else:
                log("INFO", "缩略图缓存目录可用：%s" % d)
            THUMB = d
            return d
        except Exception as e:
            log("WARNING", "缓存目录不可用（%s）：%s" % (d, e))
    return THUMB

def migrate_thumb_dir():
    """把旧版放在 assets/thumb 的缓存搬到系统缓存目录（只做一次，失败不影响使用）"""
    try:
        if not os.path.isdir(THUMB_OLD) or os.path.abspath(THUMB_OLD) == os.path.abspath(THUMB):
            return
        names = [f for f in os.listdir(THUMB_OLD) if f.lower().endswith(".jpg")]
        if not names:
            try:
                os.rmdir(THUMB_OLD)
            except OSError:
                pass
            return
        moved = 0
        for f in names:
            src, dst = os.path.join(THUMB_OLD, f), os.path.join(THUMB, f)
            if os.path.exists(dst):
                continue
            try:
                shutil.move(src, dst)
                moved += 1
            except Exception as e:
                log("WARNING", "迁移缩略图失败 %s：%s" % (f, e))
                break
        log("INFO", "预览缓存已迁移到系统缓存目录：%d 个文件 → %s" % (moved, THUMB))
        try:
            os.rmdir(THUMB_OLD)
        except OSError:
            pass
    except Exception as e:
        log("WARNING", "预览缓存迁移失败（不影响使用）：%s" % e)

# ==================== 预览缩略图（降采样缓存） ====================
# 点「预览」时用降采样小图（磁盘缓存）；点开大图才读原图。
# 缩略图生成并发控制（0.4.1）：
#   · 同一目录用「目录锁」串行 —— Windows 下一次把整卷做完，只起一个 PowerShell 进程；
#   · 不同目录之间用信号量限制并发，避免同时开太多进程。
_thumb_lock = threading.Lock()          # 兼容保留
_thumb_sem = threading.Semaphore(4)     # 非批量后端的并发上限
_dir_locks = {}
_dir_locks_guard = threading.Lock()
_rotate_guard = threading.Lock()
_last_rotate = [0.0]
THUMB_ROTATE_MIN_SEC = 30               # 缩略图缓存轮换的最小间隔（避免每张图都全目录扫描）
THUMB_BATCH_MAX = 300                   # 单次批量最多处理多少张
_backend_logged = False

def _dir_lock(d):
    """同一目录一把锁，保证整卷只被批处理一次"""
    with _dir_locks_guard:
        lk = _dir_locks.get(d)
        if lk is None:
            lk = threading.Lock()
            _dir_locks[d] = lk
        return lk

# 正在读取原图的请求数：>0 时缩略图生成让路（看大图优先）
_img_lock = threading.Lock()
IMG_BUSY = [0]

def img_enter():
    with _img_lock:
        IMG_BUSY[0] += 1

def img_leave():
    with _img_lock:
        IMG_BUSY[0] = max(0, IMG_BUSY[0] - 1)

def wait_full_img(timeout=6.0):
    """等原图读完再生成缩略图；返回实际等待秒数"""
    t0 = time.time()
    while IMG_BUSY[0] > 0 and (time.time() - t0) < timeout:
        time.sleep(0.05)
    return time.time() - t0

def thumb_params():
    s = load_settings()
    on = bool(s.get("thumb", True))
    try:
        lvl = int(s.get("thumbQ", 1))
    except (TypeError, ValueError):
        lvl = 3
    lvl = max(1, min(5, lvl))
    edge, quality = THUMB_LEVELS[lvl]
    return on, lvl, edge, quality

def thumb_limits():
    """返回 (字节上限, 文件数安全上限)；大小可在设置里改，默认 200 MB"""
    s = load_settings()
    try:
        mb = int(float(s.get("thumbMaxMB", 200)))
    except (TypeError, ValueError):
        mb = 200
    mb = max(20, min(5000, mb))
    return mb * 1048576, max(2000, mb * 20)

def thumb_backend():
    """优先用系统自带工具（Mac: sips；Win: PowerShell/GDI+；Linux: convert）"""
    if sys.platform == "darwin" and shutil.which("sips"):
        return "sips"
    if os.name == "nt" and shutil.which("powershell"):
        return "powershell"
    if shutil.which("convert"):
        return "imagemagick"
    return None

def thumb_path(src, edge, quality):
    try:
        st = os.stat(src)
    except OSError:
        return None
    key = "%s|%d|%d|%d|%d" % (os.path.abspath(src), int(st.st_mtime), st.st_size, edge, quality)
    h = hashlib.md5(key.encode("utf-8")).hexdigest()[:20]
    return os.path.join(THUMB, "%s_%d_%d.jpg" % (h, edge, quality))

def _run(cmd, timeout, env=None):
    return subprocess.check_call(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                 timeout=timeout, env=env)

def make_thumb(src, dst, edge, quality, rotate=0):
    """先生成到临时文件再原子替换，避免半截 JPEG 被当成缓存
    rotate: 0=不动，90=顺时针 90°，-90=逆时针 90°（导出 Index Paper 竖图用）"""
    tmp = "%s.%d.part" % (dst, threading.get_ident())
    try:
        if _make_thumb_to(src, tmp, edge, quality, rotate):
            os.replace(tmp, dst)
            return True
        return False
    finally:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass

def _make_thumb_to(src, dst, edge, quality, rotate=0):
    """返回 True 表示 dst 已生成"""
    be = thumb_backend()
    if be == "sips":
        out = subprocess.check_output(["sips", "-g", "pixelWidth", "-g", "pixelHeight", src],
                                      stderr=subprocess.DEVNULL, timeout=30).decode("utf-8", "ignore")
        mw = re.search(r"pixelWidth:\s*(\d+)", out)
        mh = re.search(r"pixelHeight:\s*(\d+)", out)
        args = ["sips", "-s", "format", "jpeg", "-s", "formatOptions", str(quality)]
        if rotate:
            args += ["-r", str(int(rotate))]
        if mw and mh and max(int(mw.group(1)), int(mh.group(1))) > edge:
            args += ["--resampleHeightWidthMax", str(edge)]
        args += [src, "--out", dst]
        _run(args, 120)
        return os.path.exists(dst) and os.path.getsize(dst) > 0
    if be == "powershell":
        script = (
            "$ErrorActionPreference='Stop';Add-Type -AssemblyName System.Drawing;"
            "$img=[System.Drawing.Image]::FromFile($env:SRC);try{"
            "$w=$img.Width;$h=$img.Height;$m=[Math]::Max($w,$h);$e=%d;"
            "if($m -gt $e){$s=$e/$m;$w=[int]($w*$s);$h=[int]($h*$s)}"
            "$bmp=New-Object System.Drawing.Bitmap $w,$h;"
            "$g=[System.Drawing.Graphics]::FromImage($bmp);"
            "$g.InterpolationMode=[System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic;"
            "$g.DrawImage($img,0,0,$w,$h);"
            "$ci=[System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders()|Where-Object{$_.MimeType -eq 'image/jpeg'};"
            "$p=New-Object System.Drawing.Imaging.EncoderParameters 1;"
            "$p.Param[0]=New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality,[long]%d);"
            "if([int]$env:ROT -ne 0){if([int]$env:ROT -gt 0){"
            "$bmp.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipNone)}"
            "else{$bmp.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipNone)}}"
            "$bmp.Save($env:DST,$ci,$p);$bmp.Dispose();$g.Dispose()"
            "}finally{$img.Dispose()}"
        ) % (edge, quality)
        _run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script], 180,
             env=dict(os.environ, SRC=src, DST=dst, ROT=str(int(rotate))))
        return os.path.exists(dst) and os.path.getsize(dst) > 0
    if be == "imagemagick":
        args = ["convert", src + "[0]", "-auto-orient"]
        if rotate:
            args += ["-rotate", str(int(rotate))]
        args += ["-resize", "%dx%d>" % (edge, edge), "-quality", str(quality), dst]
        _run(args, 120)
        return os.path.exists(dst) and os.path.getsize(dst) > 0
    return False

def _log_batch_stats(statfile):
    """把批量的逐张耗时汇总写进日志，方便定位 Windows 上到底慢在哪一步。"""
    try:
        if not statfile or not os.path.exists(statfile) or os.path.getsize(statfile) == 0:
            return
        with open(statfile, encoding="utf-8-sig") as fh:      # PowerShell 会带 BOM
            data = json.load(fh)
        if isinstance(data, dict):
            data = [data]
        if not data:
            return
        ms = [int(x.get("ms") or 0) for x in data]
        top = sorted(data, key=lambda x: -int(x.get("ms") or 0))[:3]

        def desc(x):
            return "%s(%sx%s / %.1fMB / %dms)" % (
                os.path.basename(x.get("dst") or ""), x.get("w"), x.get("h"),
                int(x.get("bytes") or 0) / 1048576.0, int(x.get("ms") or 0))

        log("INFO", "批量缩略图耗时明细：%d 张 / 总 %.1fs / 平均 %dms / 最慢 3 张：%s"
            % (len(ms), sum(ms) / 1000.0, sum(ms) // len(ms), "；".join(desc(x) for x in top)))
    except Exception as e:
        log("WARNING", "缩略图耗时统计读取失败：%s" % e)


_waiters = {}
_waiters_guard = threading.Lock()
STREAM_POLL = 0.08


def _notify_thumb(dst):
    with _waiters_guard:
        for ev in _waiters.pop(dst, []) or []:
            ev.set()


def _wait_thumb(dst, timeout):
    """等某一张缩略图出现；出现即返回 True。"""
    if os.path.exists(dst):
        return True
    ev = threading.Event()
    with _waiters_guard:
        _waiters.setdefault(dst, []).append(ev)
    try:
        return ev.wait(timeout)
    finally:
        with _waiters_guard:
            lst = _waiters.get(dst)
            if lst and ev in lst:
                lst.remove(ev)


def _claim_tmp(pending):
    """把已经落盘的 tmp 立刻认领成正式缓存文件，并唤醒等待它的请求。"""
    n = 0
    for tmp in list(pending):
        dst = pending[tmp]
        if os.path.exists(tmp) and os.path.getsize(tmp) > 0:
            try:
                os.replace(tmp, dst)
                n += 1
                _notify_thumb(dst)
            except OSError:
                pass
            pending.pop(tmp, None)
    return n


def _log_ps_diag(dbgfile, failfile):
    """把 PowerShell 批量脚本的自诊断信息与失败样本写进日志。"""
    try:
        if dbgfile and os.path.exists(dbgfile) and os.path.getsize(dbgfile) > 0:
            with open(dbgfile, encoding="utf-8-sig") as fh:
                t = fh.read().strip().replace("\n", " | ")[:400]
            if t:
                log("INFO", "批量缩略图诊断：" + t)
    except Exception as e:
        log("WARNING", "批量诊断读取失败：%s" % e)
    try:
        if failfile and os.path.exists(failfile) and os.path.getsize(failfile) > 0:
            with open(failfile, encoding="utf-8-sig") as fh:
                rows = [x.strip() for x in fh.read().splitlines() if x.strip()]
            if rows:
                log("WARNING", "批量失败样本（共 %d 条）：%s"
                    % (len(rows), " ／ ".join(x[:220] for x in rows[:2])))
    except Exception as e:
        log("WARNING", "批量失败样本读取失败：%s" % e)


def make_thumb_batch(items, edge, quality, timeout=None):
    """一次进程调用尽量多做几张缩略图。items = [(src, dst, rotate), ...]，返回成功生成的 dst 列表。

    Windows(PowerShell)：整批只起一个 PowerShell 进程。
    任务改为「一行一张、制表符分隔」的纯文本传输（不再用 JSON，避开 PS 5.1 的解析差异），
    并把每一张的失败原因原样带回日志。"""
    ok = []
    if not items:
        return ok
    if thumb_backend() != "powershell":
        for src, dst, rot in items:
            try:
                if make_thumb(src, dst, edge, quality, rot):
                    ok.append(dst)
            except Exception as e:
                log("WARNING", "缩略图生成失败：%s: %s" % (type(e).__name__, e))
        return ok
    jobs = []
    jobfile = failfile = statfile = dbgfile = None
    try:
        lines = []
        for src, dst, rot in items:
            tmp = "%s.%d.part" % (dst, threading.get_ident())
            jobs.append({"src": src, "dst": dst, "tmp": tmp})
            lines.append("\t".join([src, dst, tmp, str(int(edge)), str(int(rot))]))
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except OSError:
                pass
        fd, jobfile = tempfile.mkstemp(prefix="filmthumb_", suffix=".txt")
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as fh:
            fh.write("\n".join(lines))
        fd2, failfile = tempfile.mkstemp(prefix="filmthumbfail_", suffix=".txt")
        os.close(fd2)
        fd3, statfile = tempfile.mkstemp(prefix="filmthumbstat_", suffix=".json")
        os.close(fd3)
        fd4, dbgfile = tempfile.mkstemp(prefix="filmthumbdbg_", suffix=".txt")
        os.close(fd4)
        script = (
            "$ErrorActionPreference='Stop';"
            "Add-Type -AssemblyName System.Drawing;"
            "$wic=$false;try{Add-Type -AssemblyName PresentationCore -ErrorAction Stop;$wic=$true}catch{$wic=$false};"
            "$lines=[System.IO.File]::ReadAllLines($env:FILM_JOBS,[System.Text.Encoding]::UTF8);"
            "$dbg=@();$dbg+=('lines=' + $lines.Count);$dbg+=('wic=' + $wic);"
            "if($lines.Count -gt 0){$dbg+=('first=' + $lines[0])}"
            "$ci=[System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders()|Where-Object{$_.MimeType -eq 'image/jpeg'};"
            "if($ci -eq $null){$dbg+='codec=NULL'}"
            "$p=New-Object System.Drawing.Imaging.EncoderParameters 1;"
            "$p.Param[0]=New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality,[long]$env:FILM_QUAL);"
            "$bad=@();$stat=@();"
            "foreach($ln in $lines){"
            "try{"
            "$a=$ln -split ([char]9);"
            "if($a.Count -lt 5){throw ('fields=' + $a.Count)}"
            "$src=$a[0];$tmp=$a[2];$edg=[int]$a[3];$rot=[int]$a[4];"
            "$sw=[System.Diagnostics.Stopwatch]::StartNew();"
            "$used='gdi';$ok=$false;$iw=0;$ih=0;"
            "if($wic -and $rot -eq 0){try{"
            "$fr=[System.Windows.Media.Imaging.BitmapFrame]::Create((New-Object System.Uri($src)),[System.Windows.Media.Imaging.BitmapCreateOptions]::DelayCreation,[System.Windows.Media.Imaging.BitmapCacheOption]::None);"
            "$ow=[int]$fr.PixelWidth;$oh=[int]$fr.PixelHeight;$fr=$null;$iw=$ow;$ih=$oh;"
            "$bi=New-Object System.Windows.Media.Imaging.BitmapImage;$bi.BeginInit();$bi.UriSource=New-Object System.Uri($src);"
            "$bi.CacheOption=[System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad;"
            "if($ow -ge $oh){if($ow -gt $edg){$bi.DecodePixelWidth=$edg}}else{if($oh -gt $edg){$bi.DecodePixelHeight=$edg}};"
            "$bi.EndInit();"
            "$enc=New-Object System.Windows.Media.Imaging.JpegBitmapEncoder;$enc.QualityLevel=[int]$env:FILM_QUAL;"
            "$enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bi));"
            "$fs=[System.IO.File]::Create($tmp);try{$enc.Save($fs)}finally{$fs.Close()};$bi=$null;$enc=$null;$ok=$true;$used='wic'"
            "}catch{$ok=$false;$used='wicfail'}}"
            "if(-not $ok){"
            "$img=[System.Drawing.Image]::FromFile($src);"
            "$iw=[int]$img.Width;$ih=[int]$img.Height;"
            "try{"
            "$w=$iw;$h=$ih;$m=[Math]::Max($w,$h);"
            "if($m -gt $edg){$s=$edg/$m;$w=[int]($w*$s);$h=[int]($h*$s)}"
            "$bmp=New-Object System.Drawing.Bitmap $w,$h;"
            "try{"
            "$g=[System.Drawing.Graphics]::FromImage($bmp);"
            "try{"
            "$g.InterpolationMode=[System.Drawing.Drawing2D.InterpolationMode]::HighQualityBilinear;"
            "$g.DrawImage($img,0,0,$w,$h)"
            "}finally{$g.Dispose()}"
            "if($rot -ne 0){if($rot -gt 0){$bmp.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipNone)}else{$bmp.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipNone)}}"
            "$bmp.Save($tmp,$ci,$p)"
            "}finally{$bmp.Dispose()}"
            "}finally{$img.Dispose()}"
            "}"
            "$sw.Stop();"
            "$fb=0;try{$fb=[long](Get-Item -LiteralPath $src).Length}catch{};"
            "$stat+=[pscustomobject]@{dst=$tmp;ms=[int]$sw.ElapsedMilliseconds;w=$iw;h=$ih;bytes=$fb;via=$used}"
            "}catch{$bad+=('{0} :: {1}: {2}' -f $tmp, $_.Exception.GetType().Name, $_.Exception.Message);"
            "if($bad.Count -le 2){$dbg+=('FAIL exists=' + (Test-Path -LiteralPath $a[0]) + ' src=' + $a[0])}}"
            "}"
            "if($env:FILM_DBG){$dbg | Set-Content -Encoding UTF8 $env:FILM_DBG}"
            "if($env:FILM_STAT){try{$stat | ConvertTo-Json -Compress | Set-Content -Encoding UTF8 $env:FILM_STAT}catch{}}"
            "if($bad.Count -gt 0){$bad | Set-Content -Encoding UTF8 $env:FILM_FAIL; exit 2}"
            "exit 0"
        )
        if timeout is None:
            timeout = max(60, min(240, 30 + 4 * len(jobs)))
        proc = None
        try:
            proc = subprocess.Popen(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                                     "-Command", script], stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
                                    env=dict(os.environ, FILM_JOBS=jobfile, FILM_FAIL=failfile,
                                             FILM_STAT=statfile, FILM_DBG=dbgfile,
                                             FILM_QUAL=str(int(quality))))
            pending = dict((j["tmp"], j["dst"]) for j in jobs)
            t0 = time.time()
            rc = None
            while pending:
                _claim_tmp(pending)               # 出一张就立刻可见，前端会一张一张冒出来
                if not pending:
                    break
                rc = proc.poll()
                if rc is not None:
                    _claim_tmp(pending)
                    break
                if time.time() - t0 > timeout:
                    log("WARNING", "批量缩略图超时（%.0fs），还剩 %d 张没出来" % (timeout, len(pending)))
                    proc.kill()
                    break
                time.sleep(STREAM_POLL)
            try:
                rc = proc.wait(timeout=10)
            except Exception:
                try:
                    proc.kill()
                except Exception:
                    pass
                rc = -1
            if rc not in (0, None):
                try:
                    err = (proc.stderr.read() or b"").decode("utf-8", "ignore").strip().replace("\n", " ")[:400]
                except Exception:
                    err = ""
                log("WARNING", "批量缩略图进程退出码 %s%s" % (rc, ("；stderr: " + err) if err else ""))
        except Exception as e:
            log("WARNING", "批量缩略图进程异常：%s" % e)
            try:
                if proc:
                    proc.kill()
            except Exception:
                pass
        _log_ps_diag(dbgfile, failfile)
        _log_batch_stats(statfile)
        for j in jobs:
            try:
                if os.path.exists(j["tmp"]) and os.path.getsize(j["tmp"]) > 0:
                    os.replace(j["tmp"], j["dst"])
                    ok.append(j["dst"])
            except OSError as e:
                log("WARNING", "缩略图落盘失败：%s" % e)
    except Exception as e:
        log("WARNING", "批量缩略图生成出错：%s: %s" % (type(e).__name__, e))
    finally:
        for p in (jobfile, failfile, statfile, dbgfile):
            try:
                if p and os.path.exists(p):
                    os.remove(p)
            except OSError:
                pass
        for j in jobs:
            try:
                if os.path.exists(j["tmp"]):
                    os.remove(j["tmp"])
            except OSError:
                pass
    return ok



def thumb_rotate(force=False):
    """超过设置的大小上限（默认 200MB）就删最旧的缩略图。
       0.4.1：加最小间隔，避免每生成一张就全目录扫描一次。"""
    now = time.time()
    with _rotate_guard:
        if not force and (now - _last_rotate[0]) < THUMB_ROTATE_MIN_SEC:
            return
        _last_rotate[0] = now
    try:
        max_bytes, max_files = thumb_limits()
        items = []
        total = 0
        for f in os.listdir(THUMB):
            p = os.path.join(THUMB, f)
            try:
                st = os.stat(p)
            except OSError:
                continue
            items.append((st.st_mtime, st.st_size, f))
            total += st.st_size
        if len(items) <= max_files and total <= max_bytes:
            return
        items.sort()
        removed = 0
        while items and (len(items) > max_files or total > max_bytes):
            mt, sz, f = items.pop(0)
            try:
                os.remove(os.path.join(THUMB, f))
                total -= sz
                removed += 1
            except OSError:
                break
        if removed:
            log("INFO", "预览缓存超限（上限 %d MB），删除 %d 个最旧缩略图，现剩 %d 个 / %.1f MB"
                % (max_bytes // 1048576, removed, len(items), total / 1048576.0))
    except Exception as e:
        log("WARNING", "预览缓存轮换失败：%s" % e)

def thumb_stats():
    n, total = 0, 0
    try:
        for f in os.listdir(THUMB):
            try:
                total += os.stat(os.path.join(THUMB, f)).st_size
                n += 1
            except OSError:
                pass
    except OSError:
        pass
    max_bytes, max_files = thumb_limits()
    return {"dir": THUMB, "files": n, "bytes": total, "mb": round(total / 1048576.0, 1),
            "maxMB": max_bytes // 1048576, "maxFiles": max_files,
            "backend": thumb_backend() or "（无，将直接读原图）"}

def thumb_clear():
    n, total = 0, 0
    try:
        for f in os.listdir(THUMB):
            if not f.lower().endswith(".jpg"):
                continue
            p = os.path.join(THUMB, f)
            try:
                total += os.stat(p).st_size
                os.remove(p)
                n += 1
            except OSError:
                pass
    except OSError as e:
        log("ERROR", "清空预览缓存失败：%s" % e)
        return {"ok": False, "msg": str(e)}
    log("INFO", "已清空预览缩略图缓存：删除 %d 个文件，释放 %.1f MB" % (n, total / 1048576.0))
    return {"ok": True, "removed": n, "freed": round(total / 1048576.0, 1)}

# ==================== 缩略图：后台队列（0.4.2） ====================
# Windows 上「每张图起一个 PowerShell」+ 大图解码本身就慢（一卷几十秒）。
# 与其让用户对着空白等，不如：缺缩略图时**先返回原图**（浏览器自己缩放，立刻可见），
# 同时把整卷排队到后台，用「一个进程批量生成」，下次再看就走缓存。
_bg_q = queue.Queue()
_bg_pending = set()
_bg_guard = threading.Lock()
_bg_started = [False]
BG_START_DELAY = 1.5          # 先把 CPU/磁盘让给浏览器渲染原图


def _thumb_wait_mode():
    """True = 等缩略图生成完再返回（慢但省内存）；False = 先返回原图、后台生成。
       默认：Windows(PowerShell) 不等待，其它平台（sips 很快）等待；可用 settings.thumbWait 覆盖。"""
    try:
        v = load_settings().get("thumbWait")
    except Exception:
        v = None
    if v is None:
        return True          # 0.4.3：默认等待（先发原图对前端开销太大）
    return bool(v)


_batch_cool = {}
_dir_batch = {}
_dir_batch_guard = threading.Lock()
_thumb_demand = {}
_demand_seq = [0]
BATCH_COALESCE = 0.25
THUMB_WAIT_MAX = 120


def _batch_in_cooldown(d):
    return _batch_cool.get(d, 0.0) > time.time()


def _batch_set_cooldown(d, secs=600):
    _batch_cool[d] = time.time() + secs
    log("WARNING", "该目录批量生成连续失败，%d 分钟内直接用单张方式生成：%s"
        % (secs // 60, os.path.basename(d)))


def _bump_demand(dst):
    _demand_seq[0] += 1
    _thumb_demand[dst] = _demand_seq[0]


def _dir_batch_thread(d, edge, quality, lvl):
    """按「谁在被看就先做谁」排序，再把整卷切成若干份并行批量；每出一张立刻可见。"""
    time.sleep(BATCH_COALESCE)               # 先收集一下这一瞬间前端要哪几张
    try:
        if _batch_in_cooldown(d):
            return
        with _dir_lock(d):
            try:
                names = [f for f in sorted(os.listdir(d)) if f.lower().endswith(IMG_EXT)]
            except OSError:
                names = []
            imgs, _thumbs = split_thumbs(names)

            def prio(f):
                try:
                    dd = thumb_path(os.path.join(d, f), edge, quality) or ""
                except Exception:
                    dd = ""
                return -_thumb_demand.get(dd, 0)
            imgs.sort(key=prio)
            jobs = []
            for f in imgs:
                src = os.path.join(d, f)
                try:
                    d2 = thumb_path(src, edge, quality)
                except Exception:
                    d2 = None
                if not d2 or os.path.exists(d2):
                    continue
                jobs.append((src, d2, 0))
                if len(jobs) >= THUMB_BATCH_MAX:
                    break
            if not jobs:
                return
            wait_full_img(2.0)
            try:
                nproc = int(load_settings().get("thumbProcs") or 2)
            except Exception:
                nproc = 2
            nproc = max(1, min(4, nproc, len(jobs)))
            chunks = [jobs[i::nproc] for i in range(nproc)]
            t0 = time.time()
            counts = [0] * nproc

            def one(i):
                try:
                    counts[i] = len(make_thumb_batch(chunks[i], edge, quality,
                                                     max(60, min(300, 30 + 4 * len(chunks[i])))))
                except Exception as e:
                    log("WARNING", "批量分片失败：%s: %s" % (type(e).__name__, e))
            ts = [threading.Thread(target=one, args=(i,), daemon=True) for i in range(nproc) if chunks[i]]
            for t in ts:
                t.start()
            for t in ts:
                t.join()
            made = sum(counts)
            log("INFO", "批量生成预览缩略图：档%d/%dpx/%dq，%d/%d 张，%d 个进程并行，用时 %.1fs（%s）"
                % (lvl, edge, quality, made, len(jobs), len(ts), time.time() - t0, os.path.basename(d)))
            if made == 0:
                _batch_set_cooldown(d)
            thumb_rotate()
    except Exception as e:
        log("WARNING", "目录批量任务异常：%s: %s" % (type(e).__name__, e))
    finally:
        with _dir_batch_guard:
            _dir_batch.pop(d, None)



def _bg_worker():
    while True:
        d, edge, quality, lvl = _bg_q.get()
        try:
            time.sleep(BG_START_DELAY)
            _dir_batch_thread(d, edge, quality, lvl)
        except Exception as e:
            log("WARNING", "后台生成缩略图失败：%s: %s" % (type(e).__name__, e))
        finally:
            with _bg_guard:
                _bg_pending.discard(d)
            _bg_q.task_done()


def _bg_enqueue_dir(d, edge, quality, lvl):
    """把某个目录排进后台队列（同一目录只排一次）"""
    with _bg_guard:
        if d in _bg_pending:
            return False
        _bg_pending.add(d)
        if not _bg_started[0]:
            _bg_started[0] = True
            threading.Thread(target=_bg_worker, name="thumb-bg", daemon=True).start()
    _bg_q.put((d, edge, quality, lvl))
    log("INFO", "缩略图未就绪：先返回原图，已排队后台生成整卷 —— %s" % os.path.basename(d))
    return True


# ==================== 长驻 PowerShell 缩略图引擎（0.4.7） ====================
# Windows 上每次新起 powershell.exe 要 2~3 秒（加载 .NET / 程序集 / AMSI 扫描）。
# 这里让一个 PowerShell 进程**常驻**，逐张喂任务、逐张出结果：
#   · 冷启动只付一次（服务启动时预热，用户打开预览时已是热的）
#   · 按「谁在看就先做谁」的优先级出图，不预生成整卷
# 引擎起不来 / 中途挂掉时，自动退回 v0.4.6 的「按目录批量」，功能不受影响。
_ENG = {"proc": None, "q": None, "ready": False, "dead": False,
        "lock": threading.Lock(), "sguard": threading.Lock()}

_ENGINE_PS = r"""
$ErrorActionPreference='Continue'
[Console]::InputEncoding  = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
Add-Type -AssemblyName System.Drawing
$wic=$false
try{Add-Type -AssemblyName PresentationCore -ErrorAction Stop;$wic=$true}catch{$wic=$false}
$ci=[System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders()|Where-Object{$_.MimeType -eq 'image/jpeg'}
$p=New-Object System.Drawing.Imaging.EncoderParameters 1
$p.Param[0]=New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality,55)
[Console]::Out.WriteLine('READY' + [char]9 + $wic)
[Console]::Out.Flush()
while($true){
  $ln=[Console]::In.ReadLine()
  if($null -eq $ln){break}
  if($ln -eq 'QUIT'){break}
  $a=$ln -split ([char]9)
  if($a.Count -lt 5){[Console]::Out.WriteLine('ERR'+[char]9+'badline');[Console]::Out.Flush();continue}
  $src=$a[0];$tmp=$a[1];$edg=[int]$a[2];$q=[int]$a[3];$rot=[int]$a[4]
  $sw=[System.Diagnostics.Stopwatch]::StartNew()
  $used='gdi';$ok=$false
  try{
    if($wic -and $rot -eq 0){
      try{
        $fr=[System.Windows.Media.Imaging.BitmapFrame]::Create((New-Object System.Uri($src)),[System.Windows.Media.Imaging.BitmapCreateOptions]::DelayCreation,[System.Windows.Media.Imaging.BitmapCacheOption]::None)
        $ow=[int]$fr.PixelWidth;$oh=[int]$fr.PixelHeight;$fr=$null
        $bi=New-Object System.Windows.Media.Imaging.BitmapImage
        $bi.BeginInit();$bi.UriSource=New-Object System.Uri($src)
        $bi.CacheOption=[System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        if($ow -ge $oh){if($ow -gt $edg){$bi.DecodePixelWidth=$edg}}else{if($oh -gt $edg){$bi.DecodePixelHeight=$edg}}
        $bi.EndInit()
        $enc=New-Object System.Windows.Media.Imaging.JpegBitmapEncoder
        $enc.QualityLevel=$q
        $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bi))
        $fs=[System.IO.File]::Create($tmp)
        try{$enc.Save($fs)}finally{$fs.Close()}
        $bi=$null;$enc=$null;$ok=$true;$used='wic'
      }catch{$ok=$false;$used='wicfail'}
    }
    if(-not $ok){
      $p.Param[0]=New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality,[long]$q)
      $img=[System.Drawing.Image]::FromFile($src)
      try{
        $w=[int]$img.Width;$h=[int]$img.Height;$m=[Math]::Max($w,$h)
        if($m -gt $edg){$s=$edg/$m;$w=[int]($w*$s);$h=[int]($h*$s)}
        $bmp=New-Object System.Drawing.Bitmap $w,$h
        try{
          $g=[System.Drawing.Graphics]::FromImage($bmp)
          try{$g.InterpolationMode=[System.Drawing.Drawing2D.InterpolationMode]::HighQualityBilinear;$g.DrawImage($img,0,0,$w,$h)}finally{$g.Dispose()}
          if($rot -ne 0){if($rot -gt 0){$bmp.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipNone)}else{$bmp.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipNone)}}
          $bmp.Save($tmp,$ci,$p)
        }finally{$bmp.Dispose()}
      }finally{$img.Dispose()}
    }
    $sw.Stop()
    [Console]::Out.WriteLine('OK'+[char]9+$tmp+[char]9+$sw.ElapsedMilliseconds+[char]9+$used)
  }catch{
    [Console]::Out.WriteLine('ERR'+[char]9+$tmp+[char]9+($_.Exception.Message -replace [char]10,' '))
  }
  [Console]::Out.Flush()
}
"""


def _engine_alive():
    p = _ENG["proc"]
    return bool(p) and p.poll() is None and _ENG["ready"]


def _engine_start(timeout=25.0):
    """启动常驻引擎并等 READY；失败返回 False（之后一直走批量兜底）。"""
    if _engine_alive():
        return True
    with _ENG["sguard"]:
        if _engine_alive():
            return True
        if _ENG["dead"]:
            return False
        try:
            ps1 = os.path.join(THUMB, "thumb_engine.ps1")
            with open(ps1, "w", encoding="utf-8") as fh:
                fh.write(_ENGINE_PS)
        except OSError as e:
            log("WARNING", "写引擎脚本失败：%s" % e)
            _ENG["dead"] = True
            return False
        try:
            flags = 0x08000000 if os.name == "nt" else 0     # CREATE_NO_WINDOW
            proc = subprocess.Popen(
                ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", ps1],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                text=True, encoding="utf-8", errors="replace", bufsize=1, creationflags=flags)
        except Exception as e:
            log("WARNING", "启动缩略图引擎失败：%s" % e)
            _ENG["dead"] = True
            return False
        q = queue.Queue()

        def reader():
            try:
                for line in iter(proc.stdout.readline, ""):
                    if not line:
                        break
                    q.put(line.rstrip("\r\n"))
            except Exception:
                pass
            q.put(None)
        threading.Thread(target=reader, name="thumb-engine-r", daemon=True).start()
        t0 = time.time()
        ready_line = None
        while time.time() - t0 < timeout:
            try:
                line = q.get(timeout=0.5)
            except queue.Empty:
                if proc.poll() is not None:
                    break
                continue
            if line is None:
                break
            if line.startswith("READY"):
                ready_line = line
                break
        if not ready_line:
            try:
                proc.kill()
            except Exception:
                pass
            log("WARNING", "缩略图引擎握手失败（%.0fs 内没有 READY），改用批量兜底" % timeout)
            _ENG["dead"] = True
            return False
        _ENG["proc"] = proc
        _ENG["q"] = q
        _ENG["ready"] = True
        log("INFO", "缩略图引擎已就绪：%s" % ready_line)
        return True


def _engine_do(src, dst, edge, quality, rot=0, timeout=45.0):
    """用常驻引擎生成一张，返回 (ok, ms, via)。"""
    if not _engine_alive():
        return (False, 0, "")
    tmp = "%s.%d.part" % (dst, threading.get_ident())
    try:
        if os.path.exists(tmp):
            os.remove(tmp)
    except OSError:
        pass
    line = "\t".join([src, tmp, str(int(edge)), str(int(quality)), str(int(rot))])
    q = _ENG["q"]
    with _ENG["lock"]:
        p = _ENG["proc"]
        if not p or p.poll() is not None:
            _ENG["ready"] = False
            return (False, 0, "")
        try:
            p.stdin.write(line + "\n")
            p.stdin.flush()
        except Exception as e:
            log("WARNING", "引擎写入失败：%s" % e)
            _ENG["ready"] = False
            return (False, 0, "")
        t0 = time.time()
        got = None
        while time.time() - t0 < timeout:
            try:
                r = q.get(timeout=0.5)
            except queue.Empty:
                if p.poll() is not None:
                    break
                continue
            if r is None:
                break
            parts = r.split("\t")
            if parts and parts[0] == "OK" and len(parts) > 1 and parts[1] == tmp:
                ms = int(parts[2]) if len(parts) > 2 and parts[2].isdigit() else 0
                got = (ms, parts[3] if len(parts) > 3 else "")
                break
            if parts and parts[0] == "ERR":
                log("WARNING", "引擎生成失败：%s" % (parts[2] if len(parts) > 2 else r))
                return (False, 0, "")
        if not got:
            log("WARNING", "缩略图引擎无响应/已退出，标记为不可用")
            _ENG["ready"] = False
            try:
                p.kill()
            except Exception:
                pass
            return (False, 0, "")
    try:
        if os.path.exists(tmp) and os.path.getsize(tmp) > 0:
            os.replace(tmp, dst)
            return (True, got[0], got[1])
    except OSError as e:
        log("WARNING", "引擎缩略图落盘失败：%s" % e)
    return (False, 0, "")


_eq = []
_eqset = set()
_eqguard = threading.Lock()
_eqseq = [0]
_eqworker = [False]


def _engine_enqueue(src, dst, edge, quality):
    """按需求优先级入队（谁在看排前面）"""
    with _eqguard:
        if os.path.exists(dst) or dst in _eqset:
            return
        _eqseq[0] += 1
        heapq.heappush(_eq, (-_thumb_demand.get(dst, 0), _eqseq[0], (src, dst, edge, quality)))
        _eqset.add(dst)
        if not _eqworker[0]:
            _eqworker[0] = True
            threading.Thread(target=_engine_worker, name="thumb-engine-w", daemon=True).start()


def _engine_worker():
    while True:
        with _eqguard:
            item = heapq.heappop(_eq) if _eq else None
        if not item:
            time.sleep(0.15)
            continue
        _, _, (src, dst, edge, quality) = item
        with _eqguard:
            _eqset.discard(dst)
        if os.path.exists(dst):
            _notify_thumb(dst)
            continue
        ok = False
        if _engine_alive():
            ok = _engine_do(src, dst, edge, quality)[0]
        if not ok:                      # 引擎不可用 → 单张兜底（慢但能用）
            try:
                ok = bool(make_thumb(src, dst, edge, quality))
            except Exception as e:
                log("WARNING", "引擎兜底单张失败：%s: %s" % (type(e).__name__, e))
        if ok and os.path.exists(dst):
            _notify_thumb(dst)


def _ensure_thumbs(path_p, edge, quality, lvl):
    """确保 path_p 的缩略图存在。

    0.4.7：若常驻引擎已就绪 → 直接把这张丢给它（按需求排序，一张张出）；
    引擎不可用 → 退回 0.4.6 的「按目录批量」；再不行 → 单张兜底。"""
    try:
        dst = thumb_path(path_p, edge, quality)
    except Exception as e:
        log("WARNING", "缩略图路径计算失败：%s" % e)
        return
    if not dst or os.path.exists(dst):
        return
    if thumb_backend() != "powershell":
        with _thumb_sem:
            if os.path.exists(dst):
                return
            try:
                waited = wait_full_img(6.0)
                if waited > 0.5:
                    log("INFO", "缩略图让路给原图读取 %.1fs：%s" % (waited, os.path.basename(path_p)))
                if make_thumb(path_p, dst, edge, quality):
                    log("INFO", "生成预览缩略图：档%d/%dpx/%dq %s → %s"
                        % (lvl, edge, quality, os.path.basename(path_p), os.path.basename(dst)))
                    thumb_rotate()
            except Exception as e:
                log("WARNING", "缩略图生成失败（回退原图）：%s: %s" % (type(e).__name__, e))
        return
    d = os.path.dirname(os.path.abspath(path_p))
    _bump_demand(dst)                        # 记住「这张正在被看」，优先做
    if not _engine_alive() and not _ENG["dead"]:      # 引擎正在预热：稍等一下再决定
        t0 = time.time()
        while time.time() - t0 < 6.0 and not _engine_alive() and not _ENG["dead"]:
            time.sleep(0.2)
    if _engine_alive():
        _engine_enqueue(path_p, dst, edge, quality)
        if _wait_thumb(dst, THUMB_WAIT_MAX) or os.path.exists(dst):
            return
        if os.path.exists(dst):
            return
    if not _batch_in_cooldown(d):
        with _dir_batch_guard:
            running = d in _dir_batch
            if not running:
                _dir_batch[d] = True
        if not running:
            threading.Thread(target=_dir_batch_thread, args=(d, edge, quality, lvl), daemon=True).start()
    if _wait_thumb(dst, THUMB_WAIT_MAX):
        return
    if not os.path.exists(dst):              # 超时或都不可用 → 单张兜底
        try:
            if make_thumb(path_p, dst, edge, quality):
                log("INFO", "已改用单张方式生成：%s" % os.path.basename(path_p))
                thumb_rotate()
        except Exception as e:
            log("WARNING", "单张生成失败（回退原图）：%s: %s" % (type(e).__name__, e))



def serve_thumb(path_p, qs):
    """返回 (bytes, content_type, cache_header)。

    0.4.2：缺缩略图时，Windows 默认**先返回原图**（浏览器自己缩放，立刻可见），
    同时把整卷排进后台队列慢慢生成；下次再看这一卷就走缓存。这样打开预览不再干等。"""
    ext = os.path.splitext(path_p)[1].lower()
    on, lvl, edge, quality = thumb_params()
    dst = None
    if on:
        try:
            dst = thumb_path(path_p, edge, quality)
        except Exception as e:
            log("WARNING", "缩略图路径计算失败：%s" % e)
            dst = None
        if dst and not os.path.exists(dst):
            if thumb_backend() is None:
                log("WARNING", "缩略图生成不可用（没有可用的降采样工具），回退原图：%s"
                    % os.path.basename(path_p))
                dst = None
            elif _thumb_wait_mode():
                _ensure_thumbs(path_p, edge, quality, lvl)
            else:
                _bg_enqueue_dir(os.path.dirname(os.path.abspath(path_p)), edge, quality, lvl)
    if dst and os.path.exists(dst):
        with open(dst, "rb") as fh:
            return fh.read(), "image/jpeg", "private, max-age=3600"
    with open(path_p, "rb") as fh:
        return fh.read(), MIME.get(ext, "application/octet-stream"), "no-store"

# ==================== 导出 Index Paper（生成可打印的 docx） ====================
# 规则（与 000_INDEX PAPER/ 里既有的 docx 保持一致）：
#   · 标题 = 胶卷文件夹名；表头两列三行：型号/编号、相机/拍摄日期、镜头/冲洗日期
#   · 正文沿用模板里的「图片阵列」段落样式（制表位 2520 / 5040 / 7560 twips）
#   · 135：每行 4 张，图片宽度 4.19cm；120：每行 3 张，图片宽度 5.59cm
#   · 竖图（高 > 宽）旋转 90°；旋转后一律按宽度等比缩放，保证每张图宽度一致
IP_ORDER = ["011_scanPost", "010_scan", "021_camPost"]
IP_TPL_NAME = "index_template.docx"                       # assets/ 下的模板副本（优先）
IP_TPL_FALLBACK = os.path.join(ROOT, "000_INDEX PAPER", "000_胶卷Index模板.docx")
IP_PER_ROW = {"135": 4, "120": 3}
IP_WIDTH_CM = {"135": 4.19, "120": 5.59}
# 同一行图片之间的分隔：135 = 制表位；120 = 两个空格
IP_SEP_TAB = '<w:r><w:tab/></w:r>'
IP_SEP_SP2 = ('<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
              '<w:t xml:space="preserve">  </w:t></w:r>')
IP_EDGE_DEFAULT, IP_Q_DEFAULT, IP_ROT_DEFAULT = 1000, 76, "cw"
EMU_CM = 360000
_docx_lock = threading.Lock()
_docx_cache = {}
_DOC_HEAD = ('<w:rPr><w:rFonts w:ascii="Amasis MT Pro Black" w:hAnsi="Amasis MT Pro Black"/></w:rPr>')
_EMOJI_HEAD = ('<w:rPr><w:rFonts w:ascii="Segoe UI Emoji" w:hAnsi="Segoe UI Emoji" '
               'w:cs="Segoe UI Emoji"/></w:rPr>')


def ip_template_path():
    p = os.path.join(ASSETS, IP_TPL_NAME)
    if os.path.exists(p):
        return p
    return IP_TPL_FALLBACK if os.path.exists(IP_TPL_FALLBACK) else None


def ip_out_dir():
    """导出目录：设置 ipDir（相对 FILM 目录，或绝对路径）"""
    d = (load_settings().get("ipDir") or "000_INDEX PAPER").strip() or "000_INDEX PAPER"
    return os.path.normpath(d if os.path.isabs(d) else os.path.join(ROOT, d))


def ip_dir_rel():
    p = ip_out_dir()
    try:
        return os.path.relpath(p, ROOT)
    except ValueError:
        return p


def xml_esc(s):
    return (str(s if s is not None else "").replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def img_size(path):
    """纯标准库读 JPEG / PNG 的像素尺寸；读不到返回 None"""
    try:
        with open(path, "rb") as fh:
            data = fh.read(600000)
    except OSError:
        return None
    if len(data) >= 24 and data[:8] == b"\x89PNG\r\n\x1a\n":
        return (struct.unpack(">I", data[16:20])[0], struct.unpack(">I", data[20:24])[0])
    if data[:2] != b"\xff\xd8":
        return None
    i = 2
    while i + 9 < len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        m = data[i + 1]
        if m in (0xD8, 0x01) or 0xD0 <= m <= 0xD7:
            i += 2
            continue
        if m in (0xDA, 0xD9):
            return None
        ln = struct.unpack(">H", data[i + 2:i + 4])[0]
        if 0xC0 <= m <= 0xCF and m not in (0xC4, 0xC8, 0xCC):
            h, w = struct.unpack(">HH", data[i + 5:i + 9])
            return (w, h)
        i += 2 + ln
    return None


def ip_date(v):
    """2025-01-03 → 2025/01/03（不是这个格式就原样返回）"""
    v = str(v or "").strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", v)
    if not m:
        return v
    return "%s/%s/%s" % (m.group(1), m.group(2).zfill(2), m.group(3).zfill(2))


def ip_tmp_root():
    p = os.path.join(cache_root(), "indexpaper")
    try:
        os.makedirs(p, exist_ok=True)
    except OSError:
        p = tempfile.gettempdir()
    return p


def ip_parts():
    """读模板 docx（缓存）；改模板文件后自动重载"""
    tpl = ip_template_path()
    if not tpl:
        raise RuntimeError("找不到 docx 模板：assets/%s 或 000_INDEX PAPER/000_胶卷Index模板.docx"
                           % IP_TPL_NAME)
    key = (tpl, os.path.getmtime(tpl), os.path.getsize(tpl))
    with _docx_lock:
        if _docx_cache.get("key") == key:
            return _docx_cache["data"]
    z = zipfile.ZipFile(tpl)
    data = {n: z.read(n) for n in z.namelist() if not n.endswith("/")}
    z.close()
    with _docx_lock:
        _docx_cache.clear()
        _docx_cache["key"] = key
        _docx_cache["data"] = data
    return data


def _ip_run(text, emoji=False):
    rpr = _EMOJI_HEAD if emoji else _DOC_HEAD
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, xml_esc(text))


def _ip_cell(emoji, label, value):
    body = _ip_run(emoji, True) + _ip_run(" " + label + ": ")
    if value:
        body += _ip_run(str(value))
    return ('<w:tc><w:tcPr><w:tcW w:w="5021" w:type="dxa"/></w:tcPr>'
            '<w:p><w:pPr>%s</w:pPr>%s</w:p></w:tc>' % (_DOC_HEAD, body))


def _ip_table(rec):
    d1, d2 = ip_date(rec.get("d1")), ip_date(rec.get("d2"))
    cap = d1 if (not d2 or d2 == d1) else (d1 + "~" + d2)
    if not d1:
        cap = d2
    dev = ip_date(rec.get("dev"))
    note = str(rec.get("devnote") or "").strip()
    if dev and note:
        dev = dev + " （" + note + "）"
    rows = [("🎞️", "Film", rec.get("model")), ("🗂️", "Index", rec.get("code")),
            ("📷", "Cam", rec.get("cam")), ("📅", "Capture Date", cap),
            ("🔭", "Lens", rec.get("lens")), ("🖨️", "Processed Date", dev)]
    tr = ""
    for k in range(0, 6, 2):
        tr += "<w:tr>" + _ip_cell(*rows[k]) + _ip_cell(*rows[k + 1]) + "</w:tr>"
    return ('<w:tbl><w:tblPr><w:tblStyle w:val="af4"/><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders>'
            '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" '
            'w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="5021"/><w:gridCol w:w="5021"/></w:tblGrid>'
            + tr + "</w:tbl>")


def _ip_drawing(idx, cx, cy):
    did = 100 + idx
    return ('<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:drawing>'
            '<wp:inline distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%(cx)d" cy="%(cy)d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:docPr id="%(id)d" name="图片 %(id)d"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:nvPicPr><pic:cNvPr id="%(id)d" name="image%(n)d.jpeg"/><pic:cNvPicPr/>'
            '</pic:nvPicPr><pic:blipFill><a:blip r:embed="rId%(id)d"/>'
            '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%(cx)d" cy="%(cy)d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
            % {"id": did, "n": idx + 1, "cx": cx, "cy": cy})


def _ip_document(rec, media, per_row, sep):
    base = ip_parts()
    raw = base["word/document.xml"].decode("utf-8")
    head = raw[:raw.index("<w:body>") + len("<w:body>")]
    s0 = raw.index("<w:sectPr")
    sect = raw[s0:raw.index("</w:sectPr>", s0) + len("</w:sectPr>")]
    out = [head,
           '<w:p><w:pPr><w:pStyle w:val="1"/><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr>'
           '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t xml:space="preserve">%s</w:t>'
           '</w:r></w:p>' % xml_esc(rec.get("name") or ""),
           _ip_table(rec)]
    for i in range(0, len(media), per_row):
        chunk = media[i:i + per_row]
        runs = sep.join(_ip_drawing(i + k, it[1], it[2])
                        for k, it in enumerate(chunk))
        out.append('<w:p><w:pPr><w:pStyle w:val="ae"/><w:spacing w:after="156"/>'
                   '<w:ind w:left="420" w:hanging="420"/></w:pPr>%s</w:p>' % runs)
    out.append(sect + "</w:body></w:document>")
    return "".join(out).encode("utf-8")


def _ip_rels(n):
    base = ip_parts()
    s = base["word/_rels/document.xml.rels"].decode("utf-8")
    extra = "".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/image" Target="media/image%d.jpeg"/>'
                    % (100 + i, i + 1) for i in range(n))
    return s.replace("</Relationships>", extra + "</Relationships>").encode("utf-8")


def _ip_content_types():
    base = ip_parts()
    s = base["[Content_Types].xml"].decode("utf-8")
    if 'Extension="jpeg"' not in s:
        s = s.replace('<Default Extension="xml"',
                      '<Default Extension="jpeg" ContentType="image/jpeg"/>'
                      '<Default Extension="xml"', 1)
    return s.encode("utf-8")


_ROOT_LIST = {"t": 0.0, "v": []}


def _root_list():
    """胶卷目录（FILM）下的顶层名字，带 30 秒缓存 —— 避免每次检查都遍历整个目录。"""
    if time.time() - _ROOT_LIST["t"] < 30 and _ROOT_LIST["v"]:
        return _ROOT_LIST["v"]
    try:
        v = sorted(os.listdir(ROOT))
    except OSError:
        v = []
    _ROOT_LIST["t"] = time.time()
    _ROOT_LIST["v"] = v
    return v


def _invalidate_root_list():
    _ROOT_LIST["t"] = 0.0
    _ROOT_LIST["v"] = []


def ip_find_folder(seq, name):
    """定位胶卷文件夹；文件夹被改名时按「序号_index」前缀兜底"""
    p = safe_under(ROOT, name) if name else None
    if p and os.path.isdir(p):
        return p
    pref = "%s_index" % seq
    try:
        for f in _root_list():
            if f.startswith(pref) and os.path.isdir(os.path.join(ROOT, f)):
                return os.path.join(ROOT, f)
    except OSError:
        pass
    return None


def roll_name(cells):
    """按「序号_index 编号_型号_YYMMDD[-YYMMDD]」算文件夹名；信息不足返回空串。"""
    if not cells or len(cells) < 8:
        return ""
    seq = str(cells[0] or "").strip()
    code = str(cells[1] or "").strip()
    model = str(cells[2] or "").strip()
    d1 = str(cells[6] or "").strip()
    d2 = str(cells[7] or "").strip()
    if not (seq and model and re.fullmatch(r"\d{4}-\d{2}-\d{2}", d1 or "")):
        return ""

    def ymd(s):
        return s[2:4] + s[5:7] + s[8:10]

    name = "%s_index %s_%s_%s" % (seq, code, model, ymd(d1))
    if d2 and d2 != d1 and re.fullmatch(r"\d{4}-\d{2}-\d{2}", d2):
        name += "-" + ymd(d2)
    return name


def _locate_roll_dir(seq, old_cells=None, prefer_name=None):
    """找硬盘上这一卷的文件夹：先按给定名字，再按旧名字，最后按「序号_index」前缀兜底。"""
    for nm in (prefer_name, roll_name(old_cells) if old_cells else ""):
        if nm:
            p = safe_under(ROOT, nm)
            if p and os.path.isdir(p):
                return p
    if seq:
        return ip_find_folder(seq, None)
    return None


def folder_plan(seq, cells, old_cells=None):
    """编辑后是否需要把硬盘文件夹改名（只判断，不动手）。"""
    target = roll_name(cells)
    if not target:
        return {"need": False, "skip": True,
                "msg": "信息不完整（需要序号 / 编号 / 型号 / 拍摄开始），未处理文件夹"}
    old_name = roll_name(old_cells) if old_cells else ""
    cur = _locate_roll_dir(seq, old_cells, prefer_name=old_name)
    if not cur:
        return {"need": False, "missing": True, "name": target,
                "msg": "硬盘上没找到序号 %s 的文件夹（应为：%s）" % (seq, target)}
    old = os.path.basename(cur)
    if old == target:
        return {"need": False, "same": True, "old": old, "name": target}
    dst = safe_under(ROOT, target)
    if not dst:
        return {"need": False, "bad": True, "old": old, "name": target,
                "msg": "目标文件夹名不合法：%s" % target}
    if os.path.exists(dst):
        return {"need": False, "conflict": True, "old": old, "name": target,
                "msg": "已存在同名文件夹「%s」，不会改名，请手工处理" % target}
    return {"need": True, "old": old, "name": target, "seq": seq}


def folder_sync(seq, cells, old_cells=None):
    """按 folder_plan 的判断真正执行改名（由前端确认后调用）。"""
    plan = folder_plan(seq, cells, old_cells)
    if not plan.get("need"):
        out = {"ok": True, "changed": False}
        out.update(plan)
        if plan.get("conflict") or plan.get("bad"):
            out["ok"] = False
        return out
    cur = _locate_roll_dir(seq, old_cells,
                           prefer_name=roll_name(old_cells) if old_cells else None)
    dst = safe_under(ROOT, plan["name"])
    if not cur or not dst:
        return {"ok": False, "changed": False, "old": plan.get("old"), "name": plan["name"],
                "msg": "定位文件夹失败，未改名"}
    old = os.path.basename(cur)
    try:
        os.rename(cur, dst)
    except OSError as e:
        log("ERROR", "文件夹改名失败：%s → %s（%s）" % (old, plan["name"], e))
        return {"ok": False, "changed": False, "old": old, "name": plan["name"],
                "msg": "文件夹改名失败：%s" % e}
    log("INFO", "文件夹已同步改名：%s → %s" % (old, plan["name"]))
    _invalidate_root_list()
    try:
        _LIST_CACHE.clear()
    except Exception:
        pass
    return {"ok": True, "changed": True, "old": old, "name": plan["name"]}


def folder_info(seq, name):
    """给界面用：这一卷的文件夹在不在、名字一不一致。"""
    seq = str(seq or "").strip()
    name = str(name or "").strip()
    cur = None
    if name:
        p = safe_under(ROOT, name)
        if p and os.path.isdir(p):
            cur = p
    if cur is None and seq:
        cur = ip_find_folder(seq, None)
    found = os.path.basename(cur) if cur else ""
    return {"ok": True, "expected": name, "found": found, "exists": bool(cur),
            "same": bool(cur) and bool(name) and found == name}


def folder_delete(seq, name):
    """把某卷的文件夹移入废纸篓 / 回收站（可恢复）。全平台兜底为移到 ROOT/_已删除/。"""
    cur = ip_find_folder(seq, name)
    if not cur:
        return {"ok": False, "msg": "硬盘上没找到该卷的文件夹，未删除"}
    base = os.path.basename(cur)
    if not safe_under(ROOT, base):
        return {"ok": False, "msg": "路径不合法，未删除"}
    if sys.platform == "darwin":
        try:
            sc = 'tell application "Finder" to delete POSIX file "%s"' % cur.replace('"', '\\"')
            subprocess.run(["osascript", "-e", sc], check=True, timeout=60,
                           stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
            log("WARNING", "已把文件夹移入废纸篓：%s" % base)
            _invalidate_root_list()
            return {"ok": True, "name": base, "how": "废纸篓"}
        except Exception as e:
            log("WARNING", "移入废纸篓失败（%s），改用 _已删除 兜底" % e)
    if os.name == "nt":
        try:
            ps = ("Add-Type -AssemblyName Microsoft.VisualBasic;"
                  "[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory("
                  "$env:FILM_DEL,'OnlyErrorDialogs','SendToRecycleBin')")
            pr = subprocess.run(["powershell", "-NoProfile", "-NonInteractive",
                                 "-ExecutionPolicy", "Bypass", "-Command", ps],
                                timeout=120, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
                                env=dict(os.environ, FILM_DEL=cur))
            if pr.returncode == 0:
                log("WARNING", "已把文件夹移入回收站：%s" % base)
                _invalidate_root_list()
                return {"ok": True, "name": base, "how": "回收站"}
            err = (pr.stderr or b"").decode("utf-8", "ignore").strip().replace("\n", " ")[:200]
            log("WARNING", "移入回收站失败（%s），改用 _已删除 兜底" % err)
        except Exception as e:
            log("WARNING", "移入回收站异常（%s），改用 _已删除 兜底" % e)
    try:
        junk = os.path.join(ROOT, "_已删除")
        os.makedirs(junk, exist_ok=True)
        dst = os.path.join(junk, "%s-%s" % (base, dt.datetime.now().strftime("%Y%m%d-%H%M%S")))
        os.rename(cur, dst)
        log("WARNING", "文件夹已移到 %s：%s" % (junk, base))
        _invalidate_root_list()
        return {"ok": True, "name": base, "how": "_已删除"}
    except OSError as e:
        log("ERROR", "移除文件夹失败：%s（%s）" % (base, e))
        return {"ok": False, "msg": "移除文件夹失败：%s" % e}


def folder_mkdir(cells, subs=False):
    """在胶卷目录下创建这一卷的文件夹；subs=True 时同时建标准子文件夹。"""
    target = roll_name(cells)
    if not target:
        return {"ok": False, "msg": "信息不完整（需要序号 / 编号 / 型号 / 拍摄开始），无法创建"}
    dst = safe_under(ROOT, target)
    if not dst:
        return {"ok": False, "msg": "文件夹名不合法：%s" % target}
    existed = os.path.exists(dst)
    made = []
    try:
        os.makedirs(dst, exist_ok=True)
        if subs:
            for s in IP_ORDER + ["020_cam"]:
                p = os.path.join(dst, s)
                if not os.path.exists(p):
                    os.makedirs(p)
                    made.append(s)
    except OSError as e:
        log("ERROR", "创建文件夹失败：%s（%s）" % (target, e))
        return {"ok": False, "msg": "创建失败：%s" % e}
    log("INFO", "%s文件夹：%s%s" % ("已存在" if existed else "已创建", target,
                                    ("；同时建子目录 " + "、".join(made)) if made else ""))
    _invalidate_root_list()
    try:
        _LIST_CACHE.clear()
    except Exception:
        pass
    return {"ok": True, "existed": existed, "name": target, "made": made}


def ip_build_docx(rec, opt, outdir=None):
    """生成一份 Index Paper docx。返回结果字典（不抛异常）"""
    t0 = time.time()
    seq = str(rec.get("seq") or "").strip()
    folder = ip_find_folder(seq, rec.get("name"))
    if not folder:
        return {"ok": False, "seq": seq, "msg": "找不到胶卷文件夹：%s" % (rec.get("name") or "(空)"),
                "hint": "请确认该条目在 FILM 目录下的文件夹名（或先「⟳ 同步」）"}
    name = os.path.basename(folder)
    rec = dict(rec)
    rec["name"] = name
    sub, files, tfs = "", [], []
    for s in IP_ORDER:
        p = os.path.join(folder, s)
        if os.path.isdir(p):
            raw = sorted(f for f in os.listdir(p) if f.lower().endswith(IMG_EXT))
            fs, tfs = split_thumbs(raw)          # 缩略图不进 Index Paper
            if fs:
                sub, files = s, fs
                break
    if not files:
        return {"ok": False, "seq": seq, "msg": "文件夹里没找到图片（已查 %s）" % " / ".join(IP_ORDER)}
    size = str(rec.get("size") or "").strip()
    per_row = IP_PER_ROW.get(size, IP_PER_ROW["135"])
    w_cm = IP_WIDTH_CM.get(size, IP_WIDTH_CM["135"])
    # 120：同一行图片用两个空格分隔（不再插制表位）；135：制表位对齐
    sep = IP_SEP_SP2 if size == "120" else IP_SEP_TAB
    try:
        edge = max(200, min(4000, int(float(opt.get("edge") or IP_EDGE_DEFAULT))))
        qual = max(30, min(95, int(float(opt.get("q") or IP_Q_DEFAULT))))
    except (TypeError, ValueError):
        edge, qual = IP_EDGE_DEFAULT, IP_Q_DEFAULT
    rot = str(opt.get("rot") or IP_ROT_DEFAULT).lower()
    outdir = outdir or ip_out_dir()
    dst = os.path.join(outdir, name + ".docx")
    try:
        os.makedirs(outdir, exist_ok=True)
    except OSError as e:
        return {"ok": False, "seq": seq, "msg": "无法创建导出目录 %s：%s" % (outdir, e),
                "hint": "可在设置里改 ipDir，或先手动建好该文件夹"}
    if opt.get("skip") and os.path.exists(dst) and os.path.getsize(dst) > 0:
        return {"ok": True, "skipped": True, "seq": seq, "file": name + ".docx", "dir": outdir,
                "dirRel": ip_dir_rel(), "images": len(files), "bytes": os.path.getsize(dst)}
    cx = int(round(w_cm * EMU_CM))
    try:
        tmpdir = tempfile.mkdtemp(prefix="ip_", dir=ip_tmp_root())
    except OSError as e:                      # 系统缓存目录不可写：回退到系统临时目录
        log("WARNING", "Index Paper：%s 不可写（%s），改用系统临时目录" % (ip_tmp_root(), e))
        tmpdir = tempfile.mkdtemp(prefix="ip_")
    part = dst + ".part"
    media, rotated, failed = [], 0, 0
    try:
        for i, f in enumerate(files):
            src = os.path.join(folder, sub, f)
            tmp = os.path.join(tmpdir, "%04d.jpg" % i)
            rr = 0
            if rot in ("cw", "ccw"):
                wh = img_size(src)
                if wh and wh[1] > wh[0]:
                    rr = 90 if rot == "cw" else -90
            if not make_thumb(src, tmp, edge, qual, rr):
                failed += 1
                log("WARNING", "Index Paper：降采样失败，跳过 %s（%s）" % (f, name))
                continue
            wh = img_size(tmp) or img_size(src)
            if not wh or not wh[0]:
                failed += 1
                continue
            if rr:
                rotated += 1
            media.append((tmp, cx, int(round(cx * wh[1] / float(wh[0])))))
        if not media:
            return {"ok": False, "seq": seq, "msg": "没有任何图片能读取（降采样工具不可用？）",
                    "hint": "Mac 用自带 sips；Windows 需要 PowerShell"}
        with zipfile.ZipFile(part, "w", zipfile.ZIP_DEFLATED) as z:
            for n, b in ip_parts().items():
                if n in ("word/document.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"):
                    continue
                z.writestr(n, b)
            for i, (p, _cx, _cy) in enumerate(media):
                z.write(p, "word/media/image%d.jpeg" % (i + 1))
            z.writestr("word/document.xml", _ip_document(rec, media, per_row, sep))
            z.writestr("word/_rels/document.xml.rels", _ip_rels(len(media)))
            z.writestr("[Content_Types].xml", _ip_content_types())
        os.replace(part, dst)
        nb = os.path.getsize(dst)
        log("INFO", "Index Paper：生成 %s（%d 张图，%s 每行 %d 张、宽 %.2fcm%s%s%s）耗时 %.1fs，%.1f MB"
            % (name + ".docx", len(media), size or "135", per_row, w_cm,
               "，旋转 %d 张" % rotated if rotated else "",
               "，跳过 %d 张" % failed if failed else "",
               "，已排除 %d 张缩略图" % len(tfs) if tfs else "",
               time.time() - t0, nb / 1048576.0))
        return {"ok": True, "seq": seq, "folder": name, "file": name + ".docx", "dir": outdir,
                "dirRel": ip_dir_rel(), "images": len(media), "rotated": rotated,
                "thumbs": len(tfs), "failed": failed,
                "rows": (len(media) + per_row - 1) // per_row,
                "bytes": nb, "seconds": round(time.time() - t0, 1)}
    except Exception as e:
        log("ERROR", "Index Paper：生成 %s 失败：%s: %s" % (name, type(e).__name__, e))
        return {"ok": False, "seq": seq, "msg": "%s: %s" % (type(e).__name__, e),
                "hint": "若该 docx 正在 Word 里打开，请先关闭再重试"}
    finally:
        try:
            if os.path.exists(part):
                os.remove(part)
        except OSError:
            pass
        shutil.rmtree(tmpdir, ignore_errors=True)


def ip_info():
    tpl = ip_template_path()
    outdir = ip_out_dir()
    exists = []
    if os.path.isdir(outdir):
        try:
            exists = sorted(f[:-5] for f in os.listdir(outdir)
                            if f.lower().endswith(".docx") and not f.startswith("~$"))
        except OSError:
            pass
    s = load_settings()
    return {"ok": True, "dir": outdir, "dirRel": ip_dir_rel(),
            "tpl": os.path.basename(tpl) if tpl else "", "tplOK": bool(tpl),
            "exists": exists, "perRow": IP_PER_ROW, "width": IP_WIDTH_CM,
            "opts": {"edge": int(s.get("ipEdge") or IP_EDGE_DEFAULT),
                     "q": int(s.get("ipQ") or IP_Q_DEFAULT),
                     "rot": s.get("ipRot") or IP_ROT_DEFAULT,
                     "skip": bool(s.get("ipSkip", True))}}


# ==================== HTTP 服务 ====================
def find_html():
    p = os.path.join(ASSETS, "index.html")
    if os.path.exists(p):
        return p
    for f in sorted(os.listdir(ASSETS)):
        if f.lower().endswith(".html"):
            return os.path.join(ASSETS, f)
    return p

IMG_EXT = (".jpg", ".jpeg", ".png", ".heic", ".heif", ".hif")
# 冲洗店给的缩略图 / 索引图：预览不显示、统计不计入、导出 Index Paper 不导出
THUMB_WORDS = ("thumb", "缩略")

def split_thumbs(files):
    """把同一目录的图片分成 (正式图, 缩略图)。判定规则：
       1) 文件名（小写）里含 thumb / 缩略——如 Thumbnails.jpg、thumb_01.jpg；
       2) 主名是另一张图主名的严格前缀，且多出来的部分全是数字（≥ 2 位）——
          即冲洗店常见的「整卷缩略图」：082 的 00012964.jpg 对应 000129640001.jpg …；
          （要求主名 ≥ 6 位、多出 ≥ 2 位、且多出部分纯数字，避免 1.jpg/10.jpg 这类误判）"""
    stems = {f: os.path.splitext(f)[0] for f in files}
    imgs, thumbs = [], []
    for f in files:
        s = stems[f]
        low = f.lower()
        if any(w in low for w in THUMB_WORDS):
            thumbs.append(f)
            continue
        if len(s) >= 6:
            is_thumb = False
            for g in files:
                if g == f:
                    continue
                t = stems[g]
                if t.startswith(s):
                    rest = t[len(s):]
                    if len(rest) >= 2 and rest.isdigit():
                        is_thumb = True
                        break
            if is_thumb:
                thumbs.append(f)
                continue
        imgs.append(f)
    return imgs, thumbs
MIME = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
        ".heic": "image/heic", ".heif": "image/heif", ".hif": "image/heic",
        ".html": "text/html; charset=utf-8", ".json": "application/json; charset=utf-8",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".txt": "text/plain; charset=utf-8", ".css": "text/css", ".js": "application/javascript"}

class H(BaseHTTPRequestHandler):
    server_version = "FilmIndex/1.0"

    def log_message(self, fmt, *a):
        pass

    def _send(self, code, data, ctype, cache="no-store"):
        if isinstance(data, str):
            data = data.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", cache)
        self.end_headers()
        try:
            self.wfile.write(data)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _json(self, o, code=200):
        self._send(code, json.dumps(o, ensure_ascii=False), "application/json; charset=utf-8")

    def do_GET(self):
        u = urlparse(self.path)
        path, q = unquote(u.path), parse_qs(u.query)
        try:
            if path == "/" or path.lower().endswith(".html"):
                with open(find_html(), "rb") as fh:
                    self._send(200, fh.read(), MIME[".html"])
            elif path == "/__icons":
                self._json(icon_report())
            elif path == "/__iconhd":
                f = (q.get("f") or [""])[0]
                p = safe_under(ICONHD_DIR, f) if f else None
                if p and os.path.isfile(p) and f.lower().endswith(".png"):
                    with open(p, "rb") as fh:
                        self._send(200, fh.read(), "image/png", "public, max-age=3600")
                else:
                    self._send(404, "not found", "text/plain")
            elif path == "/__icon":
                f = (q.get("f") or [""])[0]
                p = safe_under(ICON_DIR, f) if f else None
                if p and os.path.isfile(p) and f.lower().endswith(".png"):
                    with open(p, "rb") as fh:
                        self._send(200, fh.read(), "image/png", "public, max-age=3600")
                else:
                    self._send(404, "not found", "text/plain")
            elif path == "/__sync":
                bad = excel_read_error()
                if bad:
                    log("ERROR", "同步失败：%s" % bad["msg"])
                    self._json(bad, 500)
                else:
                    try:
                        snap = write_snapshot()
                    except Exception as e:
                        msg = "读取 Excel 内容失败：%s（%s）" % (type(e).__name__, e)
                        log("ERROR", "同步失败：%s" % msg)
                        self._json({"ok": False, "err": "read",
                                    "excel": os.path.basename(excel_path()), "msg": msg,
                                    "hint": "请确认该 xlsx 结构正常（含「胶卷信息总表」与「选项」两个工作表）"}, 500)
                        return
                    log("INFO", "快照已同步（%d 卷）" % snap["count"])
                    snap["ok"] = True
                    warns = []
                    w = excel_fallback_warn()
                    if w:
                        warns.append(w)
                    if snap["count"] == 0:
                        warns.append("读取成功，但没解析到任何有效条目"
                                     "（检查「序号」是否为三位或四位数字、「拍摄开始」是否已填）")
                    if warns:
                        snap["warn"] = "；".join(warns)
                        log("WARNING", "同步提醒：%s" % snap["warn"])
                    self._json(snap)
            elif path == "/__snapshot":
                with open(SNAP, encoding="utf-8") as fh:
                    self._json(json.load(fh))
            elif path == "/__settings":
                self._json(load_settings())
            elif path == "/__data":
                bad = excel_read_error()
                if bad:
                    log("ERROR", "读取表格数据失败：%s" % bad["msg"])
                    return self._json(bad, 500)
                rows, opts = read_grid()
                lock = excel_lock_file()
                log("INFO", "读取表格数据供编辑：%d 行 / 选项 %d·%d·%d%s"
                    % (len(rows), len(opts.get("cam") or []), len(opts.get("size") or []),
                       len(opts.get("model") or []), "（Excel 正打开）" if lock else ""))
                self._json({"ok": True, "rows": rows, "options": opts,
                            "modelIcons": _model_icons((opts or {}).get("model") or []),
                            "nextseq": next_seq(rows), "fields": COLS,
                            "locked": bool(lock), "excel": os.path.basename(excel_path()),
                            "lockfile": os.path.basename(lock) if lock else "",
                            "lastdesc": LAST.get("desc")})
            elif path == "/__lock":
                lock = excel_lock_file()
                xlsx = os.path.basename(excel_path())
                self._json({"ok": True, "locked": bool(lock), "excel": xlsx,
                            "lockfile": os.path.basename(lock) if lock else "",
                            "msg": ("Excel 正打开 %s，请先关闭 Excel 再增删改" % xlsx) if lock else ""})
            elif path == "/__options":
                _rows, opts = read_grid()
                self._json({"ok": True, "options": opts, "max": OPT_MAX,
                            "counts": {k: len(opts.get(k) or []) for k, _ in OPT_KEYS}})
            elif path == "/__template":
                with open(TPL, "rb") as fh:
                    data = fh.read()
                log("INFO", "下载 Excel 模板 template.xlsx（%.1f KB）" % (len(data) / 1024.0))
                self._send(200, data, MIME[".xlsx"])
            elif path == "/__xlsx":
                with open(excel_path(), "rb") as fh:
                    self._send(200, fh.read(), MIME[".xlsx"])
            elif path == "/__folderinfo":
                self._json(folder_info((q.get("seq") or [""])[0], (q.get("name") or [""])[0]))
            elif path == "/__list":
                d = (q.get("d") or [""])[0]
                p = safe_under(ROOT, d)
                if not p or not os.path.isdir(p):
                    return self._json({"ok": False, "msg": "目录不存在"})
                key = os.path.abspath(p)
                hit = _LIST_CACHE.get(key)
                if hit and time.time() - hit[0] < LIST_TTL:
                    return self._json({"ok": True, "dir": d, "files": hit[1], "thumbs": hit[2],
                                       "cached": True})
                fs, tfs = split_thumbs(sorted(f for f in os.listdir(p)
                                              if f.lower().endswith(IMG_EXT)))
                if len(_LIST_CACHE) > 800:
                    _LIST_CACHE.clear()
                _LIST_CACHE[key] = (time.time(), fs, tfs)
                self._json({"ok": True, "dir": d, "files": fs, "thumbs": tfs})
            elif path == "/__img":
                d = (q.get("d") or [""])[0]
                f = (q.get("f") or [""])[0]
                p = safe_under(ROOT, d + "/" + f) if f else None
                if not p or not os.path.isfile(p):
                    return self._send(404, "not found", "text/plain")
                ext = os.path.splitext(p)[1].lower()
                # 正在读原图：告知缩略图生成让路，把磁盘/连接让给大图
                img_enter()
                try:
                    with open(p, "rb") as fh:
                        data = fh.read()
                finally:
                    img_leave()
                self._send(200, data, MIME.get(ext, "application/octet-stream"),
                           "private, max-age=3600")
            elif path == "/__thumb":
                d = (q.get("d") or [""])[0]
                f = (q.get("f") or [""])[0]
                p = safe_under(ROOT, d + "/" + f) if f else None
                if not p or not os.path.isfile(p):
                    return self._send(404, "not found", "text/plain")
                data, ctype, cache = serve_thumb(p, q)
                self._send(200, data, ctype, cache)
            elif path == "/__thumbinfo":
                st = thumb_stats()
                st["ok"] = True
                st["levels"] = {str(k): list(v) for k, v in THUMB_LEVELS.items()}
                st["level"] = thumb_params()[1]
                st["on"] = thumb_params()[0]
                self._json(st)
            elif path == "/__ipinfo":
                self._json(ip_info())
            elif path == "/__ipopen":
                d = ip_out_dir()
                if not os.path.isdir(d):
                    try:
                        os.makedirs(d, exist_ok=True)
                    except OSError:
                        pass
                self._open_path(d)
            elif path == "/__open":
                self._open_path(safe_under(ROOT, (q.get("d") or [""])[0]))
            elif path == "/__openexcel":
                self._open_path(excel_path())
            elif path == "/__openlog":
                self._open_path(LOGDIR)
            elif path == "/__openbackup":
                self._open_path(BACKUP)
            elif path == "/__opendir":
                k = (q.get("k") or [""])[0]
                self._open_path({"log": LOGDIR, "backup": BACKUP, "assets": ASSETS,
                                 "thumb": THUMB}.get(k))
            else:
                self._send(404, "not found", "text/plain")
        except Exception as e:
            log("ERROR", "GET %s 失败：%s: %s" % (path, type(e).__name__, e))
            self._json({"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}, 500)

    def _open_path(self, p):
        if not p or not os.path.exists(p):
            return self._json({"ok": False, "msg": "路径不存在"})
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", p])
            elif os.name == "nt":
                os.startfile(p)          # noqa
            else:
                subprocess.Popen(["xdg-open", p])
            log("INFO", "已在访达中打开：%s" % os.path.basename(p))
            self._json({"ok": True})
        except Exception as e:
            log("ERROR", "打开路径失败 %s：%s" % (p, e))
            self._json({"ok": False, "msg": str(e)})

    def do_POST(self):
        path = unquote(urlparse(self.path).path)
        n = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(n) if n else b""
        try:
            if path == "/__settings":
                o = json.loads(body.decode("utf-8") or "{}")
                self._json({"ok": save_settings(o), "file": SETT})
            elif path == "/__options":
                o = json.loads(body.decode("utf-8") or "{}")
                self._json(op_save_options(o.get("options") or {}, bool(o.get("force"))))
            elif path == "/__clientdiag":
                o = json.loads(body.decode("utf-8") or "{}")
                msg = str(o.get("msg") or "")[:400].replace("\n", " ")
                log("WARNING", "[前端诊断] %s" % msg)
                self._json({"ok": True})
            elif path == "/__thumbclear":
                self._json(thumb_clear())
            elif path == "/__indexpaper":
                o = json.loads(body.decode("utf-8") or "{}")
                seq = str(o.get("seq") or "").strip()
                rec = None
                try:
                    with open(SNAP, encoding="utf-8") as fh:
                        for r in (json.load(fh).get("data") or []):
                            if str(r.get("seq")) == seq:
                                rec = r
                                break
                except Exception:
                    rec = None
                if rec is None:
                    return self._json({"ok": False, "seq": seq,
                                       "msg": "快照里没有序号 %s（可先点「⟳ 同步」再导出）" % seq})
                img_enter()          # 告诉缩略图队列：正在读大图，先让路
                try:
                    r = ip_build_docx(rec, o)
                finally:
                    img_leave()
                self._json(r, 200 if r.get("ok") else 500)
            elif path == "/__snapshot":
                o = json.loads(body.decode("utf-8") or "{}")
                with open(SNAP, "w", encoding="utf-8") as fh:
                    json.dump(o, fh, ensure_ascii=False)
                log("INFO", "快照已手动保存")
                self._json({"ok": True, "file": SNAP})
            elif path == "/__icons":
                o = json.loads(body.decode("utf-8") or "{}")
                mp = load_icon_map()
                act = str(o.get("action") or "set")
                if act == "clearall":
                    mp = {}
                    save_icon_map(mp)
                    log("INFO", "已清除全部手动图标映射（%d 条）" % len(load_icon_map()))
                    self._json({"ok": True, "cleared": True})
                else:
                    m = str(o.get("model") or "").strip()
                    f = str(o.get("file") or "").strip()
                    if not m:
                        self._json({"ok": False, "msg": "缺少型号"})
                    elif f and not os.path.isfile(os.path.join(ICON_DIR, f)):
                        self._json({"ok": False, "msg": "图标文件不存在：%s" % f})
                    else:
                        if f:
                            mp[m] = f
                            log("INFO", "手动指定图标：%s → %s" % (m, f))
                        else:
                            mp.pop(m, None)
                            log("INFO", "已取消手动指定：%s（回到自动匹配）" % m)
                        save_icon_map(mp)
                        try:
                            write_snapshot()
                        except Exception as e:
                            log("WARNING", "图标变更后刷新快照失败：%s" % e)
                        self._json({"ok": True})
            elif path == "/__mkdir":
                o = json.loads(body.decode("utf-8") or "{}")
                self._json(folder_mkdir(o.get("cells") or [], bool(o.get("subs"))))
            elif path == "/__syncfolder":
                o = json.loads(body.decode("utf-8") or "{}")
                cells = o.get("cells") or []
                seq = str(o.get("seq") or (cells[0] if cells else "") or "").strip()
                if cells:
                    self._json(folder_sync(seq, cells))
                else:
                    rows, _o = read_grid()
                    row = next((r for r in rows if r[0].strip() == seq), None)
                    self._json(folder_sync(seq, row) if row
                               else {"ok": False, "msg": "找不到序号 %s" % seq})
            elif path in ("/__add", "/__update", "/__delete", "/__undo"):
                o = json.loads(body.decode("utf-8") or "{}")
                self._json(self._crud(path, o))
            else:
                self._send(404, "not found", "text/plain")
        except json.JSONDecodeError as e:
            log("ERROR", "POST %s 请求体解析失败：%s" % (path, e))
            self._json({"ok": False, "msg": "请求格式错误"}, 400)
        except Exception as e:
            log("ERROR", "POST %s 失败：%s: %s" % (path, type(e).__name__, e))
            self._json({"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}, 500)

    def _crud(self, path, o):
        force = bool(o.get("force"))
        if path == "/__add":
            return op_add(o.get("cells") or [], force)
        if path == "/__update":
            return op_update(str(o.get("seq") or "").strip(), o.get("cells") or [], force)
        if path == "/__delete":
            return op_delete(str(o.get("seq") or "").strip(), force, bool(o.get("delFolder")))
        return op_undo(force)

def free_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    p = s.getsockname()[1]
    s.close()
    return p

def _launch(url):
    if os.environ.get("FILM_NO_OPEN"):
        return
    try:
        if sys.platform == "darwin":
            if subprocess.call(["open", "-a", "Safari", url]) != 0:
                subprocess.call(["open", url])
        elif os.name == "nt":
            os.startfile(url)          # noqa
        else:
            subprocess.call(["xdg-open", url])
    except Exception as e:
        log("WARNING", "自动打开浏览器失败：%s（请手动复制地址）" % e)

def main():
    if "--build" in sys.argv:
        ensure_excel()
        log("INFO", "手动重建快照")
        snap = write_snapshot()
        print("快照已更新：%d 卷" % snap["count"])
        return 0
    ensure_excel()
    log("INFO", "服务启动；版本 %s；Excel=%s" % (VERSION, os.path.basename(excel_path())))
    if thumb_backend() == "powershell":
        threading.Thread(target=_engine_start, name="thumb-engine-warm", daemon=True).start()
    ensure_thumb_dir()
    log("INFO", "预览缓存目录：%s" % THUMB)
    try:
        migrate_thumb_dir()
    except Exception as e:
        log("WARNING", "预览缓存目录初始化异常：%s" % e)
    try:
        backup_on_start()
    except Exception as e:
        log("ERROR", "启动备份失败：%s" % e)
    try:
        snap = write_snapshot()
        log("INFO", "启动快照完成（%d 卷）" % snap["count"])
    except Exception as e:
        log("ERROR", "启动快照失败：%s" % e)
    port = free_port()
    url = "http://127.0.0.1:%d/" % port
    srv = ThreadingHTTPServer(("127.0.0.1", port), H)
    print("=" * 62)
    print("  胶卷索引本地服务已启动（v%s）" % VERSION)
    print("  请用 Safari 打开这个地址（已尝试自动打开）：")
    print("  %s" % url)
    print("  用完关闭这个终端窗口即停止服务。")
    print("=" * 62)
    threading.Timer(1.0, lambda: _launch(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        srv.server_close()
        log("INFO", "服务已停止")
    return 0

if __name__ == "__main__":
    sys.exit(main())
