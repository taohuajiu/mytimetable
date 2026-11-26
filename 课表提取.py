#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
金山文档课表提取器（公开版 · 不含单价）
========================================
用法与完整版完全一样，区别只有两点：

1. 一对一学生单价不写在代码里。优先读环境变量 STUDENT_PRICES（GitHub Actions 用 Secret 注入），
   其次读脚本同目录的 prices.json，都没有就按兜底价算并给出警告。
   格式： {"学生A": 500, "学生B": 620}（键是学生名，值是课单价）
2. 输出的 timetable.json 只给"节数 + 金额合计"：
   每节课不带 price、不带 student，也不输出按学生的 breakdown，避免把学生单价公开。

默认行为
--------
1. 取数据源：优先用「课表源文件」里的本地表；没有则尝试从金山文档下载；
2. 在【所有工作表】里找出"姓名/老师"列等于目标姓名的行（支持合并单元格自动向下填充）；
3. 只保留未来 N 天（默认含今天）的行；
4. 单元格字体、底色、边框、对齐、数字格式、合并单元格、列宽、行高全部从源表复制，
   即"按原来的格式"输出；
5. 同一个文件里再附一个「本月课时与课酬」工作表；
6. 结果保存到桌面，一个文件包含全部内容；
7. 如果指定了 --json-out，或运行在 GitHub Actions 中，额外输出 timetable.json。

准备（二选一）
--------------
A. 本地文件：把课表从金山文档下载/另存为 xlsx，放进本脚本同目录的「课表源文件」文件夹
   （也可以用 --file 直接指定路径）。
B. 自动下载：在本脚本同目录新建 cookie.txt，或设置环境变量 WPS_COOKIE。
   GitHub Actions 中务必使用 WPS_COOKIE Secret。

用法
----
    python 课表提取.py
    python 课表提取.py --name 段锋杰 --days 7
    python 课表提取.py --refresh --json-out timetable.json
    python 课表提取.py --stats --month 2026-12

依赖：pip install openpyxl
"""

import argparse
import datetime as dt
import json
import os
import re
import sys
import urllib.error
import urllib.request
from copy import copy
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("缺少依赖 openpyxl，请先执行： pip install openpyxl")


# ============================== 配置区 ==============================
KDOCS_URL = "https://www.kdocs.cn/l/co6ZLHXYhHuL"
KDOCS_FILE_ID = "mz1ofYQ1q1MXuz9B8oMz1xsgDoo8HT5vY"

TARGET_NAME = "段锋杰"
DAYS = 7
INCLUDE_TODAY = True

NAME_HEADER_HINTS = ("姓名", "老师", "教师", "任课老师", "授课老师")
DATE_HEADER_HINTS = ("日期", "上课日期", "时间")

GROUP_HINTS = ("班", "年级", "高复", "艺考")
GRADE_PATTERN = re.compile(r"高[一二三]|初[一二三]|小学|大班|小班")
SPLIT_PATTERN = re.compile(r"[、,，/;；\n]+")

# ---------- 课时费规则（按学生精确计算）----------
COMMISSION = 0.35          # 一对一：老师到手 = 学生课单价 × 0.35
BASE_SALARY = 1600.0       # 底薪（每月固定，计入"当前到手/本月到手"）

RATE_BAND = 180.0          # 班课：到手固定 180 元/节
RATE_TRIAL = 160.0         # 体验课：到手固定 160 元/节

# 一对一学生的课单价（按报名/标价，到手需乘 COMMISSION）——本文件不保存，运行时注入
# 来源优先级：环境变量 STUDENT_PRICES > 脚本同目录 prices.json > 空（用兜底价）
STUDENT_PRICE = {}

# 课表里的其他写法 → 统一到上面的人名（同一人同价）
NAME_ALIAS = {
    "黄锴": "黄楷",
}

# 不按提成、直接带固定到手价的学生（唐莎：到手 180 元/节）
FIXED_STUDENT_RATE = {
    "唐莎": 180.0,
}

DEFAULT_ONE_RATE = 210.0   # 名单外的学生：沿用原来的一对一 210 元/节兜底


def load_student_prices():
    """从环境变量或 prices.json 读一对一单价；读不到就返回空表。"""
    raw = os.environ.get("STUDENT_PRICES", "").strip()
    if raw:
        try:
            return {str(k): float(v) for k, v in json.loads(raw).items()}
        except Exception as e:
            sys.exit("STUDENT_PRICES 解析失败（应为 {{\"姓名\": 单价}}）：{}".format(e))

    p = Path(__file__).resolve().parent / "prices.json"
    if p.exists():
        try:
            return {str(k): float(v) for k, v in
                    json.loads(p.read_text(encoding="utf-8")).items()}
        except Exception as e:
            sys.exit("prices.json 解析失败：{}".format(e))

    print("！没有找到一对一单价（STUDENT_PRICES 或 prices.json），"
          "名单内学生会按 {} 元/节兜底，金额偏低。".format(DEFAULT_ONE_RATE))
    return {}


STUDENT_PRICE = load_student_prices()
UNKNOWN_LABEL = "未登记学生"

KINDS = ("班课", "一对一", "体验课")
TRIAL_PATTERN = re.compile(r"(?:[Tt]|体验课?)$")
# ===================================================================

SCRIPT_DIR = Path(__file__).resolve().parent
# GitHub Actions 的日志对公开仓库任何人可见，跑在云端时只输出总数
QUIET = os.environ.get("GITHUB_ACTIONS") == "true"
SOURCE_DIR = SCRIPT_DIR / "课表源文件"


# ------------------------------ 小工具 ------------------------------
def clean_text(value):
    if value is None:
        return ""
    if isinstance(value, str):
        s = value.replace("\r", " ").replace("\n", " ")
        s = re.sub(r"<br\s*/?>", " ", s, flags=re.I)
        s = re.sub(r"[\s\u3000\u00a0]+", "", s)
        return s
    return str(value).strip()


def resolve_desktop():
    home = Path.home()
    for cand in (home / "Desktop", home / "OneDrive" / "Desktop", home / "桌面"):
        if cand.is_dir():
            return cand
    return home


def to_date(value, today):
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)) and 20000 < value < 80000:
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(value))

    s = clean_text(value)
    if not s:
        return None

    def build(y, m, d):
        try:
            return dt.date(y, m, d)
        except ValueError:
            return None

    m = re.search(r"(\d{2,4})\s*[年/\-.]\s*(\d{1,2})\s*[月/\-.]\s*(\d{1,2})", s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        return build(y, mo, d)

    m = re.search(r"(\d{1,2})\s*[月/\-.]\s*(\d{1,2})", s)
    if m:
        mo, d = int(m.group(1)), int(m.group(2))
        cands = [build(y, mo, d) for y in (today.year - 1, today.year, today.year + 1)]
        cands = [c for c in cands if c]
        if cands:
            return min(cands, key=lambda c: abs((c - today).days))
    return None


# ---------------------------- 取数据源 ----------------------------
def download_from_kdocs(file_id, cookie, dest):
    api = ("https://www.kdocs.cn/api/v3/office/file/{}/download"
           "?isblocks=false&support_checksums=1").format(file_id)
    headers = {
        "Cookie": cookie,
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/124.0 Safari/537.36"),
        "Referer": KDOCS_URL,
        "Accept": "*/*",
    }

    def fetch(url):
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=90) as resp:
            return resp.headers.get("Content-Type", ""), resp.read()

    try:
        ctype, data = fetch(api)
    except urllib.error.HTTPError as e:
        body = e.read()[:300].decode("utf-8", "ignore")
        if e.code in (401, 403):
            raise RuntimeError(
                "金山文档返回未登录/无权限，说明 cookie 已过期或没有该文档权限。\n"
                "请重新从浏览器复制 Cookie，或改用 --file 指定本地已下载的 xlsx。\n"
                "服务端返回：{}".format(body))
        raise RuntimeError("下载失败 HTTP {}：{}".format(e.code, body))

    if not data[:2] == b"PK":
        try:
            info = json.loads(data.decode("utf-8", "ignore"))
        except Exception:
            info = {}
        inner = info.get("url") or (info.get("data") or {}).get("url")
        if not inner:
            raise RuntimeError("下载接口没有返回文件流：" + data[:300].decode("utf-8", "ignore"))
        ctype, data = fetch(inner)

    dest.write_bytes(data)
    print("已从金山文档下载：{}（{} 字节）".format(dest, len(data)))
    return dest


def pick_local_source(explicit):
    if explicit:
        p = Path(explicit).expanduser()
        if not p.is_absolute():
            p = SCRIPT_DIR / p
        if not p.exists():
            sys.exit("找不到文件：{}".format(p))
        return p

    cache = SOURCE_DIR / "26年度上学期课表.xlsx"
    if cache.exists():
        return cache
    if SOURCE_DIR.is_dir():
        xlsx = [p for p in SOURCE_DIR.glob("*.xlsx") if not p.name.startswith("~$")]
        if xlsx:
            return max(xlsx, key=lambda p: p.stat().st_mtime)
    return None


def read_cookie():
    # 优先环境变量（GitHub Actions Secrets）
    env_cookie = os.environ.get("WPS_COOKIE")
    if env_cookie and env_cookie.strip():
        return env_cookie.strip()

    cookie_file = SCRIPT_DIR / "cookie.txt"
    if not cookie_file.exists():
        return None
    cookie = cookie_file.read_text(encoding="utf-8").strip()
    cookie = re.sub(r"^cookie\s*:\s*", "", cookie, flags=re.I).strip()
    return cookie or None


def get_source(explicit_file=None, refresh=False):
    if explicit_file:
        src = pick_local_source(explicit_file)
        print("数据源：本地文件 {}".format(src))
        return src

    cache = SOURCE_DIR / "26年度上学期课表.xlsx"
    cookie = read_cookie()

    if refresh:
        if not cookie:
            print("！--refresh 需要 WPS_COOKIE 或 cookie.txt，但没有找到，改用本地已下载的表")
        else:
            SOURCE_DIR.mkdir(parents=True, exist_ok=True)
            print("数据源：金山文档在线下载（--refresh）")
            try:
                return download_from_kdocs(KDOCS_FILE_ID, cookie, cache)
            except Exception as e:
                print("！自动下载失败：{}".format(e))
                if not cache.exists():
                    sys.exit("自动下载失败，本地也没有缓存表，请手动放一份到 {}".format(SOURCE_DIR))
                print("  改用本地缓存表继续，日期口径仍按今天算")

    src = pick_local_source(None)
    if src:
        print("数据源：本地文件 {}".format(src))
        if cookie:
            print("  （想让脚本先拉最新表，加 --refresh）")
        return src

    if cookie:
        SOURCE_DIR.mkdir(parents=True, exist_ok=True)
        print("数据源：金山文档在线下载")
        return download_from_kdocs(KDOCS_FILE_ID, cookie, cache)

    sys.exit(
        "没有找到数据源，也没有 WPS_COOKIE / cookie.txt。\n"
        "请任选一种：\n"
        "  1) 浏览器打开 {} ，下载/另存为 xlsx，放进 {} ；\n"
        "  2) 设置环境变量 WPS_COOKIE，或在本目录建 cookie.txt，让脚本自动下载。"
        .format(KDOCS_URL, SOURCE_DIR))


# ---------------------------- 定位表头 ----------------------------
def locate_header(ws, max_scan=5):
    name_col = date_col = None
    header_row = 1
    for r in range(1, min(max_scan, ws.max_row or 1) + 1):
        for c in range(1, (ws.max_column or 1) + 1):
            txt = clean_text(ws.cell(row=r, column=c).value)
            if not txt:
                continue
            if name_col is None and len(txt) <= 6 and any(h in txt for h in NAME_HEADER_HINTS):
                name_col, header_row = c, r
            if date_col is None and len(txt) <= 6 and any(h in txt for h in DATE_HEADER_HINTS):
                date_col = c
        if name_col:
            break
    if name_col is None:
        name_col, header_row = 1, 1
    if date_col is None:
        date_col = name_col + 1
    return header_row, name_col, date_col


def match_name(cell_value, target):
    s = clean_text(cell_value)
    if not s:
        return False
    return s == clean_text(target) or clean_text(target) in s


def extract_rows(ws, target, today, date_from=None, date_to=None):
    header_row, name_col, date_col = locate_header(ws)
    rows = []
    current_name = ""
    for r in range(header_row + 1, (ws.max_row or 0) + 1):
        filled = clean_text(ws.cell(row=r, column=name_col).value)
        if filled:
            current_name = filled
        if not current_name or not match_name(current_name, target):
            continue
        d = to_date(ws.cell(row=r, column=date_col).value, today)
        if date_from or date_to:
            if d is None:
                continue
            if date_from and d < date_from:
                continue
            if date_to and d > date_to:
                continue
        rows.append((r, d))
    return header_row, name_col, date_col, rows


# -------------------------- 按原格式复制 --------------------------
def copy_cell_style(src, dst):
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def build_sheet(src_ws, dst_ws, keep_rows, header_row):
    last_col = src_ws.max_column or 1
    src_rows = [header_row] + [r for r in keep_rows if r != header_row]

    for key, dim in src_ws.column_dimensions.items():
        if dim.width:
            dst_ws.column_dimensions[key].width = dim.width
        if dim.hidden:
            dst_ws.column_dimensions[key].hidden = True
    for key, dim in src_ws.row_dimensions.items():
        if dim.height:
            dst_ws.row_dimensions[key].height = dim.height
    try:
        if src_ws.sheet_format.defaultRowHeight:
            dst_ws.sheet_format.defaultRowHeight = src_ws.sheet_format.defaultRowHeight
        if src_ws.sheet_format.defaultColWidth:
            dst_ws.sheet_format.defaultColWidth = src_ws.sheet_format.defaultColWidth
    except Exception:
        pass

    merges = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in src_ws.merged_cells.ranges]

    def owner(r, c):
        for r1, c1, r2, c2 in merges:
            if r1 <= r <= r2 and c1 <= c <= c2:
                return (r1, c1, r2, c2)
        return None

    dst_of = {}
    for i, sr in enumerate(src_rows):
        dr = i + 1
        dst_of[sr] = dr
        for c in range(1, last_col + 1):
            ow = owner(sr, c)
            if ow and (sr, c) != (ow[0], ow[1]):
                continue
            copy_cell_style(src_ws.cell(row=sr, column=c), dst_ws.cell(row=dr, column=c))

    kept = set(src_rows)
    for r1, c1, r2, c2 in merges:
        hit = [r for r in src_rows if r1 <= r <= r2]
        if not hit:
            continue
        if hit[-1] - hit[0] + 1 != len(hit):
            continue
        dr1, dr2 = dst_of[hit[0]], dst_of[hit[-1]]
        if dr1 != dr2 or c1 != c2:
            try:
                dst_ws.merge_cells(start_row=dr1, start_column=c1, end_row=dr2, end_column=c2)
            except Exception:
                pass
        if r1 not in kept:
            target = dst_ws.cell(row=dr1, column=c1)
            if target.value in (None, ""):
                copy_cell_style(src_ws.cell(row=r1, column=c1), target)
                first_kept = src_ws.cell(row=hit[0], column=c1)
                try:
                    if first_kept.fill is not None and first_kept.fill.fill_type == "solid":
                        target.fill = copy(first_kept.fill)
                except Exception:
                    pass

    try:
        dst_ws.freeze_panes = src_ws.freeze_panes
    except Exception:
        pass
    try:
        dst_ws.print_title_rows = "1:1"
        dst_ws.page_setup.orientation = src_ws.page_setup.orientation or "landscape"
        dst_ws.page_setup.fitToWidth = 1
        dst_ws.page_setup.fitToHeight = 0
        dst_ws.sheet_properties.pageSetUpPr.fitToPage = True
    except Exception:
        pass


# -------------------------- 课时统计 --------------------------
WEEKDAY_CN = ("星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日")
TIME_HEADER_PATTERN = re.compile(r"\d{1,2}\s*[:：]\s*\d{2}")


def is_group_course(text):
    return any(h in text for h in GROUP_HINTS) or bool(GRADE_PATTERN.search(text))


def classify_course(text):
    s = clean_text(text)
    if not s:
        return None
    if is_group_course(s):
        return "班课"
    if TRIAL_PATTERN.search(s):
        return "体验课"
    return "一对一"


def strip_trial(text):
    """去掉体验课标记（末尾的 T / 体验课），只留学生名+备注。"""
    s = clean_text(text)
    s = re.sub(r"(?:体验课|体验)$", "", s)
    s = re.sub(r"[Tt]$", "", s)
    return s


def student_of(text):
    """从课程文字里认出学生姓名，认不出返回 None。"""
    s = strip_trial(text)
    if not s:
        return None
    keys = list(STUDENT_PRICE) + list(FIXED_STUDENT_RATE) + list(NAME_ALIAS)
    for name in sorted(keys, key=len, reverse=True):
        if name in s:
            return NAME_ALIAS.get(name, name)
    return None


def one_rate(name):
    """一对一老师到手单价。"""
    if name in FIXED_STUDENT_RATE:
        return FIXED_STUDENT_RATE[name]
    if name in STUDENT_PRICE:
        return STUDENT_PRICE[name] * COMMISSION
    return DEFAULT_ONE_RATE


def item_money(kind, part):
    """单节课老师到手金额。"""
    if kind == "班课":
        return RATE_BAND
    if kind == "体验课":
        return RATE_TRIAL
    return one_rate(student_of(part))


def category_rates():
    """所有可能的收费条目 → 到手单价。"""
    rates = {"班课": RATE_BAND, "体验课": RATE_TRIAL, UNKNOWN_LABEL: DEFAULT_ONE_RATE}
    for n, p in STUDENT_PRICE.items():
        rates[n] = p * COMMISSION
    rates.update(FIXED_STUDENT_RATE)
    return rates


def category_of(kind, part):
    """把一节课归到"班课/体验课/某学生"这一栏。"""
    if kind == "班课":
        return "班课"
    if kind == "体验课":
        return "体验课"
    return student_of(part) or UNKNOWN_LABEL


def summarize_by_category(detail, day_limit):
    """按"班课 / 体验课 / 每个学生"汇总全月与已上的节数和金额。"""
    rates = category_rates()
    rows = {}
    for d in sorted(detail):
        done = d <= day_limit
        b = detail[d]
        for kind in KINDS:
            for _, _, part in b[kind]:
                key = category_of(kind, part)
                r = rows.setdefault(key, {
                    "label": key, "rate": rates.get(key),
                    "all_n": 0, "all_m": 0.0, "done_n": 0, "done_m": 0.0,
                })
                m = item_money(kind, part)
                r["all_n"] += 1
                r["all_m"] += m
                if done:
                    r["done_n"] += 1
                    r["done_m"] += m

    ordered = [rows.pop(k) for k in ("班课", "体验课") if k in rows]
    ordered += sorted(rows.values(), key=lambda x: (-x["all_m"], x["label"]))
    return ordered


def round2(v):
    """四舍五入到分（229.075 → 229.08），避免浮点把 .5 舍掉。"""
    if v is None:
        return None
    return float(Decimal(str(round(v, 6))).quantize(Decimal("0.01"),
                                                    rounding=ROUND_HALF_UP))


def fmt_money(v):
    return "{:,.2f}".format(round2(v))


def collect_month(src_wb, name, year, month, today):
    detail = {}
    sheet_titles = []
    for ws in src_wb.worksheets:
        header_row, name_col, date_col, rows = extract_rows(ws, name, today)
        if not rows:
            continue
        time_cols = []
        for c in range(1, (ws.max_column or 1) + 1):
            h = clean_text(ws.cell(row=header_row, column=c).value)
            if TIME_HEADER_PATTERN.search(h):
                time_cols.append((c, h))
        if not time_cols:
            continue

        hit = False
        for r, d in rows:
            if d is None or d.year != year or d.month != month:
                continue
            for c, h in time_cols:
                raw = ws.cell(row=r, column=c).value
                for part in SPLIT_PATTERN.split(clean_text(raw) or ""):
                    kind = classify_course(part)
                    if not kind:
                        continue
                    bucket = detail.setdefault(d, {k: [] for k in KINDS})
                    bucket[kind].append((c, h, part))
                    hit = True
        if hit:
            sheet_titles.append(ws.title)

    summary = {k: sum(len(v[k]) for v in detail.values()) for k in KINDS}
    return detail, summary, sheet_titles


def collect_range(src_wb, name, start, end, today):
    """收集 start~end 范围内的课程，结构同 collect_month 的 detail。"""
    detail = {}
    for ws in src_wb.worksheets:
        header_row, name_col, date_col, rows = extract_rows(ws, name, today)
        if not rows:
            continue
        time_cols = []
        for c in range(1, (ws.max_column or 1) + 1):
            h = clean_text(ws.cell(row=header_row, column=c).value)
            if TIME_HEADER_PATTERN.search(h):
                time_cols.append((c, h))
        if not time_cols:
            continue
        for r, d in rows:
            if d is None or d < start or d > end:
                continue
            for c, h in time_cols:
                raw = ws.cell(row=r, column=c).value
                for part in SPLIT_PATTERN.split(clean_text(raw) or ""):
                    kind = classify_course(part)
                    if not kind:
                        continue
                    bucket = detail.setdefault(d, {k: [] for k in KINDS})
                    bucket[kind].append((c, h, part))
    return detail


def day_courses(bucket):
    items = [x for k in KINDS for x in bucket[k]]
    return [(h, s) for _, h, s in sorted(items, key=lambda x: x[0])]


def bucket_money(bucket):
    return sum(item_money(k, part) for k in KINDS for _, _, part in bucket[k])


def disp_width(text):
    import unicodedata
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in text)


def pad(text, width):
    return text + " " * max(0, width - disp_width(text))


def split_month(detail, today):
    done_days = [d for d in detail if d <= today]
    all_sum = {k: sum(len(v[k]) for v in detail.values()) for k in KINDS}
    done_sum = {k: sum(len(detail[d][k]) for d in done_days) for k in KINDS}
    return all_sum, done_sum


def print_stats(name, year, month, detail, summary, today=None):
    today = today or dt.date.max
    rows = summarize_by_category(detail, today)
    all_sum, done_sum = split_month(detail, today)
    total_all, total_done = sum(all_sum.values()), sum(done_sum.values())
    money_all = sum(r["all_m"] for r in rows)
    money_done = sum(r["done_m"] for r in rows)

    if QUIET:
        print("\n{} · {}年{}月：全月 {} 节，课酬 {} 元；已上 {} 节，课酬 {} 元；"
              "加底薪 {} 元后到手 {:.2f} / {:.2f} 元（明细不在命令行输出）".format(
                  name, year, month, total_all, fmt_money(money_all),
                  total_done, fmt_money(money_done), fmt_money(BASE_SALARY),
                  money_done + BASE_SALARY, money_all + BASE_SALARY))
        return

    weekdays = sum(1 for d in detail if d.weekday() < 5)
    pending = sum(1 for d in detail if d > today)

    widths = (14, 12, 10, 14, 10, 14)
    print("\n{} · {}年{}月课时与课酬".format(name, year, month))
    print("-" * 78)
    print("  " + "".join(pad(h, w) for h, w in zip(
        ("类型/学生", "到手单价", "全月节数", "全月课酬", "已上节数", "已上课酬"), widths)))
    for r in rows:
        print("  " + pad(r["label"], widths[0])
              + fmt_money(r["rate"]).rjust(widths[1])
              + str(r["all_n"]).rjust(widths[2])
              + fmt_money(r["all_m"]).rjust(widths[3])
              + str(r["done_n"]).rjust(widths[4])
              + fmt_money(r["done_m"]).rjust(widths[5]))
    print("  " + "-" * 76)
    print("  " + pad("课酬合计", widths[0])
          + "".rjust(widths[1])
          + str(total_all).rjust(widths[2])
          + fmt_money(money_all).rjust(widths[3])
          + str(total_done).rjust(widths[4])
          + fmt_money(money_done).rjust(widths[5]))
    print("  " + pad("底薪", widths[0])
          + "".rjust(widths[1]) + "".rjust(widths[2])
          + fmt_money(BASE_SALARY).rjust(widths[3])
          + "".rjust(widths[4])
          + fmt_money(BASE_SALARY).rjust(widths[5]))
    print("  " + pad("到手总计", widths[0])
          + "".rjust(widths[1])
          + str(total_all).rjust(widths[2])
          + fmt_money(money_all + BASE_SALARY).rjust(widths[3])
          + str(total_done).rjust(widths[4])
          + fmt_money(money_done + BASE_SALARY).rjust(widths[5]))
    print("  （有课 {} 天，工作日 {} 天、周末 {} 天；今天 {}，还有 {} 天待上）"
          .format(len(detail), weekdays, len(detail) - weekdays,
                  today if today != dt.date.max else "-", pending))
    print("-" * 78)

    running = 0.0
    for d in sorted(detail):
        b = detail[d]
        cells = ["{} {}".format(h, s) for h, s in day_courses(b)]
        n = len(b["班课"]) + len(b["一对一"]) + len(b["体验课"])
        m = bucket_money(b)
        if d <= today:
            running += m
            tail = "累计{:>10}".format(fmt_money(running))
        else:
            tail = "    未上    "
        print("  {:>2}/{:02d} {}  {}节 {:>9}元 {}  {}".format(
            d.month, d.day, WEEKDAY_CN[d.weekday()], n, fmt_money(m), tail,
            "；".join(cells)))
    print("-" * 78)


def write_month_sheet(wb, name, year, month, detail, summary, today=None, sheet_title=None):
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    day_limit = today or dt.date.max
    thin = Side(style="thin", color="FFBFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="FF92D050")
    sum_fill = PatternFill("solid", fgColor="FFE2F0D9")
    today_fill = PatternFill("solid", fgColor="FFFFF2CC")
    center = Alignment(horizontal="center", vertical="center")
    grey = Font(name="微软雅黑", size=11, color="FFA6A6A6", italic=True)

    rows = summarize_by_category(detail, day_limit)
    all_sum, done_sum = split_month(detail, day_limit)
    money_all = sum(r["all_m"] for r in rows)
    money_done = sum(r["done_m"] for r in rows)
    total_all, total_done = sum(all_sum.values()), sum(done_sum.values())
    pending = sum(1 for d in detail if d > day_limit)

    ws = wb.create_sheet((sheet_title or "课时与课酬")[:31])

    ws["A1"] = "{} · {}年{}月课时与课酬".format(name, year, month)
    ws["A1"].font = Font(name="微软雅黑", size=14, bold=True)
    ws.merge_cells("A1:I1")
    ws.row_dimensions[1].height = 26

    if today:
        head = ("截至 {}月{}日（含今天）已上 {} 节，课酬 {} 元；加底薪 {} 元，"
                "到手合计 {} 元。本月已排 {} 节，课酬 {} 元，到手合计 {} 元，"
                "还有 {} 天课未上。").format(
            today.month, today.day, total_done, fmt_money(money_done),
            fmt_money(BASE_SALARY), fmt_money(money_done + BASE_SALARY),
            total_all, fmt_money(money_all), fmt_money(money_all + BASE_SALARY), pending)
    else:
        head = "本月已排 {} 节，课酬 {} 元，到手合计 {} 元。".format(
            total_all, fmt_money(money_all), fmt_money(money_all + BASE_SALARY))
    c = ws.cell(row=2, column=1, value=head)
    c.font = Font(name="微软雅黑", size=11, bold=True, color="FFC00000")
    ws.merge_cells("A2:I2")

    headers = ("类型/学生", "到手单价(元/节)", "全月节数", "全月小计(元)",
               "已上节数", "已上小计(元)")
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=i, value=h)
        cell.font = Font(name="微软雅黑", size=11, bold=True)
        cell.fill = head_fill
        cell.alignment = center
        cell.border = border

    r = 5
    for item in rows:
        vals = (item["label"],
                None if item["rate"] is None else round2(item["rate"]),
                item["all_n"], round2(item["all_m"]),
                item["done_n"], round2(item["done_m"]))
        for i, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=i, value=v)
            cell.font = Font(name="微软雅黑", size=11)
            cell.alignment = center
            cell.border = border
        ws.cell(row=r, column=2).number_format = "#,##0.00"
        ws.cell(row=r, column=4).number_format = "#,##0.00"
        ws.cell(row=r, column=6).number_format = "#,##0.00"
        r += 1

    def put_sum_row(label, n1, m1, n2, m2, bold_money=False):
        nonlocal r
        for i, v in enumerate((label, None, n1, m1, n2, m2), 1):
            cell = ws.cell(row=r, column=i, value=v)
            cell.font = Font(name="微软雅黑", size=11, bold=True)
            cell.fill = sum_fill
            cell.alignment = center
            cell.border = border
        ws.cell(row=r, column=4).number_format = "#,##0.00"
        ws.cell(row=r, column=6).number_format = "#,##0.00"
        if bold_money:
            ws.cell(row=r, column=6).font = Font(
                name="微软雅黑", size=12, bold=True, color="FFC00000")
        r += 1

    put_sum_row("课酬合计", total_all, round2(money_all), total_done, round2(money_done))
    put_sum_row("底薪", None, round2(BASE_SALARY), None, round2(BASE_SALARY))
    put_sum_row("到手总计", total_all, round2(money_all + BASE_SALARY),
                total_done, round2(money_done + BASE_SALARY), bold_money=True)
    total_row = r - 1

    r = total_row + 2
    c = ws.cell(row=r, column=1, value="逐日明细（累计列只算到今天，灰字为未上的课）")
    c.font = Font(name="微软雅黑", size=12, bold=True)
    r += 1
    headers = ("日期", "星期", "班课", "一对一", "体验课", "节数", "课酬(元)", "累计(元)", "课程")
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.font = Font(name="微软雅黑", size=11, bold=True)
        cell.fill = head_fill
        cell.alignment = center
        cell.border = border
    head_row = r
    r += 1

    running = 0
    for d in sorted(detail):
        b = detail[d]
        courses = ["{} {}".format(h, s) for h, s in day_courses(b)]
        n = len(b["班课"]) + len(b["一对一"]) + len(b["体验课"])
        m = bucket_money(b)
        if d <= day_limit:
            running += m
            cum = running
        else:
            cum = None
        row_vals = ("{:d}月{:d}号".format(d.month, d.day), WEEKDAY_CN[d.weekday()],
                    len(b["班课"]), len(b["一对一"]), len(b["体验课"]), n, m, cum,
                    "；".join(courses))
        for i, v in enumerate(row_vals, 1):
            cell = ws.cell(row=r, column=i, value=v)
            cell.font = Font(name="微软雅黑", size=11)
            cell.border = border
            cell.alignment = center if i <= 8 else Alignment(
                horizontal="left", vertical="center", wrap_text=True)
            if d == day_limit:
                cell.fill = today_fill
            elif d > day_limit:
                cell.font = grey
        ws.cell(row=r, column=7).number_format = "#,##0.00"
        ws.cell(row=r, column=8).number_format = "#,##0.00"
        lines = max(1, -(-disp_width(row_vals[-1]) // 44))
        if lines > 1:
            ws.row_dimensions[r].height = 16 * lines
        r += 1

    for i, v in enumerate(("合计", None, all_sum["班课"], all_sum["一对一"],
                           all_sum["体验课"], total_all, round2(money_all),
                           round2(money_done), None), 1):
        cell = ws.cell(row=r, column=i, value=v)
        cell.font = Font(name="微软雅黑", size=11, bold=True)
        cell.fill = sum_fill
        cell.border = border
        cell.alignment = center
    ws.cell(row=r, column=7).number_format = "#,##0.00"
    ws.cell(row=r, column=8).number_format = "#,##0.00"
    cell = ws.cell(row=r, column=9, value="课酬合计＝全月；累计列＝截至今天（不含底薪）")
    cell.font = Font(name="微软雅黑", size=9, color="FF808080")
    cell.alignment = Alignment(horizontal="left", vertical="center")

    r += 2
    price_txt = "、".join("{} {}".format(n, fmt_money(p * COMMISSION))
                         for n, p in STUDENT_PRICE.items())
    note = ("说明：班课 {:.0f} 元/节；体验课（名字带 T）{:.0f} 元/节；"
            "一对一按「学生课单价 × 提成 {:.0%}」算到手（{}），唐莎固定 {:.0f} 元/节；"
            "底薪 {:.0f} 元/月。金额按精确值累加，显示保留两位小数。").format(
        RATE_BAND, RATE_TRIAL, COMMISSION, price_txt,
        FIXED_STUDENT_RATE.get("唐莎", RATE_BAND), BASE_SALARY)
    cell = ws.cell(row=r, column=1, value=note)
    cell.font = Font(name="微软雅黑", size=9, color="FF808080")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)

    for col, width in zip("ABCDEFGHI", (12, 12, 9, 9, 9, 9, 11, 11, 44)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = ws.cell(row=head_row + 1, column=1)
    return ws


# ------------------------------ JSON 输出 ------------------------------
def output_json(name, today, month_detail, month_summary, range_detail,
                month_year, month_month, json_path):
    """输出 timetable.json 供手机端 Scriptable 读取。"""
    all_sum, done_sum = split_month(month_detail, today)
    rows = summarize_by_category(month_detail, today)
    money_all = sum(r["all_m"] for r in rows)
    money_done = sum(r["done_m"] for r in rows)

    # 兼容手机端旧脚本：band/one/trial 仍然给 count/money/rate 三个键
    def kind_money(labels):
        return sum(r["all_m"] for r in rows if r["label"] in labels)

    band_labels = {"班课"}
    trial_labels = {"体验课"}
    one_labels = {r["label"] for r in rows} - band_labels - trial_labels
    one_money = kind_money(one_labels)
    one_count = all_sum["一对一"]

    data = {
        "name": name,
        "updated": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "month_total": round2(money_all),
        "month": {
            "year": month_year,
            "month": month_month,
            "total_hours": sum(all_sum.values()),
            "total_money": round2(money_all),
            "done_hours": sum(done_sum.values()),
            "done_money": round2(money_done),
            "base_salary": round2(BASE_SALARY),
            "total_income": round2(money_all + BASE_SALARY),
            "done_income": round2(money_done + BASE_SALARY),
            "band": {"count": all_sum["班课"], "money": round2(kind_money(band_labels)),
                     "rate": RATE_BAND},
            "one": {"count": one_count, "money": round2(one_money),
                    "rate": round2(one_money / one_count) if one_count else DEFAULT_ONE_RATE},
            "trial": {"count": all_sum["体验课"], "money": round2(kind_money(trial_labels)),
                      "rate": RATE_TRIAL},
            # 公开版：不输出按学生的 breakdown，避免暴露学生单价
        },
        "courses": [],
    }

    for d in sorted(range_detail):
        b = range_detail[d]
        items = []
        for kind in KINDS:
            for c, h, part in b[kind]:
                items.append((c, h, part, kind))
        items.sort(key=lambda x: x[0])
        for c, h, part, kind in items:
            data["courses"].append({
                "date": d.isoformat(),
                "weekday": WEEKDAY_CN[d.weekday()],
                "time": h,
                "course": part,
                "kind": kind,
                # 公开版：不带 student / price
            })

    out = Path(json_path).expanduser()
    if not out.is_absolute():
        out = Path.cwd() / out
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print("JSON 已输出：{}".format(out))
    return out


# ------------------------------ 主流程 ------------------------------
def parse_month(text):
    m = re.match(r"\s*(\d{4})\s*[-/年.]\s*(\d{1,2})\s*月?\s*$", text or "")
    if not m:
        sys.exit("--month 格式应为 2026-11 这种，收到：{}".format(text))
    year, month = int(m.group(1)), int(m.group(2))
    if not 1 <= month <= 12:
        sys.exit("月份要在 1–12 之间：{}".format(text))
    return year, month


def main():
    global RATE_BAND, RATE_TRIAL, COMMISSION, BASE_SALARY, DEFAULT_ONE_RATE
    ap = argparse.ArgumentParser(description="金山文档课表提取器")
    ap.add_argument("--name", default=TARGET_NAME, help="要筛选的姓名，默认 %s" % TARGET_NAME)
    ap.add_argument("--days", type=int, default=DAYS, help="未来天数，默认 %d" % DAYS)
    ap.add_argument("--file", default=None, help="本地源 xlsx 路径（不传则自动找/自动下载）")
    ap.add_argument("--out", default=None, help="输出文件路径（默认桌面）")
    ap.add_argument("--start-tomorrow", action="store_true", help="从明天开始算（默认含今天）")
    ap.add_argument("--with-all", action="store_true", help="额外输出一个「全部」工作表")
    ap.add_argument("--stats", action="store_true",
                    help="统计某个月课时与课酬，不导出七天表")
    ap.add_argument("--month", default=None,
                    help="统计哪个月，格式 2026-11（默认本月）")
    ap.add_argument("--no-stats", action="store_true",
                    help="只出课表，不在同一文件里附带本月课时统计")
    ap.add_argument("--refresh", action="store_true",
                    help="先从金山文档重下最新课表（需要 WPS_COOKIE 或 cookie.txt），失败则用本地缓存")
    ap.add_argument("--as-of", default=None,
                    help="课酬算到哪天为止，如 2026-11-15（默认今天）")
    ap.add_argument("--rate-band", type=float, default=RATE_BAND,
                    help="班课到手价，默认 180 元/节")
    ap.add_argument("--rate-trial", type=float, default=RATE_TRIAL,
                    help="体验课（名字带 T）到手价，默认 160 元/节")
    ap.add_argument("--commission", type=float, default=COMMISSION,
                    help="一对一课时费提成比例，默认 0.35")
    ap.add_argument("--base-salary", type=float, default=BASE_SALARY,
                    help="底薪，默认 1600 元/月")
    ap.add_argument("--rate-one", type=float, default=DEFAULT_ONE_RATE,
                    help="名单外学生的一对一兜底到手价，默认 210 元/节")
    ap.add_argument("--json-out", default=None,
                    help="输出 JSON 文件路径（供手机端读取）；GitHub Actions 中默认 timetable.json")
    args = ap.parse_args()

    RATE_BAND, RATE_TRIAL = args.rate_band, args.rate_trial
    COMMISSION = args.commission
    BASE_SALARY = args.base_salary
    DEFAULT_ONE_RATE = args.rate_one

    today = dt.date.today()
    as_of = today
    if args.as_of:
        m = re.match(r"\s*(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})", args.as_of)
        if not m:
            sys.exit("--as-of 格式应为 2026-11-15 这种，收到：{}".format(args.as_of))
        as_of = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    # 是否输出 JSON
    json_path = args.json_out
    if json_path is None and os.environ.get("GITHUB_ACTIONS") == "true":
        json_path = "timetable.json"

    # ---------- 统计模式 ----------
    if args.stats:
        if args.month:
            year, month = parse_month(args.month)
        else:
            year, month = today.year, today.month

        src_wb = openpyxl.load_workbook(get_source(args.file, args.refresh), data_only=True)
        detail, summary, sheets = collect_month(src_wb, args.name, year, month, today)
        if not detail:
            sys.exit("{} 在 {}年{}月 没有排课记录。".format(args.name, year, month))
        print("数据来源工作表：{}".format("、".join(sheets)))
        print_stats(args.name, year, month, detail, summary, as_of)

        out_path = Path(args.out).expanduser() if args.out else (
            resolve_desktop() / "{}_{}年{}月课时与课酬.xlsx".format(args.name, year, month))
        out_path.parent.mkdir(parents=True, exist_ok=True)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        write_month_sheet(wb, args.name, year, month, detail, summary, as_of,
                          "{}年{}月课时与课酬".format(year, month))
        wb.save(out_path)
        print("统计表已导出：{}".format(out_path))

        # 统计模式下也输出 JSON（包含本月统计 + 未来 N 天课表）
        if json_path:
            include_today = INCLUDE_TODAY and not args.start_tomorrow
            first_day = today if include_today else today + dt.timedelta(days=1)
            last_day = first_day + dt.timedelta(days=max(args.days, 1) - 1)
            range_detail = collect_range(src_wb, args.name, first_day, last_day, today)
            output_json(args.name, today, detail, summary, range_detail, year, month, json_path)
        return

    include_today = INCLUDE_TODAY and not args.start_tomorrow
    first_day = today if include_today else today + dt.timedelta(days=1)
    last_day = first_day + dt.timedelta(days=max(args.days, 1) - 1)

    src_path = get_source(args.file, args.refresh)
    src_wb = openpyxl.load_workbook(src_path, data_only=True)

    # 1) 扫一遍所有工作表，收集命中行
    hits = []
    for ws in src_wb.worksheets:
        header_row, name_col, date_col, rows_all = extract_rows(ws, args.name, today)
        if not rows_all:
            continue
        rows_window = [(r, d) for r, d in rows_all
                       if d is not None and first_day <= d <= last_day]
        hits.append((ws, header_row, rows_all, rows_window))
        print("命中工作表「{}」：全部 {} 行，其中 {} ~ {} 共 {} 行"
              .format(ws.title, len(rows_all), first_day, last_day, len(rows_window)))

    if not hits:
        sys.exit("没有找到姓名为「{}」的行，请检查姓名写法或源文件。".format(args.name))

    # 2) 生成结果
    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    for ws, hr, rows_all, rows_window in hits:
        if not rows_window:
            continue
        dst = out_wb.create_sheet("{}-未来{}天".format(args.name, args.days)[:31])
        build_sheet(ws, dst, [r for r, _ in rows_window], hr)

    if args.with_all:
        for ws, hr, rows_all, rows_window in hits:
            dst = out_wb.create_sheet("{}-全部".format(args.name)[:31])
            build_sheet(ws, dst, [r for r, _ in rows_all], hr)

    if not out_wb.worksheets:
        sys.exit("定位到了「{}」的行，但 {} ~ {} 这 {} 天里没有排课。"
                 .format(args.name, first_day, last_day, args.days))

    # 3) 同一个月度课时/课酬也放进同一个文件（--no-stats 可关掉）
    month_detail, month_summary = {}, {}
    y, m = (today.year, today.month) if not args.month else parse_month(args.month)
    if not args.no_stats:
        month_detail, month_summary, _ = collect_month(src_wb, args.name, y, m, today)
        if month_detail:
            write_month_sheet(out_wb, args.name, y, m, month_detail, month_summary, as_of,
                              "{}年{}月课时与课酬".format(y, m))
            all_sum, done_sum = split_month(month_detail, as_of)
            rows = summarize_by_category(month_detail, as_of)
            m_all = sum(r["all_m"] for r in rows)
            m_done = sum(r["done_m"] for r in rows)
            print("已附带 {}年{}月：全月 班课 {} / 一对一 {} / 体验课 {} 节，课酬 {} 元，"
                  "加底薪到手 {} 元；截至 {} 已上 {} 节、课酬 {} 元、到手 {} 元"
                  .format(y, m, all_sum["班课"], all_sum["一对一"], all_sum["体验课"],
                          fmt_money(m_all), fmt_money(m_all + BASE_SALARY),
                          as_of.strftime("%m-%d"), sum(done_sum.values()),
                          fmt_money(m_done), fmt_money(m_done + BASE_SALARY)))

    # 4) 保存 xlsx
    out_path = Path(args.out).expanduser() if args.out else (
        resolve_desktop() / "{}_课表_{}.xlsx".format(
            args.name, first_day.strftime("%Y%m%d")))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_wb.save(out_path)
    print("\n已导出：{}".format(out_path))
    for ws in out_wb.worksheets:
        print("  - 工作表「{}」：{} 行".format(ws.title, ws.max_row))

    # 5) 输出 JSON（供手机端 Scriptable 读取）
    if json_path:
        if not month_detail:
            month_detail, month_summary, _ = collect_month(src_wb, args.name, y, m, today)
        range_detail = collect_range(src_wb, args.name, first_day, last_day, today)
        output_json(args.name, today, month_detail, month_summary,
                    range_detail, y, m, json_path)


if __name__ == "__main__":
    main()
