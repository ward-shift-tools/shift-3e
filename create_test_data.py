#!/usr/bin/env python3
"""3E病棟テストデータ生成（2026年5月）
スタッフ約40名、クラス5種、遅出可フラグ付き
列: 名前, クラス, 遅出可, 週勤務, 前月末, 夜勤Min, 夜勤Max, 連勤Max, 勤務曜日, 祝日不可, 土日不可
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

import calendar
from datetime import date
import jpholiday

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

year, month = 2026, 5
num_days = calendar.monthrange(year, month)[1]  # 31

wb = Workbook()
thin = Side(style="thin")
bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
staff_hdr_fill = PatternFill("solid", fgColor="548235")
staff_hdr_font = Font(bold=True, color="FFFFFF", size=10)
hdr_fill = PatternFill("solid", fgColor="4472C4")
hdr_font = Font(bold=True, color="FFFFFF", size=11)

# ====== スタッフデータ ======
# 列: 名前, クラス, 遅出可, 週勤務, 前月末, 夜勤Min, 夜勤Max, 連勤Max, 勤務曜日, 祝日不可, 土日不可
# クラス: ERリーダー / リーダー / ER可 / HCU可 / 病棟可
staff_data = [
    # ── ERリーダー（ER・HCU・病棟・共リーダー全て可）──
    ["山田",   "ERリーダー", "○", None, "明", None, None, None, None, None, None],
    ["中村",   "ERリーダー", "○", None, None, None, None, None, None, None, None],
    ["田中",   "ERリーダー", None, None, "夜", None, None, None, None, None, None],
    ["佐藤",   "ERリーダー", "○", None, None, None, None, None, None, None, None],
    ["鈴木",   "ERリーダー", None, None, None, None, None, None, None, None, None],

    # ── リーダー（HCU・病棟・共リーダー可、ER不可）──
    ["伊藤",   "リーダー", "○", None, "明", None, None, None, None, None, None],
    ["渡辺",   "リーダー", "○", None, None, None, None, None, None, None, None],
    ["小林",   "リーダー", None, None, None, None, None, None, None, None, None],
    ["加藤",   "リーダー", "○", None, None, None, None, None, None, None, None],
    ["吉田",   "リーダー", None, None, "夜", None, None, None, None, None, None],
    ["山口",   "リーダー", "○", None, None, None, None, None, None, None, None],
    ["松本",   "リーダー", None, None, None, None, None, None, None, None, None],
    ["井上",   "リーダー", "○", None, "明", None, None, None, None, None, None],

    # ── ER可（ER・HCU・病棟可、リーダー不可）──
    ["木村",   "ER可", "○", None, None, None, None, None, None, None, None],
    ["林",     "ER可", None, None, None, None, None, None, None, None, None],
    ["清水",   "ER可", "○", None, None, None, None, None, None, None, None],
    ["山本",   "ER可", None, None, "夜", None, None, None, None, None, None],
    ["中山",   "ER可", "○", None, None, None, None, None, None, None, None],
    ["小川",   "ER可", None, None, None, None, None, None, None, None, None],
    ["池田",   "ER可", None, None, None, None, None, None, None, None, None],

    # ── HCU可（HCU・病棟可、ER・リーダー不可）──
    ["橋本",   "HCU可", None, None, None, None, None, None, None, None, None],
    ["石田",   "HCU可", None, None, "明", None, None, None, None, None, None],
    ["前田",   "HCU可", "○", None, None, None, None, None, None, None, None],
    ["岡田",   "HCU可", None, None, None, None, None, None, None, None, None],
    ["長田",   "HCU可", None, None, None, None, None, None, None, None, None],
    ["村田",   "HCU可", "○", None, None, None, None, None, None, None, None],
    ["藤田",   "HCU可", None, None, None, None, None, None, None, None, None],
    ["坂本",   "HCU可", None, None, "夜", None, None, None, None, None, None],

    # ── 病棟可（病棟のみ）──
    ["斎藤",   "病棟可", None, None, None, None, None, None, None, None, None],
    ["西田",   "病棟可", None, None, "明", None, None, None, None, None, None],
    ["福田",   "病棟可", None, None, None, None, None, None, None, None, None],
    ["中田",   "病棟可", None, None, None, None, None, None, None, None, None],
    ["原田",   "病棟可", None, None, None, None, None, None, None, None, None],
    ["宮田",   "病棟可", None, None, None, None, None, None, None, None, None],
    ["田村",   "病棟可", None, None, None, None, None, None, None, None, None],
    ["川村",   "病棟可", None, None, None, None, None, None, None, None, None],
    ["高橋",   "病棟可", None, None, None, None, None, None, None, None, None],
    ["三宅",   "病棟可", None, None, None, None, None, None, None, None, None],
    ["上田",   "病棟可", None, None, None, None, None, None, None, None, None],
    ["青木",   "病棟可", None, None, "夜", None, None, None, None, None, None],
]

staff_headers = ["名前", "クラス", "遅出可", "週勤務", "前月末", "夜勤Min", "夜勤Max", "連勤Max", "勤務曜日", "祝日不可", "土日不可"]

# ====== 勤務希望テストデータ ======
# 各スタッフに適当な希望を設定
import random
random.seed(42)
valid_wishes = ["休", "夜", "日", "遅希", "夜不"]

def make_requests(staff_data, year, month, num_days):
    rows = []
    for sd in staff_data:
        name = sd[0]
        row = [name] + [""] * num_days
        # 月3〜5日のランダム希望
        days_with_request = random.sample(range(1, num_days + 1), k=random.randint(2, 4))
        for d in days_with_request:
            w = random.choice(valid_wishes[:3])  # 休/夜/日
            row[d] = w
        rows.append(row)
    return rows

request_rows = make_requests(staff_data, year, month, num_days)

# ====== Excelシート作成 ======
ws_si = wb.active
ws_si.title = "スタッフ情報"

# ヘッダー行1: タイトル
ws_si.merge_cells(f"A1:{get_column_letter(len(staff_headers))}1")
ws_si["A1"] = f"3E スタッフ情報 — {year}年{month}月"
ws_si["A1"].font = Font(bold=True, size=12)
ws_si["A1"].fill = hdr_fill
ws_si["A1"].font = hdr_font
ws_si["A1"].alignment = Alignment(horizontal="center")

# ヘッダー行2: 説明
ws_si.merge_cells(f"A2:{get_column_letter(len(staff_headers))}2")
ws_si["A2"] = "クラス: ERリーダー / リーダー / ER可 / HCU可 / 病棟可"
ws_si["A2"].font = Font(size=9, italic=True)

# ヘッダー行3: 列名
for ci, h in enumerate(staff_headers, 1):
    c = ws_si.cell(row=3, column=ci, value=h)
    c.font = staff_hdr_font
    c.fill = staff_hdr_fill
    c.alignment = Alignment(horizontal="center")
    c.border = bdr

# データ行
for ri, row in enumerate(staff_data, 4):
    for ci, val in enumerate(row, 1):
        c = ws_si.cell(row=ri, column=ci, value=val)
        c.border = bdr
        c.alignment = Alignment(horizontal="center")

# 列幅
ws_si.column_dimensions["A"].width = 10
ws_si.column_dimensions["B"].width = 12
for ci in range(3, len(staff_headers) + 1):
    ws_si.column_dimensions[get_column_letter(ci)].width = 9

# ====== 勤務希望シート ======
ws_rq = wb.create_sheet("勤務希望")

# 希望記号凡例
ws_rq.merge_cells(f"A1:G1")
ws_rq["A1"] = f"3E 勤務希望 — {year}年{month}月"
ws_rq["A1"].font = hdr_font
ws_rq["A1"].fill = hdr_fill
ws_rq["A1"].alignment = Alignment(horizontal="center")

ws_rq.merge_cells(f"A2:{get_column_letter(num_days + 2)}2")
ws_rq["A2"] = "希望記号: 休=公休希望 / 夜=夜勤希望 / 日=日勤希望 / 遅希=遅出希望 / 夜不=夜勤不可"
ws_rq["A2"].font = Font(size=9, italic=True)

# ヘッダー行3: 名前 + 日付
ws_rq.cell(row=3, column=1, value="名前").font = staff_hdr_font
ws_rq.cell(row=3, column=1).fill = staff_hdr_fill
ws_rq.cell(row=3, column=1).alignment = Alignment(horizontal="center")
ws_rq.cell(row=3, column=2, value="クラス").font = staff_hdr_font
ws_rq.cell(row=3, column=2).fill = staff_hdr_fill
ws_rq.cell(row=3, column=2).alignment = Alignment(horizontal="center")

_wdj = ["月", "火", "水", "木", "金", "土", "日"]
_fwd = date(year, month, 1).weekday()
_holidays = set()
for d in range(1, num_days + 1):
    if jpholiday.is_holiday(date(year, month, d)):
        _holidays.add(d)

for d in range(1, num_days + 1):
    wd = _wdj[(_fwd + d - 1) % 7]
    hol = "祝" if d in _holidays else ""
    cell = ws_rq.cell(row=3, column=d + 2, value=f"{d}\n{wd}{hol}")
    cell.font = staff_hdr_font
    cell.fill = staff_hdr_fill
    cell.alignment = Alignment(horizontal="center", wrap_text=True)

# データ行
for ri, row in enumerate(request_rows, 4):
    name = row[0]
    cls = next((sd[1] for sd in staff_data if sd[0] == name), "")
    ws_rq.cell(row=ri, column=1, value=name).border = bdr
    ws_rq.cell(row=ri, column=2, value=cls).border = bdr
    for d in range(1, num_days + 1):
        c = ws_rq.cell(row=ri, column=d + 2, value=row[d] if row[d] else None)
        c.border = bdr
        c.alignment = Alignment(horizontal="center")

ws_rq.column_dimensions["A"].width = 10
ws_rq.column_dimensions["B"].width = 12
for d in range(1, num_days + 1):
    ws_rq.column_dimensions[get_column_letter(d + 2)].width = 5

ws_rq.row_dimensions[3].height = 32

# ====== 設定シート ======
ws_cfg = wb.create_sheet("設定")
ws_cfg["A1"] = "設定"
ws_cfg["A1"].font = Font(bold=True, size=12)

cfg_data = [
    ("対象年", year, ""),
    ("対象月", month, ""),
    ("公休日数（空=自動）", "", ""),
    ("病棟最低人数（平日）", 4, ""),
    ("HCU最低人数（平日）", 2, ""),
    ("ER最低人数（平日）", 3, ""),
    ("病棟最低人数（休日）", 4, ""),
    ("HCU最低人数（休日）", 2, ""),
    ("夜勤上限/月", 5, ""),
    ("夜勤推奨/月", 4, ""),
    ("最大連勤", 5, ""),
    ("推奨連勤", 4, ""),
    ("計算時間上限（秒）", 120, ""),
    ("祝日（追加: カンマ区切り日付）", "", ""),
]
ws_cfg["A2"] = "項目"
ws_cfg["B2"] = "値"
ws_cfg["C2"] = "備考"
for ci_idx in range(1, 4):
    ws_cfg.cell(row=2, column=ci_idx).font = Font(bold=True)

for ri, (label, val, note) in enumerate(cfg_data, 3):
    ws_cfg.cell(row=ri, column=1, value=label)
    ws_cfg.cell(row=ri, column=2, value=val if val != "" else None)
    ws_cfg.cell(row=ri, column=3, value=note)

ws_cfg.column_dimensions["A"].width = 28
ws_cfg.column_dimensions["B"].width = 12
ws_cfg.column_dimensions["C"].width = 30

# ====== 保存 ======
outfile = f"3e_test_{year}_{month:02d}.xlsx"
wb.save(outfile)
print(f"✅ テストデータ生成完了: {outfile}")
print(f"   スタッフ数: {len(staff_data)}名")
print(f"   期間: {year}年{month}月 ({num_days}日)")
