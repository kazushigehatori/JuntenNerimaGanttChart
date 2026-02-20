"""
手術室ガントチャート生成スクリプト
===================================
手術実施データ（Excel）を読み込み、日付ごとに手術室の稼働状況を
ガントチャートとしてExcelファイルに出力します。

使い方:
    python generate_gantt_chart.py

入力: ガントチャート-元データ.xlsx（同一フォルダに配置）
出力: 手術室ガントチャート-結果.xlsx（同一フォルダに生成）
"""

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy
from datetime import datetime, timedelta
import os
import sys

# ========== 設定 ==========
# PyInstaller exe の場合は exe の場所、通常実行の場合はスクリプトの場所を基準にする
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(BASE_DIR, "ガントチャート-元データ.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "手術室ガントチャート-結果.xlsx")

# 手術室の表示順
ROOM_ORDER = ["01A", "01B", "02", "03", "05", "06", "07", "08", "09", "10", "ｱﾝｷﾞｵ"]

# 時間範囲: 8:00 ~ 22:00（10分刻み）
TIME_START_HOUR = 8
TIME_END_HOUR = 22
COLS_PER_HOUR = 6  # 1時間=6列（10分刻み）

# 診療科の略称マッピング
DEPT_SHORT = {
    "総合外科": "総",
    "乳腺外科": "乳",
    "心臓血管外科": "心",
    "整形外科・スポーツ診療科": "整",
    "形成外科": "形",
    "産科・婦人科": "産",
    "腎・高血圧内科": "腎",
    "小児外科": "小",
    "眼科": "眼",
    "循環器内科": "循",
    "呼吸器外科": "呼",
    "泌尿器科": "泌",
    "耳鼻咽喉・頭頚科": "耳",
    "脳神経外科": "脳",
    "麻酔科・ペインクリニック": "麻",
}

# 実施区分の色分け
COLOR_NORMAL = "A0C8E4"   # 水色（定時・臨時）- テンプレートC3準拠
COLOR_EMERGENCY = "FF8CCC" # ピンク（緊急）- テンプレートC4準拠

# フォント設定
FONT_NAME = "Meiryo UI"
HEADER_FONT = Font(name=FONT_NAME, size=9, bold=True)
CELL_FONT = Font(name=FONT_NAME, size=7)
SMALL_FONT = Font(name=FONT_NAME, size=6)
DATE_FONT = Font(name=FONT_NAME, size=9, bold=True)

# 罫線
THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
THIN_SIDE = Side(style='thin')
HAIR_SIDE = Side(style='hair')


def time_to_col(time_val, col_offset=4):
    """時刻をExcel列番号に変換（列Dが8:00開始）"""
    import datetime as dt_module
    if isinstance(time_val, str):
        parts = time_val.split(":")
        hours, minutes = int(parts[0]), int(parts[1])
    elif isinstance(time_val, timedelta):
        total_seconds = int(time_val.total_seconds())
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
    elif isinstance(time_val, dt_module.time):
        hours = time_val.hour
        minutes = time_val.minute
    else:
        hours = time_val.hour
        minutes = time_val.minute

    total_minutes = (hours - TIME_START_HOUR) * 60 + minutes
    col = col_offset + int(total_minutes / 10)
    return col


def shorten_surgery_name(name, max_chars=20):
    """手術名を短縮"""
    if not name or not isinstance(name, str):
        return ""
    # 括弧内の詳細を省略
    result = name
    if len(result) > max_chars:
        result = result[:max_chars]
    return result


def calculate_utilization(day_data, rooms, weekday=""):
    """稼働率を計算
    - 平日: 9:00-17:00（8h×9室）、土曜: 9:00-13:00（4h×9室）
    - 01A・01Bは各0.5室換算、アンギオ室は除外
    """
    import datetime as dt_module

    if "土" in weekday:
        calc_start = 9 * 60
        calc_end = 13 * 60
        standard_minutes = 240  # 4時間
    else:
        calc_start = 9 * 60
        calc_end = 17 * 60
        standard_minutes = 480  # 8時間

    # 分母: 9室換算（01A=0.5, 01B=0.5, 他8室=8.0 → 合計9.0、アンギオ除外）
    room_count = 9.0
    total_available = standard_minutes * room_count
    total_used = 0

    # 室ごとの重み（01A/01Bは0.5、アンギオは除外）
    ROOM_WEIGHT = {
        "01A": 0.5,
        "01B": 0.5,
        "ｱﾝｷﾞｵ": 0,  # 除外
    }

    for _, row in day_data.iterrows():
        try:
            room_name = str(row["実施手術室名"])
            weight = ROOM_WEIGHT.get(room_name, 1.0)
            if weight == 0:
                continue

            t = row["入室時刻"]
            if isinstance(t, str):
                parts = t.split(":")
                start_min = int(parts[0]) * 60 + int(parts[1])
            elif isinstance(t, timedelta):
                start_min = int(t.total_seconds()) // 60
            elif isinstance(t, dt_module.time):
                start_min = t.hour * 60 + t.minute
            else:
                start_min = t.hour * 60 + t.minute

            t2 = row["麻酔終了時刻"]
            if isinstance(t2, str):
                parts = t2.split(":")
                end_min = int(parts[0]) * 60 + int(parts[1])
            elif isinstance(t2, timedelta):
                end_min = int(t2.total_seconds()) // 60
            elif isinstance(t2, dt_module.time):
                end_min = t2.hour * 60 + t2.minute
            else:
                end_min = t2.hour * 60 + t2.minute

            clipped_start = max(start_min, calc_start)
            clipped_end = min(end_min, calc_end)
            if clipped_end > clipped_start:
                total_used += (clipped_end - clipped_start) * weight
        except Exception:
            pass

    if total_available > 0:
        return total_used / total_available
    return 0


def write_day_block(ws, start_row, date_str, weekday, day_data, rooms):
    """1日分のガントチャートブロックを書き込む"""

    # --- ヘッダ行（時間軸） ---
    header_row = start_row

    # 稼働率を計算（曜日を渡す）
    utilization = calculate_utilization(day_data, rooms, weekday)

    # 最終列（22:00の6列分の末尾）
    last_col = 4 + (TIME_END_HOUR - TIME_START_HOUR + 1) * COLS_PER_HOUR - 1

    # 列B: 日付
    ws.cell(row=header_row, column=2, value="日付")
    ws.cell(row=header_row, column=2).font = HEADER_FONT
    ws.cell(row=header_row, column=2).border = Border(top=THIN_SIDE, bottom=THIN_SIDE, left=THIN_SIDE)
    ws.cell(row=header_row, column=2).alignment = Alignment(horizontal='center', vertical='center')

    # 列C: 部屋名
    ws.cell(row=header_row, column=3, value="部屋名")
    ws.cell(row=header_row, column=3).font = HEADER_FONT
    ws.cell(row=header_row, column=3).border = Border(top=THIN_SIDE, bottom=THIN_SIDE)
    ws.cell(row=header_row, column=3).alignment = Alignment(horizontal='center', vertical='center')

    # 時間ヘッダ（8:00～22:00すべて6列結合）
    for h in range(TIME_START_HOUR, TIME_END_HOUR + 1):
        col = 4 + (h - TIME_START_HOUR) * COLS_PER_HOUR
        time_label = h * 100
        cell = ws.cell(row=header_row, column=col, value=time_label)
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(top=THIN_SIDE, bottom=THIN_SIDE, left=THIN_SIDE, right=THIN_SIDE)
        # 全時間帯（22:00含む）で6列結合
        end_col = col + COLS_PER_HOUR - 1
        ws.merge_cells(start_row=header_row, start_column=col, end_row=header_row, end_column=end_col)

    # --- 部屋ごとの行 ---
    last_room_row = start_row + len(rooms)  # 最後の部屋行

    for room_idx, room in enumerate(rooms):
        row = start_row + 1 + room_idx
        room_data = day_data[day_data["実施手術室名"] == room]
        is_last_room = (room_idx == len(rooms) - 1)

        # 日付列（最初の部屋行のみ表示、全部屋を縦結合）
        if room_idx == 0:
            if "土" in weekday:
                util_label = f"{date_str}\n{utilization:.1%}"
            else:
                util_label = f"{date_str}\n{utilization:.1%}"
            ws.cell(row=row, column=2, value=util_label)
            ws.cell(row=row, column=2).font = DATE_FONT
            ws.cell(row=row, column=2).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws.cell(row=row, column=2).border = Border(top=THIN_SIDE, left=THIN_SIDE, bottom=THIN_SIDE, right=THIN_SIDE)
            if len(rooms) > 1:
                ws.merge_cells(start_row=row, start_column=2, end_row=row + len(rooms) - 1, end_column=2)

        # 部屋名
        bottom = THIN_SIDE if is_last_room else None
        ws.cell(row=row, column=3, value=room)
        ws.cell(row=row, column=3).font = CELL_FONT
        ws.cell(row=row, column=3).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=row, column=3).border = Border(left=THIN_SIDE, right=THIN_SIDE, bottom=bottom)

        # 時間軸の罫線（1時間ごとに縦線 + 最下行は下罫線）
        for c in range(4, last_col + 1):
            # 1時間ごとの左罫線判定
            is_hour_start = ((c - 4) % COLS_PER_HOUR == 0)
            left = THIN_SIDE if is_hour_start else None
            # 最終列の右罫線
            right = THIN_SIDE if c == last_col else None
            # 最下行の下罫線
            bot = THIN_SIDE if is_last_room else None

            existing = ws.cell(row=row, column=c).border
            ws.cell(row=row, column=c).border = Border(
                left=left if left else existing.left,
                right=right if right else existing.right,
                bottom=bot if bot else existing.bottom,
                top=existing.top
            )

        # 手術バーを描画
        for _, op in room_data.iterrows():
            try:
                start_col = time_to_col(op["入室時刻"])
                end_col = time_to_col(op["麻酔終了時刻"])

                start_col = max(start_col, 4)
                end_col = min(end_col, last_col)

                if end_col <= start_col:
                    end_col = start_col + 1

                dept_short = DEPT_SHORT.get(op["執刀診療科名"], op["執刀診療科名"][0])

                urgency = op.get("実施申込区分", "定時")
                if urgency == "緊急":
                    color = COLOR_EMERGENCY
                else:
                    color = COLOR_NORMAL
                fill = PatternFill('solid', fgColor=color)

                surgery_name = op.get("実施手術名０１", "")
                if not isinstance(surgery_name, str):
                    surgery_name = ""
                short_name = shorten_surgery_name(surgery_name, max_chars=40)

                bar_label = f"【{dept_short}】-{short_name}"
                bar_font = Font(name=FONT_NAME, size=6, color="000000")

                for c in range(start_col, end_col + 1):
                    cell = ws.cell(row=row, column=c)
                    cell.fill = fill
                    # 塗りつぶし後も罫線を保持
                    is_hour_start = ((c - 4) % COLS_PER_HOUR == 0)
                    left = THIN_SIDE if is_hour_start else None
                    right = THIN_SIDE if c == last_col else None
                    bot = THIN_SIDE if is_last_room else None
                    cell.border = Border(left=left, right=right, bottom=bot)

                ws.cell(row=row, column=start_col, value=bar_label)
                ws.cell(row=row, column=start_col).font = bar_font
                ws.cell(row=row, column=start_col).alignment = Alignment(vertical='center', wrap_text=False)

            except Exception:
                pass

    return start_row + 1 + len(rooms)


def main():
    print(f"入力ファイル読み込み: {INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE, sheet_name="ガントチャートデータ", dtype={"実施手術室名": str})

    # 日付でソート
    df["手術実施日_sort"] = pd.to_datetime(df["手術実施日"], format="%Y/%m/%d")
    df = df.sort_values(["手術実施日_sort", "実施手術室名", "入室時刻"])

    dates = df["手術実施日"].unique()
    weekday_map = dict(zip(df["手術実施日"], df["曜日"]))

    wb = Workbook()
    ws = wb.active
    ws.title = "手術室ガントチャート"

    # 列幅設定
    ws.column_dimensions['A'].width = 2
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 6

    # 時間列の幅（8:00-22:00、各時間6列 = 15時間×6列 = 90列）
    total_time_cols = (TIME_END_HOUR - TIME_START_HOUR + 1) * COLS_PER_HOUR
    for i in range(4, 4 + total_time_cols):
        ws.column_dimensions[get_column_letter(i)].width = 2.5

    # 行の高さ
    ws.sheet_properties.defaultRowHeight = 18

    # タイトル
    ws.cell(row=1, column=2, value="手術室 ガントチャート（2025年9月）")
    ws.cell(row=1, column=2).font = Font(name=FONT_NAME, size=14, bold=True)

    # 凡例
    legend_row = 2
    legend_col = 2
    ws.cell(row=legend_row, column=legend_col, value="■凡例:")
    ws.cell(row=legend_row, column=legend_col).font = Font(name=FONT_NAME, size=8, bold=True)

    cell = ws.cell(row=legend_row, column=legend_col + 2, value="定時・臨時")
    cell.fill = PatternFill('solid', fgColor=COLOR_NORMAL)
    cell.font = Font(name=FONT_NAME, size=8)
    cell.alignment = Alignment(horizontal='center')

    cell = ws.cell(row=legend_row, column=legend_col + 4, value="緊急")
    cell.fill = PatternFill('solid', fgColor=COLOR_EMERGENCY)
    cell.font = Font(name=FONT_NAME, size=8)
    cell.alignment = Alignment(horizontal='center')

    ws.cell(row=legend_row + 1, column=legend_col, value="※稼働率 = 平日:9:00-17:00（8h×9室）、土曜:9:00-13:00（4h×9室）")
    ws.cell(row=legend_row + 1, column=legend_col).font = Font(name=FONT_NAME, size=8)
    ws.cell(row=legend_row + 2, column=legend_col, value="※01A・01Bは各0.5室換算、アンギオ室は除外")
    ws.cell(row=legend_row + 2, column=legend_col).font = Font(name=FONT_NAME, size=8)

    # 日付ごとにガントチャートブロックを生成
    current_row = 6
    for date_str in dates:
        day_data = df[df["手術実施日"] == date_str]
        weekday = weekday_map.get(date_str, "")
        # 曜日を1文字に
        weekday_short = weekday.replace("曜日", "") if isinstance(weekday, str) else ""

        # MM/DD形式に変換
        try:
            dt = pd.to_datetime(date_str)
            date_display = f"{dt.month:02d}/{dt.day:02d}({weekday_short})"
        except Exception:
            date_display = date_str

        next_row = write_day_block(ws, current_row, date_display, weekday_short, day_data, ROOM_ORDER)
        current_row = next_row + 1  # 1行空けて次の日

    # 印刷設定
    ws.sheet_view.zoomScale = 80
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    # 元データの「ガントチャートデータ」シートをコピー
    src_wb = load_workbook(INPUT_FILE)
    if "ガントチャートデータ" in src_wb.sheetnames:
        src_ws = src_wb["ガントチャートデータ"]
        dst_ws = wb.create_sheet("ガントチャートデータ")

        # 列幅をコピー
        for col_letter, dim in src_ws.column_dimensions.items():
            dst_ws.column_dimensions[col_letter].width = dim.width
            dst_ws.column_dimensions[col_letter].hidden = dim.hidden

        # 行の高さをコピー
        for row_num, dim in src_ws.row_dimensions.items():
            dst_ws.row_dimensions[row_num].height = dim.height
            dst_ws.row_dimensions[row_num].hidden = dim.hidden

        # セルデータと書式をコピー
        for row in src_ws.iter_rows(min_row=1, max_row=src_ws.max_row, max_col=src_ws.max_column):
            for cell in row:
                dst_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    dst_cell.font = copy(cell.font)
                    dst_cell.fill = copy(cell.fill)
                    dst_cell.border = copy(cell.border)
                    dst_cell.alignment = copy(cell.alignment)
                    dst_cell.number_format = cell.number_format

        # 結合セルをコピー
        for merged_range in src_ws.merged_cells.ranges:
            dst_ws.merge_cells(str(merged_range))

    src_wb.close()

    # 保存
    wb.save(OUTPUT_FILE)
    print(f"ガントチャート生成完了: {OUTPUT_FILE}")
    print(f"全{len(dates)}日分のガントチャートを出力しました。")


if __name__ == "__main__":
    main()
