import streamlit as st
import openpyxl
import datetime
import re
import os

# システム環境設定（不具合防止）
os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

st.set_page_config(page_title="シフト配置確認システム", layout="wide")

# ────────────────────────────────────────────────────────
# 📌 週間予定表シートを実際のセル配置から直接読み込む関数
# ────────────────────────────────────────────────────────
def read_weekly_excel_schedule(ws, calendar_data, target_year, target_month):
    """
    アップロードされた「週間予定表」シートの実際のレイアウトから
    サービス1/2/3・ヘルパー1/2/3をそのまま読み取り、該当する曜日の
    日付すべて（その月の全ての火/水/木/…曜日）に反映する。

    列の対応:
        火曜日: C=サービス1 D=ヘルパー1 E=サービス2 F=ヘルパー2 G=サービス3 H=ヘルパー3
        水曜日: I,J,K,L,M,N（同じ並び）
        木曜日: O,P,Q,R,S,T
        金曜日: U,V,W,X,Y,Z
        土曜日: C,D,E,F,G,H
        日曜日: I,J,K,L,M,N
        月曜日: O,P,Q,R,S,T

    行の対応:
        火〜金ブロック: 8行目=5:00 〜 47行目=0:30（30分刻み、40コマ）
        土〜月ブロック: 58行目=5:00 〜 97行目=0:30（30分刻み、40コマ）
    """
    from openpyxl.utils import column_index_from_string

    weekday_blocks = {
        "火": (8, "C", "D", "E", "F", "G", "H"),
        "水": (8, "I", "J", "K", "L", "M", "N"),
        "木": (8, "O", "P", "Q", "R", "S", "T"),
        "金": (8, "U", "V", "W", "X", "Y", "Z"),
        "土": (58, "C", "D", "E", "F", "G", "H"),
        "日": (58, "I", "J", "K", "L", "M", "N"),
        "月": (58, "O", "P", "Q", "R", "S", "T"),
    }
    SLOTS_PER_BLOCK = 40  # 5:00 から 翌0:30 まで 30分刻みで40コマ

    def get_fill_hex(cell):
        """セルの塗りつぶし色をHEXで取得（テーマ色・塗りなしはNone扱い）"""
        fill = cell.fill
        if not fill or fill.fill_type != "solid":
            return None
        fg = fill.fgColor
        if getattr(fg, "type", None) == "rgb":
            rgb = fg.rgb
            if isinstance(rgb, str) and len(rgb) >= 6:
                return "#" + rgb[-6:]
        return None

    # 曜日ごとに1週間分のテンプレート（時間帯→値）を作成
    weekday_template = {}
    for w_str, (base_row, c_s1, c_h1, c_s2, c_h2, c_s3, c_h3) in weekday_blocks.items():
        col_map = {
            "s1": column_index_from_string(c_s1), "h1": column_index_from_string(c_h1),
            "s2": column_index_from_string(c_s2), "h2": column_index_from_string(c_h2),
            "s3": column_index_from_string(c_s3), "h3": column_index_from_string(c_h3),
        }
        template = {}
        last_color = {"s1": None, "s2": None, "s3": None}

        for offset in range(SLOTS_PER_BLOCK):
            row = base_row + offset
            total_min = 5 * 60 + offset * 30
            hour = (total_min // 60) % 24
            minute = total_min % 60
            row_key = "row1" if minute == 0 else "row2"
            template.setdefault(hour, {"row1": {}, "row2": {}})

            for key, col_idx in col_map.items():
                cell = ws.cell(row=row, column=col_idx)
                val = cell.value
                val_str = str(val).strip() if val is not None else ""

                if not val_str:
                    if key in ("s1", "s2", "s3"):
                        last_color[key] = None
                    continue

                template[hour][row_key][key] = val_str

                if key in ("s1", "s2", "s3"):
                    color = get_fill_hex(cell)
                    # 「〃」など継続コマで色が塗られていない場合は直前の色を引き継ぐ
                    if not color and val_str == "〃":
                        color = last_color[key]
                    if color:
                        template[hour][row_key][f"{key}_color"] = color
                    last_color[key] = color

        weekday_template[w_str] = template

    weekday_map = {0: "月", 1: "火", 2: "水", 3: "木", 4: "金", 5: "土", 6: "日"}
    for d in range(1, 32):
        try:
            this_date = datetime.date(target_year, target_month, d)
        except ValueError:
            continue
        w_str = weekday_map[this_date.weekday()]
        template = weekday_template.get(w_str)
        if not template:
            continue
        for hour, rows in template.items():
            for row_key, values in rows.items():
                for k, v in values.items():
                    calendar_data[d][hour][row_key][k] = v

# ────────────────────────────────────────────────────────
# 📌 勤務表シート（1人1行・1日1列・明/遅/早/日/夜コード方式）を読み込む関数
# ────────────────────────────────────────────────────────
def read_duty_roster_schedule(ws, calendar_data, max_days):
    """
    従来の「勤務表」形式（A列にスタッフ名、B列以降が1日ごとの列、
    各セルに 明/遅/早/日/夜 などのシフトコードが入っている表）を読み込む。
    """
    for d_num in range(1, max_days + 1):
        target_col = 2 + (d_num - 1)
        if target_col > ws.max_column: break

        for r_idx in range(4, 15):
            cell_val = ws.cell(row=r_idx, column=target_col).value
            if cell_val:
                cell_str = str(cell_val).strip()
                if not cell_str or cell_str in ["┃", "│", "↓", "｜", "〃"]: continue
                staff_name = ws.cell(row=r_idx, column=1).value
                if not staff_name: continue
                staff_name = str(staff_name).strip()

                assigned_h = None
                start_hour, start_min = None, 0
                end_hour, end_min = None, 0

                is_ake = "明" in cell_str or "ｍ" in cell_str or "m" in cell_str
                is_osor = "遅" in cell_str
                is_haya = "早" in cell_str or "ｆ" in cell_str or "f" in cell_str
                is_yoru = "夜" in cell_str or "mf" in cell_str
                is_nichi = "日" in cell_str

                if is_ake:
                    assigned_h = "h3"
                    start_hour, start_min = 5, 0
                    end_hour, end_min = 8, 0
                elif is_osor:
                    assigned_h = "h3"
                    start_hour, start_min = 10, 0
                    end_hour, end_min = 18, 30
                elif is_haya:
                    assigned_h = "h1"
                    start_hour, start_min = 7, 0
                    end_hour, end_min = 15, 0
                elif is_nichi:
                    assigned_h = "h2"
                    start_hour, start_min = 14, 0
                    end_hour, end_min = 17, 0
                elif is_yoru:
                    assigned_h = "h1"
                    start_hour, start_min = 17, 0
                    end_hour, end_min = 23, 0

                if assigned_h and start_hour is not None and end_hour is not None:
                    current_time = datetime.datetime(2026, 1, 1, start_hour, start_min)
                    end_time = datetime.datetime(2026, 1, 1, end_hour, end_min)
                    if current_time >= end_time: continue
                    time_slots = []
                    while current_time <= end_time:
                        time_slots.append((current_time.hour, current_time.minute))
                        current_time += datetime.timedelta(minutes=30)
                    for idx, (h, m) in enumerate(time_slots):
                        row_key = "row1" if m == 0 else "row2"
                        if idx == 0 or idx == len(time_slots) - 1:
                            calendar_data[d_num][h][row_key][assigned_h] = staff_name
                        else:
                            current_val = calendar_data[d_num][h][row_key][assigned_h]
                            if not current_val or current_val == "┃":
                                calendar_data[d_num][h][row_key][assigned_h] = "┃"

# ────────────────────────────────────────────────────────
# 🖨️ UI・デザイン用CSS
# ────────────────────────────────────────────────────────
st.markdown("""
    <style>
    div[data-testid="stMainBlockContainer"] { 
        max-width: 96% !important; 
        padding: 4rem 1.5rem 2rem 1.5rem !important;
    }
    div[data-testid="stSelectbox"] label { display: none !important; }
    div[data-testid="stSelectbox"] { margin-top: 0px !important; padding-top: 0px !important; }
    div[data-testid="stSlider"] label { display: none !important; }
    div[data-testid="stSlider"] { margin-top: 0px !important; padding-top: 5px !important; }
    .stButton > button { height: 42px !important; min-height: 42px !important; font-weight: bold !important; }
    
    .calendar-table { table-layout: fixed !important; width: 100% !important; border-collapse: collapse !important; }
    .calendar-table td { vertical-align: middle !important; padding: 4px 1px !important; }
    
    .staff-name-box { 
        display: block !important; 
        white-space: normal !important; 
        word-break: break-all !important; 
        font-size: 11px !important; 
        text-align: center; 
        line-height: 1.1 !important; 
    }
    
    @media print {
        body * { visibility: hidden !important; }
        .print-target, .print-target * { visibility: visible !important; }
        
        .print-target {
            position: fixed !important;
            left: 0 !important;
            top: 0 !important;
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            page-break-before: avoid !important;
            page-break-inside: avoid !important;
        }
        html, body {
            margin: 0 !important;
            padding: 0 !important;
            height: auto !important;
        }
        @page { 
            size: A3 landscape; 
            margin: 8mm !important; 
        }
        .week-print-table { border: 3px solid #000 !important; width: 100% !important; }
        .week-print-table th, .week-print-table td {
            font-size: 10.5px !important; 
            padding: 1px 1px !important;
            height: 22px !important; 
            line-height: 1.05 !important;
            border: 1px solid #000 !important;
        }
        .week-print-table .time-col { font-size: 10px !important; font-weight: bold !important; }
        .week-print-table td.day-start,
        .week-print-table th.day-start {
            border-left: 3px solid #000 !important;
        }
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h2 style='text-align: center; color: #333; margin-top: 0px;'>📅 シフト配置確認システム</h2>", unsafe_allow_html=True)

# ── ファイルアップロード欄を2列に分けて配置 ──
col1, col2 = st.columns(2)
with col1:
    weekly_file = st.file_uploader("📅 【週間予定表】のExcelを選択", type=["xlsx", "xlsm"])
with col2:
    duty_file = st.file_uploader("📋 【勤務表】のExcelを選択", type=["xlsx", "xlsm"])

calendar_data = {}
for d in range(1, 32):
    calendar_data[d] = {}
    for h in range(0, 25):
        calendar_data[d][h] = {
            "row1": {"s1": "", "h1": "", "s2": "", "h2": "", "s3": "", "h3": ""},
            "row2": {"s1": "", "h1": "", "s2": "", "h2": "", "s3": "", "h3": ""}
        }

target_year = 2026
target_month = 6

if weekly_file is not None or duty_file is not None:
    try:
        # --- 週間予定表の読み込み ---
        if weekly_file is not None:
            wb_weekly = openpyxl.load_workbook(weekly_file, data_only=True)
            weekly_sheet_name = st.selectbox("週間予定表のシートを選択してください 👇", options=wb_weekly.sheetnames, key="weekly_sheet")
            ws_weekly = wb_weekly[weekly_sheet_name]
            
            # シート名から月を抽出
            match = re.search(r'(\d+)月|\.(\d+)', weekly_sheet_name)
            if match:
                target_month = int(match.group(1) or match.group(2))
            
            read_weekly_excel_schedule(ws_weekly, calendar_data, target_year, target_month)
            st.success(f"🎉 週間予定表「{weekly_sheet_name}」を反映しました。")

        # --- 勤務表の読み込み ---
        if duty_file is not None:
            wb_duty = openpyxl.load_workbook(duty_file, data_only=True)
            duty_sheet_name = st.selectbox("勤務表のシートを選択してください 👇", options=wb_duty.sheetnames, key="duty_sheet")
            ws_duty = wb_duty[duty_sheet_name]
            
            # 週間予定表が未アップロードの場合は勤務表から月を抽出
            if weekly_file is None:
                match = re.search(r'(\d+)月|\.(\d+)', duty_sheet_name)
                if match:
                    target_month = int(match.group(1) or match.group(2))

            # max_daysの計算
            if target_month in [4, 6, 9, 11]: max_days = 30
            elif target_month == 2: max_days = 28
            else: max_days = 31
            
            read_duty_roster_schedule(ws_duty, calendar_data, max_days)
            st.success(f"🎉 勤務表「{duty_sheet_name}」を重ねて反映しました。")

    except Exception as e:
        st.error(f"Excelの読み込みエラー: {e}")

try: start_offset = (datetime.date(target_year, target_month, 1).weekday() + 1) % 7
except: start_offset = 0

if target_month in [4, 6, 9, 11]: max_days = 30
elif target_month == 2: max_days = 28
else: max_days = 31

def get_bg(val, h_type):
    if not val or val in ["┃", "｜", "↓", "〃"]: return "#ffffff"
    if h_type == "h3": return "#ffc0cb"
    if h_type == "h1": return "#ffff00"
    return "#ffffff"

def wrap_name(val, h_type):
    if not val: return ""
    if val in ["┃", "｜", "↓", "〃"]: return str(val)
    cleaned_val = str(val).replace("（", "(").replace("）", ")")
    return f"<span class='staff-name-box'>{cleaned_val}</span>"

def make_html_table_with_time(day_schedule, font_size="11px", padding="3px", is_large=False):
    html = []
    td_p_style = "padding: 0px 2px !important;" if not is_large else f"padding: {padding} 2px !important;"
    line_h_style = "line-height: 1.0 !important;" if not is_large else "line-height: 1.2 !important;"

    html.append(f"<table class='calendar-table' style='border-collapse: collapse; text-align: center; font-size: {font_size}; font-family: sans-serif; width: 100%; table-layout: fixed; border: 1px solid #333;'>")
    html.append("<tr style='background-color: #f0f0f0; font-weight: bold; border-bottom: 2px solid #333;'>")
    html.append(f"<td style='border: 1px solid #ccc; width: 22%; padding: {padding}; font-size: 10px;'>時間</td>")
    html.append("<td style='border: 1px solid #ccc; width: 13%;'>サ1</td><td style='border: 1px solid #ccc; width: 13%;'>へ1</td>")
    html.append("<td style='border: 1px solid #ccc; width: 13%;'>サ2</td><td style='border: 1px solid #ccc; width: 13%;'>へ2</td>")
    html.append("<td style='border: 1px solid #ccc; width: 13%;'>サ3</td><td style='border: 1px solid #ccc; width: 13%;'>へ3</td>")
    html.append("</tr>")
    
    rows_list = []
    for hour in range(5, 23):
        rows_list.append((f"{hour}:00", day_schedule[hour]["row1"]))
        rows_list.append((f"{hour}:30", day_schedule[hour]["row2"]))
    
    rows_list.append(("23:00", day_schedule[23]["row1"]))
    rows_list.append(("23:30", day_schedule[23]["row2"]))
    rows_list.append(("0:00", day_schedule[0]["row1"]))
    rows_list.append(("0:30", day_schedule[0]["row2"]))
    
    for idx, (time_str, row_data) in enumerate(rows_list):
        is_dash = time_str.endswith(":00")
        border_bottom_style = "border-bottom: 1px dashed #ddd;" if is_dash else "border-bottom: 1px solid #333;"
        if idx == len(rows_list) - 1: border_bottom_style = ""
            
        html.append(f"<tr style='{border_bottom_style} height: 18px;'>")
        html.append(f"<td style='border-right: 1px solid #333; border-left: 1px solid #ccc; background-color: #f9f9f9; font-weight: bold; padding: 2px 0;'>{time_str}</td>")
        
        for s_key, h_key in [("s1", "h1"), ("s2", "h2"), ("s3", "h3")]:
            val_s = row_data[s_key]
            val_h = row_data[h_key]
            bg_color_s = row_data.get(f"{s_key}_color", "#ffffff")
            bg_color_h = get_bg(val_h, h_key)
            html.append(f"<td style='border-left: 1px solid #ccc; border-right: 1px solid #ccc; background-color: {bg_color_s}; {td_p_style} {line_h_style}'>{wrap_name(val_s, '')}</td>")
            html.append(f"<td style='border-left: 1px solid #ccc; border-right: 1px solid #333; font-weight: bold; background-color: {bg_color_h}; {td_p_style} {line_h_style}'>{wrap_name(val_h, h_key)}</td>")
        html.append("</tr>")
    html.append("</table>")
    return "".join(html)

def get_era_label(year, month):
    reiwa_year = year - 2018
    era_str = f"R{reiwa_year}" if reiwa_year >= 1 else str(year)
    return f"{era_str}.{month}月"

era_label = get_era_label(target_year, target_month)

if weekly_file is not None or duty_file is not None:
    view_mode = st.tabs(["📅 1週間表示（時間軸スリム）", "📊 1ヶ月表示（カレンダー）", "🔍 1日集中表示"])

    with view_mode[0]:
        if 'current_week_idx' not in st.session_state: st.session_state.current_week_idx = 0
        weeks_list = ["第1週 (1日〜)", "第2週", "第3週", "第4週", "第5週", "第6週"]
        
        b_col1, b_col2, b_col3 = st.columns([1, 3, 1])
        with b_col1:
            if st.button("← 前の週", use_container_width=True, key="p_wk_k"):
                if st.session_state.current_week_idx > 0: st.session_state.current_week_idx -= 1
        with b_col3:
            if st.button("次の週 →", use_container_width=True, key="n_wk_k"):
                if st.session_state.current_week_idx < len(weeks_list) - 1: st.session_state.current_week_idx += 1
        with b_col2:
            week_option = st.selectbox("週選択", options=weeks_list, index=st.session_state.current_week_idx)
            st.session_state.current_week_idx = weeks_list.index(week_option)
        st.write("---")
        
        st.components.v1.html("""
            <button onclick="parent.window.print();" style="width:100%; height:45px; background-color:#e67e22; color:white; border:none; border-radius:4px; cursor:pointer; font-size:15px; font-weight:bold;">🖨️ この週間シフト表を印刷（自動1枚収め設定）</button>
        """, height=50)
        
        start_d = st.session_state.current_week_idx * 7 + 1 - start_offset
        weekdays_labels = ["日", "月", "火", "水", "木", "金", "土"]
        
        h_sheet = []
        h_sheet.append("<div class='print-target'>")
        h_sheet.append(f"<div style='text-align:right; font-size:12px; font-weight:bold; color:#333; padding:0 2px 2px 0;'>{era_label} 第{st.session_state.current_week_idx + 1}週</div>")
        
        h_sheet.append("<table class='calendar-table week-print-table' style='width:100%; border-collapse:collapse; text-align:center; font-family:sans-serif; table-layout:fixed; border:2px solid #333;'>")
        h_sheet.append("<tr style='background-color: #f0f0f0; font-weight: bold;'>")
        h_sheet.append("<td class='time-col' style='width: 4.2%; padding: 4px 0;'>時間</td>")
        
        for d_o_w in range(7):
            cur_d = start_d + d_o_w
            c_color = '#ff4b4b' if d_o_w == 0 else '#1c83e1' if d_o_w == 6 else '#333333'
            if 1 <= cur_d <= max_days:
                h_sheet.append(f"<td colspan='6' class='day-start' style='color: {c_color}; font-size: 12px; width: 13.68%;'>{weekdays_labels[d_o_w]} ({cur_d}日)</td>")
            else:
                h_sheet.append(f"<td colspan='6' class='day-start' style='color: #aaa; background-color: #fafafa; width: 13.68%;'>{weekdays_labels[d_o_w]} (外)</td>")
        h_sheet.append("</tr>")
        
        h_sheet.append("<tr style='background-color: #f9f9f9; font-size: 9px; height: 14px;'><td style='font-weight:bold; padding:0;'>-</td>")
        for _ in range(7):
            h_sheet.append("<td class='day-start' style='width:1.8%; padding:0;'>サ1</td><td style='font-weight:bold; width:2.1%; background:#fff3cd; padding:0;'>へ1</td>")
            h_sheet.append("<td style='width:1.8%; padding:0;'>サ2</td><td style='font-weight:bold; width:2.1%; background:#fff3cd; padding:0;'>へ2</td>")
            h_sheet.append("<td style='width:1.8%; padding:0;'>サ3</td><td style='font-weight:bold; width:2.1%; background:#fff3cd; padding:0;'>へ3</td>")
        h_sheet.append("</tr>")
        
        r_list = []
        for hour in range(5, 23):
            r_list.append((f"{hour}:00", True))
            r_list.append((f"{hour}:30", False))
        r_list.append(("23:00", True))
        r_list.append(("23:30", False))
        r_list.append(("0:00", True))
        r_list.append(("0:30", False))
            
        for idx, (t_str, is_even) in enumerate(r_list):
            b_style = "border-bottom:1px solid #333;" if not is_even else "border-bottom:1px dashed #ccc;"
            h_sheet.append(f"<tr style='{b_style}'><td class='time-col' style='background-color:#f2f2f2;'>{t_str}</td>")
            
            for d_o_w in range(7):
                cur_d = start_d + d_o_w
                if 1 <= cur_d <= max_days:
                    ds = calendar_data[cur_d]
                    h_num = int(t_str.split(":")[0])
                    rd = ds[h_num]["row1" if t_str.endswith(":00") else "row2"]
                        
                    first_cell = True
                    for sk, hk in [("s1", "h1"), ("s2", "h2"), ("s3", "h3")]:
                        bg_color_s = rd.get(f"{sk}_color", "#ffffff")
                        cls = " class='day-start'" if first_cell else ""
                        h_sheet.append(f"<td{cls} style='background-color:{bg_color_s};'>{wrap_name(rd[sk], '')}</td>")
                        h_sheet.append(f"<td style='font-weight:bold; background-color:{get_bg(rd[hk], hk)};'>{wrap_name(rd[hk], hk)}</td>")
                        first_cell = False
                else:
                    h_sheet.append("<td colspan='6' class='day-start' style='background-color: #fafafa; opacity:0.1;'></td>")
            h_sheet.append("</tr>")
        h_sheet.append("</table>")
        h_sheet.append("</div>")
        
        st.html(f"<div style='border: 1px solid #999; background-color: #ffffff; border-radius: 6px; padding: 5px;'>{''.join(h_sheet)}</div>")

    with view_mode[1]:
        st.components.v1.html('<button onclick="parent.window.print();" style="width:100%; height:42px; background-color:#1c83e1; color:white; border:none; border-radius:4px; cursor:pointer; font-weight:bold;">🖨️ 1ヶ月分印刷プレビュー</button>', height=45)

        st.markdown(f"<div style='text-align:right; font-size:13px; font-weight:bold; color:#555; margin-bottom:2px;'>{era_label}</div>", unsafe_allow_html=True)

        weekdays = ["日", "月", "火", "水", "木", "金", "土"]
        header_cols = st.columns(7)
        for i, day in enumerate(weekdays):
            color = '#ff4b4b' if day == '日' else '#1c83e1' if day == '土' else '#333333'
            header_cols[i].markdown(f"<h3 style='text-align: center; color: {color}; margin: 0;'>{day}</h3>", unsafe_allow_html=True)
        st.write("---")

        day_pointer = 1
        for week in range(6):
            if day_pointer > max_days: break
            row_cols = st.columns(7)
            for day_of_week in range(7):
                current_cell_idx = week * 7 + day_of_week
                with row_cols[day_of_week]:
                    if start_offset <= current_cell_idx and day_pointer <= max_days:
                        st.markdown(f"<h4 style='margin: 0 0 5px 0;'><b>{day_pointer}日</b></h4>", unsafe_allow_html=True)
                        table_html = make_html_table_with_time(calendar_data[day_pointer], font_size="9px", padding="1px", is_large=False)
                        st.html(f"<div style='border: 1px solid #999; height: 430px; overflow-y: auto; background-color: #ffffff;'>{table_html}</div>")
                        day_pointer += 1
                    else:
                        st.markdown("<div style='height: 467px;'></div>", unsafe_allow_html=True)

    with view_mode[2]:
        if 'current_day_val' not in st.session_state: st.session_state.current_day_val = 1
        d_col1, d_col2, d_col3 = st.columns([1, 3, 1])
        with d_col1:
            if st.button("← 前の日", use_container_width=True, key="p_d_b"):
                if st.session_state.current_day_val > 1: st.session_state.current_day_val -= 1
        with d_col3:
            if st.button("次の日 →", use_container_width=True, key="n_d_b"):
                if st.session_state.current_day_val < max_days: st.session_state.current_day_val += 1
        with d_col2:
            st.session_state.current_day_val = st.slider("日選択スライダー", min_value=1, max_value=max_days, value=st.session_state.current_day_val, key="d_sld")
        st.write("---")
        
        try:
            this_date = datetime.date(target_year, target_month, st.session_state.current_day_val)
            wd_str = ["月", "火", "水", "木", "金", "土", "日"][this_date.weekday()]
        except: wd_str = ""
            
        st.components.v1.html('<button onclick="parent.window.print();" style="width:100%; height:42px; background-color:#4CAF50; color:white; border:none; border-radius:4px; cursor:pointer; font-weight:bold;">🖨️ この日の詳細を印刷</button>', height=45)

        st.markdown(f"<h2 style='text-align: center; color: #1c83e1;'>🔍 {target_month}月 {st.session_state.current_day_val}日 ({wd_str}) の詳細</h2>", unsafe_allow_html=True)
        day_table_html = make_html_table_with_time(calendar_data[st.session_state.current_day_val], font_size="15px", padding="6px", is_large=True)
        st.html(f"<div class='print-target' style='max-width: 650px; margin: 0 auto; border: 2px solid #1c83e1; background-color: #ffffff; border-radius: 8px; padding: 10px;'>{day_table_html}</div>")
else:
    st.info("💡 上部のメニューから、解析元のシフトExcelファイルをアップロードしてください。「週間予定表」と「勤務表」両方を入れると合算されて表示されます。")
