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
# 📌 曜日・個別カラー対応 サービス固定シフト反映関数
# ────────────────────────────────────────────────────────
def apply_fixed_service_schedule(calendar_data, target_year, target_month):
    weekday_map = {0: "月", 1: "火", 2: "水", 3: "木", 4: "金", 5: "土", 6: "日"}

    schedules = [
        # ーーー 共通 ーーー
        (7, 0, "越智f", "s1", "全", None),
        (7, 30, "八子mh", "s1", "全", "#ff64c8"),
        (8, 0, "木村fh", "s1", "全", "#ff64c8"),
        (12, 0, "越智h", "s1", "全", "#ffff00"),
        (12, 30, "八子h", "s1", "全", "#ff64c8"),
        (15, 0, "貝森f", "s1", "全", "#ffff00"),
        (17, 0, "越智f", "s1", "全", None),
        (17, 30, "貝森c", "s1", "全", "#ffff00"),
        (18, 30, "西井eh", "s1", "全", "#ff64c8"),
        (19, 0, "照井e", "s1", "全", "#ff64c8"),
        (19, 30, "八子ef", "s1", "全", "#ff64c8"),
        (20, 0, "越智e", "s1", "全", "#ffff00"),
        (20, 30, "貝森f", "s1", "全", "#ffff00"),
        (21, 0, "平野e", "s1", "全", "#ffff00"),
        
        (17, 0, "木村", "s2", "全", None),
        
        (5, 0, "越智m", "s3", "全", "#ffff00"),
        (5, 30, "貝森m", "s3", "全", "#ffff00"),
        (6, 0, "照井mf", "s3", "全", "#ff64c8"),
        (7, 0, "佐藤m", "s3", "全", "#ff64c8"),
        (7, 30, "平野m", "s3", "全", "#ffff00"),
        (8, 0, "貝森c", "s3", "全", "#ffff00"),
        (18, 30, "佐藤e", "s3", "全", "#ff64c8"),
        
        # ーーー サービス1 (s1) ーーー
        # --- 月 ---
        (10, 0, "越智h", "s1", "月", None),
        (11, 0, "貝森b", "s1", "月", "#ffff00"),
        (11, 30, "〃", "s1", "月", "#ffff00"),
        (13, 0, "西井S", "s1", "月", None),
        (13, 30, "〃", "s1", "月", None),
        (23, 0, "貝森h", "s1", "月", "#ffff00"),
        # --- 火 ---
        (10, 0, "赤尾b", "s1", "火", None),
        (10, 30, "〃", "s1", "火", None),
        (11, 0, "越智b", "s1", "火", None),
        (11, 30, "〃", "s1", "火", None),
        (14, 0, "越智sw", "s1", "火", "#ffff00"),
        (14, 30, "〃", "s1", "火", "#ffff00"),
        # --- 水 ---
        (9, 0, "山中b", "s1", "水", None),
        (9, 30, "〃", "s1", "水", None),
        (10, 0, "越智h", "s1", "水", None),
        (13, 0, "斎藤", "s1", "水", None),
        (13, 30, "〃", "s1", "水", None),
        (14, 0, "西井b", "s1", "水", None),
        (14, 30, "〃", "s1", "水", None),
        (23, 0, "貝森h", "s1", "水", "#ffff00"),
        # --- 木 ---
        (10, 0, "脇坂sw", "s1", "木", None),
        (10, 30, "〃", "s1", "木", None),
        (11, 0, "貝森b", "s1", "木", "#ffff00"),
        (11, 30, "〃", "s1", "木", "#ffff00"),
        (13, 0, "邦洋sw", "s1", "木", None),
        (13, 30, "〃", "s1", "木", None),
        (23, 0, "貝森h", "s1", "木", "#ffff00"),
        # --- 金 ---
        (10, 0, "赤尾b", "s1", "金", None),
        (10, 30, "〃", "s1", "金", None),
        (11, 0, "越智b", "s1", "金", None),
        (11, 30, "〃", "s1", "金", None),
        (14, 0, "越智sw", "s1", "金", "#ffff00"),
        (14, 30, "〃", "s1", "金", "#ffff00"),
        # --- 土 ---
        (9, 0, "照井b", "s1", "土", None),
        (9, 30, "〃", "s1", "土", None),
        (10, 0, "越智h", "s1", "土", None),
        (13, 0, "斎藤b", "s1", "土", None),
        (13, 30, "〃", "s1", "土", None),
        (14, 0, "上吉川sw", "s1", "土", "#ffff00"),
        (14, 30, "〃", "s1", "土", "#ffff00"),
        (23, 0, "貝森h", "s1", "土", "#ffff00"),
        # --- 日 ---
        (10, 0, "越智h", "s1", "日", None),
        (13, 0, "邦洋b", "s1", "日", None),
        (13, 30, "〃", "s1", "日", None),
        (14, 0, "御前sw", "s1", "日", None),
        (14, 30, "〃", "s1", "日", None),
        
        # ーーー サービス2 (s2) ーーー
        # --- 月 ---
        (9, 0, "脇坂sw", "s2", "月", None),
        (9, 30, "〃", "s2", "月", None),
        (10, 0, "貝森h", "s2", "月", "#ffff00"),
        (11, 30, "木村fh", "s2", "月", "#ff64c8"),
        (12, 30, "貝森c", "s2", "月", "#ffff00"),
        (13, 0, "赤尾sw", "s2", "月", None),
        (13, 30, "〃", "s2", "月", None),
        (14, 0, "木村b", "s2", "月", None),
        (15, 0, "脇坂", "s2", "月", None),
        (15, 30, "〃", "s2", "月", None),
        (16, 0, "貝森sw", "s2", "月", "#ffff00"),
        (16, 30, "〃", "s2", "月", "#ffff00"),
        # --- 火 ---
        (14, 0, "山中sw", "s2", "火", None),
        (14, 30, "〃", "s2", "火", None),
        (16, 0, "上吉川sw", "s2", "火", "#ffff00"),
        (16, 30, "〃", "s2", "火", "#ffff00"),
        # --- 水 ---
        (14, 0, "棟方s", "s2", "水", None),
        (14, 30, "〃", "s2", "水", None),
        (16, 0, "斎藤sw", "s2", "水", None),
        (16, 30, "〃", "s2", "水", None),
        # --- 木 ---
        (9, 30, "越智h", "s2", "木", None),
        (10, 0, "貝森h", "s2", "木", "#ffff00"),
        (11, 30, "木村fh", "s2", "木", "#ff64c8"),
        (12, 30, "貝森c", "s2", "木", "#ffff00"),
        (13, 0, "石田sw", "s2", "木", None),
        (13, 30, "〃", "s2", "木", None),
        (14, 0, "御前sw", "s2", "木", None),
        (14, 30, "〃", "s2", "木", None),
        # --- 金 ---
        (14, 0, "棟方s", "s2", "金", None),
        (14, 30, "〃", "s2", "金", None),
        (15, 0, "〃", "s2", "金", None),
        (16, 0, "菅原s", "s2", "金", "#ffff00"),
        (16, 30, "〃", "s2", "金", "#ffff00"),
        # --- 土 ---
        (16, 0, "斎藤sw", "s2", "土", None),
        (16, 30, "〃", "s2", "土", None),
        # --- 日 ---
        (16, 0, "菅原s", "s2", "日", "#ffff00"),
        (16, 30, "〃", "s2", "日", "#ffff00"),
        
        # ーーー サービス3 (s3) ーーー
        # --- 月 ---
        (16, 30, "石田w", "s3", "月", "#ff64c8"), 
        (17, 0, "平野sw", "s3", "月", "#ffff00"),
        (17, 30, "〃", "s3", "月", "#ffff00"),
        (18, 0, "上吉川w", "s3", "月", "#ffff00"),
        # --- 火 ---
        (10, 0, "貝森h", "s3", "火", "#ffff00"),
        (10, 30, "八子sw", "s3", "火", None),
        (11, 0, "〃", "s3", "火", None),
        (11, 30, "木村fh", "s3", "火", "#ff64c8"),
        (12, 0, "貝森c", "s3", "火", "#ffff00"),
        (14, 0, "照井b", "s3", "火", None),
        (14, 30, "〃", "s3", "火", None),
        (15, 0, "八子b", "s3", "火", None),
        (15, 30, "〃", "s3", "火", None),
        (16, 0, "樺澤sw", "s3", "火", "#ffff00"),
        (16, 30, "〃", "s3", "火", "#ffff00"),
        (17, 0, "樺澤h", "s3", "火", "#ffff00"),
        (17, 30, "平野w", "s3", "火", "#ffff00"),
        # --- 水 ---
        (10, 0, "貝森h", "s3", "水", "#ffff00"),
        (10, 30, "石田b", "s3", "水", "#ffff00"),
        (12, 0, "貝森c", "s3", "水", "#ffff00"),
        (13, 0, "照井sw", "s3", "水", "#ffff00"),
        (13, 30, "〃", "s3", "水", "#ffff00"),
        (17, 0, "平野sw", "s3", "水", "#ffff00"),
        (17, 30, "〃", "s3", "水", "#ffc0cb"),
        (18, 0, "上吉川w", "s3", "水", "#ffc0cb"),
        # --- 木 ---
        (16, 0, "貝森sw", "s3", "木", "#ffff00"),
        (16, 30, "〃", "s3", "木", "#ffff00"),
        (17, 30, "平野w", "s3", "木", "#ffff00"),
        (18, 0, "上吉川w", "s3", "木", "#ffff00"),
        # --- 金 ---
        (10, 0, "貝森h", "s3", "金", "#ffff00"),
        (10, 30, "八子sw", "s3", "金", None),
        (11, 0, "〃", "s3", "金", None),
        (11, 30, "木村fh", "s3", "金", "#ff64c8"),
        (12, 0, "貝森c", "s3", "金", "#ffff00"),
        (14, 0, "木村sw", "s3", "金", None),
        (14, 30, "〃", "s3", "金", None),
        (17, 30, "平野sw", "s3", "金", "#ffff00"),
        (18, 0, "上吉川w", "s3", "金", "#ffff00"),
        # --- 土 ---
        (10, 0, "貝森h", "s3", "土", "#ffff00"),
        (11, 30, "木村fh", "s3", "土", "#ff64c8"),
        (12, 0, "貝森c", "s3", "土", "#ffff00"),
        (14, 0, "木村sw", "s3", "土", None),
        (14, 30, "〃", "s3", "土", None),
        (16, 0, "斎藤sw", "s3", "土", "#ffff00"),
        (16, 30, "〃", "s3", "土", "#ffff00"),
        (17, 0, "平野sw", "s3", "土", "#ffff00"),
        (17, 30, "〃", "s3", "土", "#ffff00"),
        # --- 日 ---
        (10, 0, "貝森h", "s3", "日", "#ffff00"),
        (11, 30, "木村fh", "s3", "日", "#ff64c8"),
        (12, 0, "貝森c", "s3", "日", "#ffff00"),
        (14, 0, "木村sw", "s3", "日", None),
        (14, 30, "〃", "s3", "日", None),
        (17, 30, "平野sw", "s3", "日", "#ffff00"),
        (18, 0, "上吉川w", "s3", "日", "#ffff00"),
    ]

    for d in range(1, 32):
        try:
            this_date = datetime.date(target_year, target_month, d)
            w_str = weekday_map[this_date.weekday()]
        except:
            continue

        for h, m, name, col, w_cond, custom_color in schedules:
            is_match = False
            if w_cond == "全" or (isinstance(w_cond, list) and w_str in w_cond) or w_cond == w_str:
                is_match = True

            if is_match:
                row_key = "row1" if m == 0 else "row2"
                calendar_data[d][h][row_key][col] = name
                if custom_color:
                    calendar_data[d][h][row_key][f"{col}_color"] = custom_color


# ────────────────────────────────────────────────────────
# 📌 共通パーツ、印刷位置の調整（余白カット用スタイル）
# ────────────────────────────────────────────────────────
st.markdown("""
<style>
@media print {
    /* 印刷対象（print-target）以外の不要なUIや余白を完全に非表示にする */
    body, html, [data-testid="stHeader"], [data-testid="stSidebar"], .stTabs, button {
        visibility: hidden !important;
        height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
    }
    /* 印刷したいメイン部分だけを1ページ目の最上部に強制表示 */
    .print-target, .print-target * { 
        visibility: visible !important; 
    }
    .print-target {
        position: absolute !important;
        left: 0 !important;
        top: 0 !important;
        width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
        page-break-inside: avoid !important;
        page-break-after: avoid !important;
    }
}
</style>
""", unsafe_allow_html=True)

st.markdown("<h2 style='text-align: center; color: #333; margin-top: 0px;'>📅 シフト配置確認システム</h2>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("シフトExcelファイルを選択してください", type=["xlsx", "xlsm"])

calendar_data = {}
for d in range(1, 32):
    calendar_data[d] = {}
    for h in range(0, 25):
        calendar_data[d][h] = {
            "row1": {"s1": "", "h1": "", "s2": "", "h2": "", "s3": "", "h3": ""},
            "row2": {"s1": "", "h1": "", "s2": "", "h2": "", "s3": "", "h3": ""}
        }

target_year = 2026
target_month = 7

if uploaded_file is not None:
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet_names = wb.sheetnames
        
        selected_sheet_name = st.selectbox("確認するシート（月）を選択してください", sheet_names)
        ws = wb[selected_sheet_name]
        
        # タイトルから年月を取得
        title_val = str(ws.cell(row=2, column=4).value or "")
        match_y = re.search(r'(令和\s*\d+|R\d+|\d+)\s*年', title_val)
        match_m = re.search(r'(\d+)\s*月', title_val)
        
        if match_y:
            y_str = match_y.group(1)
            if "令和" in y_str:
                num = int(re.sub(r'\D', '', y_str))
                target_year = 2018 + num
            elif "R" in y_str.upper():
                num = int(re.sub(r'\D', '', y_str.upper()))
                target_year = 2018 + num
            else:
                target_year = int(re.sub(r'\D', '', y_str))
        if match_m:
            target_month = int(match_m.group(1))
            
        st.info(f"📅 解析対象：{target_year}年{target_month}月")
        
        # 本文解析
        day_cols = {}
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=5, column=c).value or "").strip()
            if "日" in val:
                d_num = int(re.sub(r'\D', '', val))
                day_cols[d_num] = c

        for r in range(8, ws.max_row + 1):
            time_val = str(ws.cell(row=r, column=2).value or "").strip()
            if not time_val or time_val == "None":
                continue
            
            for d_num, col_idx in day_cols.items():
                staff_name = str(ws.cell(row=r, column=col_idx).value or "").strip()
                if not staff_name or staff_name == "None" or staff_name == "〃":
                    continue
                
                # 時間判定
                try:
                    h_part, m_part, s_part = map(int, time_val.split(':'))
                except:
                    continue
                
                row_key = "row1" if m_part == 0 else "row2"
                
                is_haya = "ｍ" in staff_name or "m" in staff_name
                is_nichi = "ｃ" in staff_name or "c" in staff_name
                is_yoru = "ｅ" in staff_name or "e" in staff_name or "ｆ" in staff_name or "f" in staff_name or "h" in staff_name or "ｈ" in staff_name
                
                assigned_h = None
                start_hour, start_min = None, None
                end_hour, end_min = None, None
                
                if is_haya:
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
                    end_hour, end_min = 21, 0
                    
                if assigned_h and start_hour is not None and end_hour is not None:
                    current_time = datetime.datetime(2026, 1, 1, start_hour, start_min)
                    end_time = datetime.datetime(2026, 1, 1, end_hour, end_min)
                    if current_time >= end_time:
                        continue
                    
                    time_slots = []
                    while current_time <= end_time:
                        time_slots.append((current_time.hour, current_time.minute))
                        current_time += datetime.timedelta(minutes=30)
                        
                    for idx, (h, m) in enumerate(time_slots):
                        r_key = "row1" if m == 0 else "row2"
                        if idx == 0 or idx == len(time_slots) - 1:
                            calendar_data[d_num][h][r_key][assigned_h] = staff_name
                        else:
                            current_val = calendar_data[d_num][h][r_key][assigned_h]
                            if not current_val or current_val == "┃":
                                calendar_data[d_num][h][r_key][assigned_h] = "┃"

        apply_fixed_service_schedule(calendar_data, target_year, target_month)
        st.success(f"🎉 シート「{selected_sheet_name}」のデータを正常に解析しました。")
        
    except Exception as e:
        st.error(f"Excelの読み込みエラー: {e}")

try:
    start_offset = (datetime.date(target_year, target_month, 1).weekday() + 1) % 7
except:
    start_offset = 0

if target_month in [4, 6, 9, 11]:
    max_days = 30
elif target_month == 2:
    max_days = 28
else:
    max_days = 31


def get_bg(val, h_type):
    if not val or val in ["┃", "None", ""]: return "#ffffff"
    if h_type == "h1": return "#e2f0d9"
    if h_type == "h2": return "#fff2cc"
    if h_type == "h3": return "#fce4d6"
    return "#ffffff"

def wrap_name(val, h_type):
    if not val or val == "None": return ""
    if val == "┃": return "┃"
    name_only = re.sub(r'[a-zA-Zｍ㎡ｃ経早日夜遅ｆｈｅｓｗＳＷｂ]', '', val).strip()
    if h_type in ["h1", "h2", "h3"]:
        suffix = "M" if "m" in val.lower() or "ｍ" in val else "C" if "c" in val.lower() or "ｃ" in val else "E" if any(x in val.lower() for x in ["e","f","h","ｅ","ｆ","ｈ"]) else ""
        return f"{name_only}<br><span style='font-size:9px;color:#666;'>{suffix}</span>" if suffix else name_only
    return name_only


# ────────────────────────────────────────────────────────
# 📌 週間予定表作成関数
# ────────────────────────────────────────────────────────
def generate_weekly_html_table(start_d, max_days, target_year, target_month, calendar_data):
    td_p_style = "padding: 3px 2px;"
    line_h_style = "line-height: 1.1;"
    
    hours_sequence = list(range(5, 24)) + [0]
    weekdays_labels = ["日", "月", "火", "水", "木", "金", "土"]
    h_sheet = []
    h_sheet.append("<div class='print-target'>")
    h_sheet.append("<table class='calendar-table' style='width:100%; border-collapse:collapse; text-align:center; font-family:sans-serif; table-layout:fixed; border:2px solid #333;'>")
    h_sheet.append("<tr style='background-color: #f0f0f0; font-weight: bold;'>")
    h_sheet.append("<td style='width: 5%; padding: 6px 0;'>時間</td>")
    
    for d_o_w in range(7):
        cur_d = start_d + d_o_w
        c_color = '#ff4b4b' if d_o_w == 0 else '#1c83e1' if d_o_w == 6 else '#333333'
        if 1 <= cur_d <= max_days:
            h_sheet.append(f"<td colspan='6' style='color: {c_color}; font-size: 13px;'>{weekdays_labels[d_o_w]} ({cur_d}日)</td>")
        else:
            h_sheet.append(f"<td colspan='6' style='color: #ccc; font-size: 13px;'>{weekdays_labels[d_o_w]} (-)</td>")
    h_sheet.append("</tr>")
    
    h_sheet.append("<tr style='background-color: #f9f9f9; font-size: 11px; font-weight: bold;'>")
    h_sheet.append("<td style='border-right: 1px solid #333;'></td>")
    for _ in range(7):
        h_sheet.append("<td colspan='2' style='border-right: 1px solid #ccc; color:#2e7d32;'>サ1</td>")
        h_sheet.append("<td colspan='2' style='border-right: 1px solid #ccc; color:#c62828;'>サ2</td>")
        h_sheet.append("<td colspan='2' style='border-right: 1px solid #333; color:#1565c0;'>サ3</td>")
    h_sheet.append("</tr>")
    
    for hour in hours_sequence:
        for r_key, label_ext in [("row1", ":00"), ("row2", ":30")]:
            t_label = "0:30" if hour == 0 and label_ext == ":30" else f"{hour}{label_ext}"
            
            is_dash = label_ext == ":00"
            border_bottom_style = "border-bottom: 1px dashed #ddd;" if is_dash else "border-bottom: 1px solid #333;"
            
            h_sheet.append(f"<tr style='{border_bottom_style} height: 24px;'>")
            h_sheet.append(f"<td style='border-right: 1px solid #333; border-left: 1px solid #ccc; background-color: #f9f9f9; font-weight: bold; font-size: 11px;'>{t_label}</td>")
            
            for d_o_w in range(7):
                cur_d = start_d + d_o_w
                day_border_right = "border-right: 1px solid #ccc;" if d_o_w < 6 else "border-right: 1px solid #333;"
                
                if 1 <= cur_d <= max_days:
                    row_data = calendar_data[cur_d][hour][r_key]
                    
                    for s_key, h_key in [("s1", "h1"), ("s2", "h2"), ("s3", "h3")]:
                        val_s = row_data[s_key]
                        val_h = row_data[h_key]
                        bg_color_s = row_data.get(f"{s_key}_color", "#ffffff")
                        bg_color_h = get_bg(val_h, h_key)
                        
                        h_border_right = "border-right: 1px solid #333;" if s_key == "s3" else "border-right: 1px solid #ccc;"
                        if s_key == "s3":
                            h_border_right = day_border_right
                        
                        h_sheet.append(f"<td style='border-left: 1px solid #ccc; border-right: 1px solid #ccc; background-color: {bg_color_s}; {td_p_style} {line_h_style} font-size:11px;'>{wrap_name(val_s, '')}</td>")
                        h_sheet.append(f"<td style='border-left: 1px solid #ccc; {h_border_right} font-weight: bold; background-color: {bg_color_h}; {td_p_style} {line_h_style} font-size:11px;'>{wrap_name(val_h, h_key)}</td>")
                else:
                    for _ in range(3):
                        h_sheet.append(f"<td style='background-color: #fafafa; border-right: 1px solid #ccc;'></td>")
                        h_sheet.append(f"<td style='background-color: #fafafa; {day_border_right if _==2 else \"border-right: 1px solid #ccc;\"}'></td>")
            h_sheet.append("</tr>")
            
    h_sheet.append("</table>")
    h_sheet.append("</div>")
    return "".join(h_sheet)


# ────────────────────────────────────────────────────────
# 📌 メイン表示ロジック
# ────────────────────────────────────────────────────────
if uploaded_file is not None:
    view_mode = st.tabs(["📊 1ヶ月表示（カレンダー）", "📅 1週間表示", "🔍 1日集中表示"])

    # ────────────────────────────────────────────────────────
    # タブ1：1ヶ月表示
    # ────────────────────────────────────────────────────────
    with view_mode[0]:
        st.components.v1.html('<button onclick="parent.window.print();" style="width:100%; height:42px; background-color:#1c83e1; color:white; border:none; border-radius:4px; cursor:pointer; font-weight:bold;">🖨️ 1ヶ月分印刷プレビュー</button>', height=45)
        
        weekdays = ["日", "月", "火", "水", "木", "金", "土"]
        html = []
        html.append("<div class='print-target'>")
        html.append(f"<h3 style='text-align:center;margin-bottom:10px;'>{target_year}年 {target_month}月 シフト配置一覧</h3>")
        html.append("<table style='width:100%; border-collapse:collapse; table-layout:fixed; font-family:sans-serif;'>")
        html.append("<tr style='background-color:#f0f0f0;font-weight:bold;text-align:center;'>")
        for w in weekdays:
            c = '#ff4b4b' if w=='日' else '#1c83e1' if w=='土' else '#333'
            html.append(f"<th style='border:1px solid #ccc; padding:6px; color:{c}; width:14.28%;'>{w}</th>")
        html.append("</tr>")
        
        day_cursor = 1 - start_offset
        for _ in range(6):
            if day_cursor > max_days: break
            html.append("<tr style='height:110px; valign:top;'>")
            for d_o_w in range(7):
                if 1 <= day_cursor <= max_days:
                    t_dt = datetime.date(target_year, target_month, day_cursor)
                    w_name = weekdays[t_dt.weekday()]
                    bg_cell = "#f9f9f9" if w_name=='日' else "#f0f8ff" if w_name=='土' else "#ffffff"
                    
                    html.append(f"<td style='border:1px solid #ccc; background-color:{bg_cell}; padding:3px; vertical-align:top;'>")
                    html.append(f"<div style='font-weight:bold; margin-bottom:4px; font-size:13px;'>{day_cursor}</div>")
                    
                    day_sched = calendar_data[day_cursor]
                    active_staffs = set()
                    for hour in range(0, 25):
                        for rk in ["row1", "row2"]:
                            for s_k in ["s1", "s2", "s3"]:
                                v = day_sched[hour][rk][s_k]
                                if v and v != "┃": active_staffs.add(re.sub(r'[a-zA-Zｍ㎡ｃ経早日夜遅ｆｈｅｓｗＳＷｂ]', '', v).strip())
                    
                    if active_staffs:
                        html.append(f"<div style='font-size:11px; color:#555; line-height:1.3;'>👥 {', '.join(sorted(list(active_staffs)))}</div>")
                    html.append("</td>")
                else:
                    html.append("<td style='border:1px solid #eee; background-color:#fafafa;'></td>")
                day_cursor += 1
            html.append("</tr>")
        html.append("</table>")
        html.append("</div>")
        st.markdown("".join(html), unsafe_allow_html=True)

    # ────────────────────────────────────────────────────────
    # タブ2：1週間表示
    # ────────────────────────────────────────────────────────
    with view_mode[1]:
        st.components.v1.html('<button onclick="parent.window.print();" style="width:100%; height:42px; background-color:#ff9800; color:white; border:none; border-radius:4px; cursor:pointer; font-weight:bold;">🖨️ 選択中の週を印刷プレビュー</button>', height=45)
        
        weeks_options = []
        d_run = 1 - start_offset
        for w_idx in range(1, 7):
            if d_run > max_days: break
            w_start = max(1, d_run)
            w_end = min(max_days, d_run + 6)
            weeks_options.append((w_idx, f"第{w_idx}週 ({w_start}日〜{w_end}日)", d_run))
            d_run += 7
            
        sel_w = st.selectbox("表示する週を選択してください", weeks_options, format_func=lambda x: x[1])
        
        st.write("---")
        st.markdown(f"<h3 style='text-align: center; margin-bottom: 15px;'>令和8年6月 週間予定表 ({sel_w[1]})</h3>", unsafe_allow_html=True)
        
        weekly_html = generate_weekly_html_table(sel_w[2], max_days, target_year, target_month, calendar_data)
        st.markdown(weekly_html, unsafe_allow_html=True)

    # ────────────────────────────────────────────────────────
    # タブ3：1日集中表示
    # ────────────────────────────────────────────────────────
    with view_mode[2]:
        if "current_day_val" not in st.session_state:
            st.session_state.current_day_val = 1
            
        d_col1, d_col2, d_col3 = st.columns([1, 4, 1])
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

        st.markdown(f"<h2 style='text-align: center; color: #333; margin-top: 10px;'>{target_year}年 {target_month}月 {st.session_state.current_day_val}日 ({wd_str}) 詳細タイムライン</h2>", unsafe_allow_html=True)
        
        day_schedule = calendar_data[st.session_state.current_day_val]
        
        single_html = []
        single_html.append("<div class='print-target'>")
        single_html.append("<table style='width:100%; border-collapse:collapse; text-align:center; font-family:sans-serif;'>")
        single_html.append("<tr style='background-color:#f0f0f0; font-weight:bold; height:32px;'>")
        single_html.append("<td style='border:1px solid #333; width:10%;'>時間</td>")
        single_html.append("<td style='border:1px solid #333; width:30%; background-color:#e2f0d9;'>サービス 1 (サ1 / ヘ1)</td>")
        single_html.append("<td style='border:1px solid #333; width:30%; background-color:#fff2cc;'>サービス 2 (サ2 / ヘ2)</td>")
        single_html.append("<td style='border:1px solid #333; width:30%; background-color:#fce4d6;'>サービス 3 (サ3 / ヘ3)</td>")
        single_html.append("</tr>")
        
        full_hours = list(range(5, 24)) + [0]
        for h in full_hours:
            for r_k, label_ext in [("row1", ":00"), ("row2", ":30")]:
                t_lbl = "0:30" if h == 0 and label_ext == ":30" else f"{h}{label_ext}"
                row_cells = day_schedule[h][r_k]
                
                single_html.append("<tr style='border-bottom:1px solid #ccc; height:24px;'>")
                single_html.append(f"<td style='border:1px solid #333; background-color:#fafafa; font-weight:bold;'>{t_lbl}</td>")
                
                for s_idx, h_idx in [("s1","h1"), ("s2","h2"), ("s3","h3")]:
                    v_s = row_cells[s_idx]
                    v_h = row_cells[h_idx]
                    bg_s = row_cells.get(f"{s_idx}_color", "#ffffff")
                    bg_h = get_bg(v_h, h_idx)
                    
                    single_html.append(f"<td style='border:1px solid #ccc; background-color:{bg_s}; text-align:left; padding-left:15px;'>{wrap_name(v_s,'')} <span style='float:right; font-weight:bold; background-color:{bg_h}; padding:0 8px; border-left:1px solid #eee;'>{wrap_name(v_h, h_idx)}</span></td>")
                single_html.append("</tr>")
                
        single_html.append("</table>")
        single_html.append("</div>")
        st.markdown("".join(single_html), unsafe_allow_html=True)
else:
    st.info("💡 左側のメニューから、解析元のシフトExcelファイルをアップロードしてください。")
