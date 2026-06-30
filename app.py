import streamlit as st
import openpyxl
import datetime
import re
import os
import json

# システム環境設定（不具合防止）
os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

st.set_page_config(page_title="シフト配置確認システム", layout="wide")

# ────────────────────────────────────────────────────────
# 📌 サービス固定シフト反映関数
# ────────────────────────────────────────────────────────
def apply_fixed_service_schedule(calendar_data):
    schedules = [
        # ーーー サービス1 (s1) ーーー
        (7, 0, "越智f", "s1"), (7, 30, "八子mh", "s1"), (8, 0, "木村fh", "s1"),
        (10, 0, "赤尾b", "s1"), (10, 30, "赤尾b", "s1"),
        (11, 0, "越智b", "s1"), (11, 30, "越智b", "s1"),
        (12, 0, "越智h", "s1"), (12, 30, "八子h", "s1"),
        (14, 0, "越智sw", "s1"), (14, 30, "越智sw", "s1"),
        (15, 0, "貝森f", "s1"),
        (17, 0, "越智f", "s1"), (17, 30, "貝森c", "s1"),
        (18, 30, "西井eh", "s1"), (19, 0, "照井e", "s1"), (19, 30, "八子ef", "s1"),
        (20, 0, "越智e", "s1"), (20, 30, "貝森f", "s1"), (21, 0, "平野e", "s1"),
        
        # ーーー サービス2 (s2) ーーー
        (14, 0, "山中sw", "s2"), (14, 30, "山中sw", "s2"),
        (16, 0, "上吉川sw", "s2"), (16, 30, "上吉川sw", "s2"),
        (17, 0, "木村", "s2"),

        # ーーー サービス3 (s3) ーーー
        (5, 0, "越智m", "s3"), (5, 30, "貝森m", "s3"), (6, 0, "照井mf", "s3"),
        (7, 0, "佐藤m", "s3"), (7, 30, "平野m", "s3"), (8, 0, "貝森c", "s3"),
        (10, 0, "貝森h", "s3"), (10, 30, "石田b", "s3"), (12, 0, "貝森c", "s3"),
        (13, 0, "照井sw", "s3"), (13, 30, "照井sw", "s3"),
        (17, 0, "平野sw", "s3"), (17, 30, "平野sw", "s3"),
        (18, 0, "上吉川w", "s3"), (18, 30, "佐藤e", "s3")
    ]
    
    for d in range(1, 32):
        for h, m, name, col in schedules:
            row_key = "row1" if m == 0 else "row2"
            calendar_data[d][h][row_key][col] = name
            
        hours_seq = list(range(5, 24)) + [0]
        for col_key in ["s1", "s2", "s3"]:
            last_name = None
            for hour in hours_seq:
                for r_key in ["row1", "row2"]:
                    current_name = calendar_data[d][hour][r_key][col_key]
                    if current_name:
                        if current_name == last_name:
                            calendar_data[d][hour][r_key][col_key] = "〃"
                        else:
                            last_name = current_name
                    else:
                        last_name = None

st.markdown("""
    <style>
    div[data-testid="stMainBlockContainer"] { max-width: 96% !important; padding: 1rem 1.5rem !important; }
    div[data-testid="stSelectbox"] label { display: none !important; }
    div[data-testid="stSelectbox"] { margin-top: 0px !important; padding-top: 0px !important; }
    div[data-testid="stSlider"] label { display: none !important; }
    div[data-testid="stSlider"] { margin-top: 0px !important; padding-top: 5px !important; }
    .stButton > button { height: 42px !important; min-height: 42px !important; font-weight: bold !important; }
    .calendar-table { table-layout: fixed !important; width: 100% !important; border-collapse: collapse !important; }
    .calendar-table td { container-type: inline-size !important; vertical-align: middle !important; padding: 4px 1px !important; height: 18px !important; }
    .staff-name-box { display: block !important; white-space: nowrap !important; overflow: visible !important; font-size: min(12px, 25cqw) !important; text-align: center; line-height: 1.2 !important; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h2 style='text-align: center; color: #333;'>📅 シフト配置確認システム</h2>", unsafe_allow_html=True)

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
target_month = 6

if uploaded_file is not None:
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        all_sheets = wb.sheetnames
        selected_sheet_name = st.selectbox("確認したいシフトのシートを選択してください 👇", options=all_sheets)
        ws_main = wb[selected_sheet_name]

        match = re.search(r'(\d+)月|\.(\d+)', selected_sheet_name)
        if match:
            extracted_month = match.group(1) or match.group(2)
            target_month = int(extracted_month)

        if target_month in [4, 6, 9, 11]: max_days = 30
        elif target_month == 2: max_days = 28
        else: max_days = 31

        for d_num in range(1, max_days + 1):
            target_col = 2 + (d_num - 1)
            if target_col > ws_main.max_column: break

            for r_idx in range(4, 15):
                cell_val = ws_main.cell(row=r_idx, column=target_col).value
                if cell_val:
                    cell_str = str(cell_val).strip()
                    if not cell_str or cell_str in ["┃", "│", "↓", "｜", "〃"]: continue
                    staff_name = ws_main.cell(row=r_idx, column=1).value
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
                        end_hour, end_min = 21, 0

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

        apply_fixed_service_schedule(calendar_data)
        st.success(f"🎉 シート「{selected_sheet_name}」のデータを正常に解析しました。")
    except Exception as e:
        st.error(f"Excelの読み込みエラー: {e}")

try: start_offset = (datetime.date(target_year, target_month, 1).weekday() + 1) % 7
except: start_offset = 0

if target_month in [4, 6, 9, 11]: max_days = 30
elif target_month == 2: max_days = 28
else: max_days = 31

hours_sequence = list(range(5, 24)) + [0]

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

    html.append(f"<table style='border-collapse: collapse; text-align: center; font-size: {font_size}; font-family: sans-serif; width: 100%; table-layout: fixed; border: 1px solid #333;'>")
    html.append("<tr style='background-color: #f0f0f0; font-weight: bold; border-bottom: 2px solid #333;'>")
    html.append(f"<td style='border: 1px solid #ccc; width: 22%; padding: {padding}; font-size: 10px;'>時間</td>")
    html.append("<td style='border: 1px solid #ccc; width: 13%;'>サ1</td><td style='border: 1px solid #ccc; width: 13%;'>へ1</td>")
    html.append("<td style='border: 1px solid #ccc; width: 13%;'>サ2</td><td style='border: 1px solid #ccc; width: 13%;'>へ2</td>")
    html.append("<td style='border: 1px solid #ccc; width: 13%;'>サ3</td><td style='border: 1px solid #ccc; width: 13%;'>へ3</td>")
    html.append("</tr>")
    
    rows_list = []
    for hour in hours_sequence:
        rows_list.append((f"{hour}:00", day_schedule[hour]["row1"]))
        time_label = "0:30" if hour == 0 else f"{hour}:30"
        rows_list.append((time_label, day_schedule[hour]["row2"]))
    
    for idx, (time_str, row_data) in enumerate(rows_list):
        is_dash = time_str.endswith(":00")
        border_bottom_style = "border-bottom: 1px dashed #ddd;" if is_dash else "border-bottom: 1px solid #333;"
        if idx == len(rows_list) - 1: border_bottom_style = ""
            
        html.append(f"<tr style='{border_bottom_style} height: 18px;'>")
        html.append(f"<td style='border-right: 1px solid #333; border-left: 1px solid #ccc; background-color: #f9f9f9; font-weight: bold; padding: 2px 0;'>{time_str}</td>")
        
        for s_key, h_key in [("s1", "h1"), ("s2", "h2"), ("s3", "h3")]:
            val_s = row_data[s_key]
            val_h = row_data[h_key]
            bg_color_h = get_bg(val_h, h_key)
            html.append(f"<td style='border-left: 1px solid #ccc; border-right: 1px solid #ccc; background-color: #ffffff; {td_p_style} {line_h_style}'>{wrap_name(val_s, '')}</td>")
            html.append(f"<td style='border-left: 1px solid #ccc; border-right: 1px solid #333; font-weight: bold; background-color: {bg_color_h}; {td_p_style} {line_h_style}'>{wrap_name(val_h, h_key)}</td>")
        html.append("</tr>")
    html.append("</table>")
    return "".join(html)

# 💡 ポップアップ自動印刷ウインドウを生成するJavaScriptコア関数
def launch_print_popup(html_body_content, page_size="A4", orientation="landscape", table_font_size="11px"):
    js_code = f"""
    <script>
    (function() {{
        var printWindow = window.open('', '_blank');
        if(!printWindow) {{
            alert('ポップアップブロックが作動しました。ブラウザの設定でポップアップを許可してください。');
            return;
        }}
        printWindow.document.write(`
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>印刷プレビュー</title>
                <style>
                    body {{ margin: 0; padding: 0; background: white; font-family: sans-serif; }}
                    * {{ -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }}
                    @page {{ size: {page_size} {orientation}; margin: 5mm; }}
                    table {{ width: 100%; border-collapse: collapse; text-align: center; table-layout: fixed; }}
                    td {{ padding: 2px 1px; vertical-align: middle; height: 19px; font-size: {table_font_size}; border: 1px solid #333; }}
                    .staff-name-box {{ display: block; white-space: nowrap; overflow: visible; text-align: center; font-size: {table_font_size}; line-height: 1.1; }}
                </style>
            </head>
            <body>
                {html_body_content}
                <script>
                    window.onload = function() {{
                        setTimeout(function() {{
                            window.print();
                        }}, 300);
                    }};
                <\/script>
            </body>
            </html>
        `);
        printWindow.document.close();
    }})();
    </script>
    """
    st.components.v1.html(js_code, height=0, width=0)

if uploaded_file is not None:
    view_mode = st.tabs(["📊 1ヶ月表示（カレンダー）", "📅 1週間表示（時間軸スリム）", "🔍 1日集中表示"])

    # ────────────────────────────────────────────────────────
    # タブ1：1ヶ月表示（カレンダー形式）
    # ────────────────────────────────────────────────────────
    with view_mode[0]:
        if st.button("🖨️ 1ヶ月分カレンダーを印刷する（別タブでA4横開き）", key="cmd_print_month", use_container_width=True):
            # 印刷用HTMLの組み立て
            m_html = [f"<h2 style='text-align:center;'>{target_month}月 シフト配置カレンダー</h2>"]
            m_html.append("<table style='width:100%; border-collapse:collapse; table-layout:fixed; border:none;'>")
            m_html.append("<tr style='background:#f0f0f0; font-weight:bold; height:30px;'>")
            for w_day in ["日", "月", "火", "水", "木", "金", "土"]:
                m_html.append(f"<td style='border:1px solid #333;'>{w_day}曜日</td>")
            m_html.append("</tr>")
            
            d_ptr = 1
            for wk in range(6):
                if d_ptr > max_days: break
                m_html.append("<tr style='vertical-align:top;'>")
                for d_of_w in range(7):
                    cell_idx = wk * 7 + d_of_w
                    if start_offset <= cell_idx and d_ptr <= max_days:
                        t_html = make_html_table_with_time(calendar_data[d_ptr], font_size="8px", padding="1px", is_large=False)
                        m_html.append(f"<td style='border:1px solid #333; padding:3px; background:#fff;'><strong>{d_ptr}日</strong><div style='margin-top:3px;'>{t_html}</div></td>")
                        d_ptr += 1
                    else:
                        m_html.append("<td style='border:1px solid #333; background:#fafafa; opacity:0.3;'></td>")
                m_html.append("</tr>")
            m_html.append("</table>")
            launch_print_popup("".join(m_html), page_size="A4", orientation="landscape", table_font_size="8px")

        # 画面上への描画
        weekdays = ["日", "月", "火", "水", "木", "金", "土"]
        header_cols = st.columns(7)
        for i, day in enumerate(weekdays):
            color = '#ff4b4b' if day == '日' else '#1c83e1' if day == '土' else '#333333'
            header_cols[i].markdown(f"<h3 style='text-align: center; color: {color}; margin: 0;'>{day}曜日</h3>", unsafe_allow_html=True)
        st.write("---")

        day_pointer = 1
        for week in range(6):
            if day_pointer > max_days: break
            row_cols = st.columns(7)
            for day_of_week in range(7):
                current_cell_idx = week * 7 + day_of_week
                with row_cols[day_of_week]:
                    if start_offset <= current_cell_idx and day_pointer <= max_days:
                        st.markdown(f"<h4 style='margin: 0 0 5px 0; height: 27px; line-height: 27px;'><b>📅 {day_pointer}日</b></h4>", unsafe_allow_html=True)
                        table_html = make_html_table_with_time(calendar_data[day_pointer], font_size="9px", padding="1px", is_large=False)
                        st.html(f"<div class='scrollable-container' style='border: 1px solid #999; height: 430px; overflow-y: auto; background-color: #ffffff; border-radius: 4px;'>{table_html}</div>")
                        day_pointer += 1
                    else:
                        st.markdown("<div style='height: 27px; margin: 0 0 5px 0;'></div>", unsafe_allow_html=True)
                        st.markdown("<div style='height: 430px; background-color: #fdfdfd; border-radius:4px; opacity:0.3; border: 1px dashed #ccc;'></div>", unsafe_allow_html=True)

    # ────────────────────────────────────────────────────────
    # タブ2：1週間表示（A3横大判印刷形式）
    # ────────────────────────────────────────────────────────
    with view_mode[1]:
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
        
        start_d = st.session_state.current_week_idx * 7 + 1 - start_offset
        weekdays_labels = ["日", "月", "火", "水", "木", "金", "土"]
        
        # 週間表示用のHTML組み立て共通ロジック
        def build_week_html_table():
            h_sheet = []
            h_sheet.append(f"<h2 style='text-align:center;'>📅 {target_month}月 週間シフト配置表 ({weeks_list[st.session_state.current_week_idx]})</h2>")
            h_sheet.append("<table style='width:100%; border-collapse:collapse; text-align:center; font-family:sans-serif; table-layout:fixed; border:2px solid #333;'>")
            h_sheet.append("<tr style='background-color: #f0f0f0; font-weight: bold;'>")
            h_sheet.append("<td style='width: 3.5%; padding: 6px 0;'>時間</td>")
            
            for d_o_w in range(7):
                cur_d = start_d + d_o_w
                c_color = '#ff4b4b' if d_o_w == 0 else '#1c83e1' if d_o_w == 6 else '#333333'
                if 1 <= cur_d <= max_days:
                    h_sheet.append(f"<td colspan='6' style='color: {c_color}; font-size: 13px; width: 13.7%;'>{weekdays_labels[d_o_w]} ({cur_d}日)</td>")
                else:
                    h_sheet.append(f"<td colspan='6' style='color: #aaa; background-color: #fafafa; width: 13.7%;'>{weekdays_labels[d_o_w]} (外)</td>")
            h_sheet.append("</tr>")
            
            h_sheet.append("<tr style='background-color: #f9f9f9; font-size: 9px;'><td style='font-weight:bold;'>-</td>")
            for _ in range(7):
                h_sheet.append("<td style='width:1.4%;'>サ1</td><td style='font-weight:bold; width:2.1%; background:#fff3cd;'>へ1</td>")
                h_sheet.append("<td style='width:1.4%;'>サ2</td><td style='font-weight:bold; width:2.1%; background:#fff3cd;'>へ2</td>")
                h_sheet.append("<td style='width:1.4%;'>サ3</td><td style='font-weight:bold; width:2.1%; background:#fff3cd;'>へ3</td>")
            h_sheet.append("</tr>")
            
            r_list = []
            for hour in hours_sequence:
                r_list.append((f"{hour}:00", True))
                r_list.append(("0:30" if hour == 0 else f"{hour}:30", False))
                
            for idx, (t_str, is_even) in enumerate(r_list):
                b_style = "border-bottom:1px solid #333;" if not is_even else "border-bottom:1px dashed #ccc;"
                h_sheet.append(f"<tr style='{b_style} height:19px;'><td style='background-color:#f2f2f2; font-weight:bold;'>{t_str}</td>")
                
                for d_o_w in range(7):
                    cur_d = start_d + d_o_w
                    if 1 <= cur_d <= max_days:
                        ds = calendar_data[cur_d]
                        h_num = int(t_str.split(":")[0])
                        rd = ds[h_num]["row1" if t_str.endswith(":00") else "row2"]
                        for sk, hk in [("s1", "h1"), ("s2", "h2"), ("s3", "h3")]:
                            h_sheet.append(f"<td style='background-color:#fff;'>{wrap_name(rd[sk], '')}</td>")
                            h_sheet.append(f"<td style='font-weight:bold; background-color:{get_bg(rd[hk], hk)};'>{wrap_name(rd[hk], hk)}</td>")
                    else:
                        h_sheet.append("<td colspan='6' style='background-color: #fafafa; opacity:0.1;'></td>")
                h_sheet.append("</tr>")
            h_sheet.append("</table>")
            return "".join(h_sheet)

        if st.button("🖨️ この週間シフト表を印刷する（別タブでA3横の大判起動）", key="cmd_print_week", use_container_width=True):
            week_html_content = build_week_html_table()
            launch_print_popup(week_html_content, page_size="A3", orientation="landscape", table_font_size="11px")

        # 画面用描画
        st.html(f"<div style='border: 2px solid #999; background-color: #ffffff; border-radius: 6px; padding: 5px;'>{build_week_html_table()}</div>")

    # ────────────────────────────────────────────────────────
    # タブ3：1日集中表示（詳細縦型形式）
    # ────────────────────────────────────────────────────────
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
            
        if st.button(f"🖨️ {st.session_state.current_day_val}日の詳細を印刷する（別タブでA4縦開き）", key="cmd_print_day", use_container_width=True):
            day_table_html = make_html_table_with_time(calendar_data[st.session_state.current_day_val], font_size="14px", padding="5px", is_large=True)
            day_content = f"""
            <h2 style='text-align: center; color: #1c83e1;'>🔍 {target_month}月 {st.session_state.current_day_val}日 ({wd_str}曜日) シフト配置詳細</h2>
            <div style='max-width: 680px; margin: 0 auto;'>{day_table_html}</div>
            """
            launch_print_popup(day_content, page_size="A4", orientation="portrait", table_font_size="14px")

        # 画面用描画
        st.markdown(f"<h2 style='text-align: center; color: #1c83e1;'>🔍 {target_month}月 {st.session_state.current_day_val}日 ({wd_str}曜日) の詳細配置</h2>", unsafe_allow_html=True)
        large_table_html = make_html_table_with_time(calendar_data[st.session_state.current_day_val], font_size="15px", padding="6px", is_large=True)
        st.html(f"<div style='max-width: 650px; margin: 0 auto; border: 2px solid #1c83e1; background-color: #ffffff; border-radius: 8px; padding: 10px;'>{large_table_html}</div>")
