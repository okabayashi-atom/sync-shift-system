import streamlit as st
import streamlit.components.v1 as components
import openpyxl
import datetime
import re
import os

# システム環境設定（不具合防止）
os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

st.set_page_config(page_title="1週間シフト印刷システム", layout="wide")

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

# ────────────────────────────────────────────────────────
# 🛠️ 印刷専用CSS（画面上は非表示、印刷時のみ全体を覆う親ウィンドウ定義）
# ────────────────────────────────────────────────────────
st.html("""
<script>
window.setupA3Print = function() {
    var oldStyle = document.getElementById('a3-print-style');
    if (oldStyle) oldStyle.remove();

    var style = document.createElement('style');
    style.id = 'a3-print-style';
    style.innerHTML = `
        @media print {
            * {
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }
            /* 不要なUIをすべて強制非表示 */
            header, footer, div[data-testid="stSidebar"], div[data-testid="stHeader"], 
            .stButton, div[data-testid="stSelectbox"], .no-print, iframe, button {
                display: none !important;
            }
            div[data-testid="stMainBlockContainer"] {
                max-width: 100% !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            div[data-testid="stAppViewContainer"] > div {
                display: block !important;
            }
            /* シフト表エリアだけを画面いっぱいに拡大 */
            #print-area-target {
                display: block !important;
                width: 100% !important;
                position: absolute;
                left: 0;
                top: 0;
                background: white;
                z-index: 9999999;
            }
            @page {
                size: A3 landscape;
                margin: 6mm 5mm;
            }
            .calendar-table {
                page-break-inside: avoid !important;
                border-collapse: collapse !important;
                width: 100% !important;
            }
            .calendar-table td {
                padding: 2px 1px !important;
                line-height: 1.1 !important;
                height: 20px !important;
                font-size: 13px !important;
            }
            .staff-name-box {
                font-size: 13px !important;
            }
        }
    `;
    document.head.appendChild(style);
};
</script>
""")

st.markdown("""
    <style>
    div[data-testid="stMainBlockContainer"] { max-width: 96% !important; padding: 1rem 1.5rem !important; }
    div[data-testid="stSelectbox"] label { display: none !important; }
    div[data-testid="stSelectbox"] { margin-top: 0px !important; }
    .calendar-table { table-layout: fixed !important; width: 100% !important; border-collapse: collapse !important; }
    .calendar-table td { container-type: inline-size !important; vertical-align: middle !important; padding: 4px 1px !important; height: 18px !important; }
    .staff-name-box { display: block !important; white-space: nowrap !important; overflow: visible !important; font-size: min(12px, 25cqw) !important; text-align: center; line-height: 1.2 !important; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h2 style='text-align: center; color: #333; margin-bottom:10px;'>📅 週間シフト配置表（印刷専用版）</h2>", unsafe_allow_html=True)

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

if uploaded_file is not None:
    st.write("---")
    
    # 選択コントロール部
    if 'current_week_idx' not in st.session_state: st.session_state.current_week_idx = 0
    weeks_list = ["第1週 (1日〜)", "第2週", "第3週", "第4週", "第5週", "第6週"]
    
    b_col1, b_col2, b_col3 = st.columns([1, 3, 1])
    with b_col1:
        if st.button("← 前の週", use_container_width=True):
            if st.session_state.current_week_idx > 0: st.session_state.current_week_idx -= 1
    with b_col3:
        if st.button("次の週 →", use_container_width=True):
            if st.session_state.current_week_idx < len(weeks_list) - 1: st.session_state.current_week_idx += 1
    with b_col2:
        week_option = st.selectbox("週選択", options=weeks_list, index=st.session_state.current_week_idx)
        st.session_state.current_week_idx = weeks_list.index(week_option)

    st.write("---")
    
    # 💡 完全独立型：A3印刷実行ボタン（iframeの壁を突破して親ウィンドウで直接print）
    components.html("""
        <button onclick="parent.setupA3Print(); parent.window.print();" style="width:100%; height:55px; background-color:#e67e22; color:white; border:none; border-radius:6px; cursor:pointer; font-size:16px; font-weight:bold; box-shadow:0 3px 6px rgba(0,0,0,0.2);">🖨️ 選択中の「週」を【A3用紙・横向き】で印刷する / PDF保存</button>
    """, height=65)
    
    start_d = st.session_state.current_week_idx * 7 + 1 - start_offset
    weekdays_labels = ["日", "月", "火", "水", "木", "金", "土"]
    
    # HTMLテーブル組み立て
    html_sheet = []
    html_sheet.append("<div id='print-area-target'>")
    html_sheet.append("<table class='calendar-table' style='width:100%; border-collapse:collapse; text-align:center; font-size:11px; font-family:sans-serif; table-layout:fixed; border:2px solid #333;'>")
    html_sheet.append("<tr style='background-color: #f0f0f0; font-weight: bold; border-bottom: 2px solid #ccc;'>")
    html_sheet.append("<td style='border: 1px solid #ccc; width: 3%; padding: 8px 0;'>時間</td>")
    
    for day_of_week in range(7):
        current_d = start_d + day_of_week
        color = '#ff4b4b' if day_of_week == 0 else '#1c83e1' if day_of_week == 6 else '#333333'
        if 1 <= current_d <= max_days:
            html_sheet.append(f"<td colspan='6' style='border: 1px solid #ccc; color: {color}; font-size: 14px; width: 13.8%;'>{weekdays_labels[day_of_week]} ({current_d}日)</td>")
        else:
            html_sheet.append(f"<td colspan='6' style='border: 1px solid #ccc; color: #aaa; background-color: #fafafa; width: 13.8%;'>{weekdays_labels[day_of_week]} (外)</td>")
    html_sheet.append("</tr>")
    
    html_sheet.append("<tr style='background-color: #f9f9f9; font-size: 9px; border-bottom: 1px solid #ccc;'><td style='border: 1px solid #ccc;'>-</td>")
    for _ in range(7):
        html_sheet.append("<td style='border: 1px solid #eee; width: 1.4%;'>サ1</td><td style='border: 1px solid #ccc; font-weight: bold; width: 2.1%;'>へ1</td>")
        html_sheet.append("<td style='border: 1px solid #eee; width: 1.4%;'>サ2</td><td style='border: 1px solid #ccc; font-weight: bold; width: 2.1%;'>へ2</td>")
        html_sheet.append("<td style='border: 1px solid #eee; width: 1.4%;'>サ3</td><td style='border: 1px solid #ccc; font-weight: bold; width: 2.1%;'>へ3</td>")
    html_sheet.append("</tr>")
    
    rows_list = []
    for hour in hours_sequence:
        rows_list.append((f"{hour}:00", True))
        rows_list.append(("0:30" if hour == 0 else f"{hour}:30", False))
        
    for idx, (time_str, is_even) in enumerate(rows_list):
        bbs = "border-bottom:1px solid #333;"
        if idx == len(rows_list) - 1: bbs = ""
        html_sheet.append(f"<tr style='{bbs} height: 18px;'><td style='border-right:1px solid #333; border-left:1px solid #ccc; background-color:#f2f2f2; font-weight:bold; padding:2px 0;'>{time_str}</td>")
        
        for day_of_week in range(7):
            current_d = start_d + day_of_week
            if 1 <= current_d <= max_days:
                ds = calendar_data[current_d]
                h_num = int(time_str.split(":")[0])
                rd = ds[h_num]["row1" if time_str.endswith(":00") else "row2"]
                for sk, hk in [("s1", "h1"), ("s2", "h2"), ("s3", "h3")]:
                    html_sheet.append(f"<td style='border-left:1px solid #ccc; border-right:1px solid #ccc; background-color:#fff; padding: 0px !important;'>{wrap_name(rd[sk], '')}</td>")
                    html_sheet.append(f"<td style='border-left:1px solid #ccc; border-right:1px solid #333; font-weight:bold; background-color:{get_bg(rd[hk], hk)}; padding: 0px !important; line-height: 1.0 !important;'>{wrap_name(rd[hk], hk)}</td>")
            else:
                html_sheet.append("<td colspan='6' style='border: 1px solid #eee; background-color: #fdfdfd; opacity:0.1;'></td>")
        html_sheet.append("</tr>")
        
    html_sheet.append("</table></div>")
    st.html(f"<div style='border: 2px solid #999; background-color: #ffffff; border-radius: 6px; padding: 5px; margin-top:10px;'>{''.join(html_sheet)}</div>")
