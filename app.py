import streamlit as st
import openpyxl
import datetime
import re
import os
import io
from ftplib import FTP

# システム環境設定（不具合・文字化け防止）
os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

st.set_page_config(page_title="シフト配置カレンダー", layout="wide")

# ────────────────────────────────────────────────────────
# ⚙️ FTPサーバーの接続設定
# ────────────────────────────────────────────────────────
FTP_HOST = "sv17062.xserver.jp"
FTP_USER = "atom_test@test.atom-makoto.com"
FTP_PASS = "d7CWAQrw"
FTP_DIR = "/Flower_House/sync_shift_System"
OUTPUT_FILE_NAME = "index.html"

# 🌐 テストサーバー用の公開URL
PUBLIC_URL = "http://test.atom-makoto.com/Flower_House/sync_shift_System/"

# ────────────────────────────────────────────────────────
# 📌 サービス固定シフト反映関数（サ1〜サ3の完全マスターデータ）
# ────────────────────────────────────────────────────────
def apply_fixed_service_schedule(calendar_data):
    # (時, 分, 名前, 列)
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

        # ーーー サービス3 (s3) 新設！ ーーー
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
            
        # 💡 サービス列（s1〜s3）で、同じ名前が連続している場合は下側を省略マーク「〃」にする
        hours_seq = list(range(5, 24)) + [0]
        for col_key in ["s1", "s2", "s3"]:
            last_name = None
            for hour in hours_seq:
                for r_key in ["row1", "row2"]:
                    current_name = calendar_data[d][hour][r_key][col_key]
                    if current_name:
                        # 直前と同じ名前なら「〃」に置き換える
                        if current_name == last_name:
                            calendar_data[d][hour][r_key][col_key] = "〃"
                        else:
                            last_name = current_name
                    else:
                        last_name = None # 空白が挟まったらリセット

# ────────────────────────────────────────────────────────
# 🛠️ PDF印刷用のJavaScript関数とCSS
# ────────────────────────────────────────────────────────
st.html("""
<script>
function printTabSection(containerId, orientation) {
    var oldStyle = document.getElementById('dynamic-print-style');
    if (oldStyle) oldStyle.remove();

    var style = document.createElement('style');
    style.id = 'dynamic-print-style';
    style.innerHTML = `
        @media print {
            * {
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }
            header, footer, div[data-testid="stSidebar"], div[data-testid="stHeader"], 
            .stButton, div[data-testid="stSelectbox"], div[data-testid="stSlider"], .no-print, .pdf-control-panel {
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
            #\${containerId} {
                display: block !important;
                width: 100% !important;
                height: auto !important;
                overflow: visible !important;
                position: absolute;
                left: 0;
                top: 0;
                background: white;
                z-index: 9999999;
            }
            #\${containerId} .scrollable-container {
                height: auto !important;
                overflow: visible !important;
            }
            @page {
                size: A4 \${orientation};
                margin: 5mm;
            }
            .calendar-table {
                page-break-inside: avoid !important;
                border-collapse: collapse !important;
            }
            .calendar-table td {
                padding-top: 0px !important;
                padding-bottom: 0px !important;
                line-height: 1.0 !important;
                height: 18px !important;
            }
        }
    `;
    document.head.appendChild(style);
    
    setTimeout(function() {
        window.print();
    }, 250);
}
</script>
""")

st.markdown("""
    <style>
    div[data-testid="stMainBlockContainer"] {
        max-width: 96% !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
    }
    div[data-testid="stSelectbox"] label {
        display: none !important;
    }
    div[data-testid="stSelectbox"] {
        margin-top: 0px !important;
        padding-top: 0px !important;
    }
    div[data-testid="stSelectbox"] > div > div {
        height: 40px !important;
        min-height: 40px !important;
        display: flex !important;
        align-items: center !important;
    }
    div[data-testid="stSlider"] label {
        display: none !important;
    }
    div[data-testid="stSlider"] {
        margin-top: 0px !important;
        padding-top: 5px !important;
    }
    .stButton > button {
        height: 40px !important;
        min-height: 40px !important;
        line-height: 40px !important;
        padding-top: 0px !important;
        padding-bottom: 0px !important;
    }
    .calendar-table {
        table-layout: fixed !important;
        width: 100% !important;
        border-collapse: collapse !important;
    }
    .calendar-table td {
        container-type: inline-size !important; 
        vertical-align: middle !important;
        padding: 4px 1px !important;
        height: 18px !important;
    }
    .staff-name-box {
        display: block !important;
        white-space: nowrap !important;
        overflow: visible !important;
        font-size: min(12px, 25cqw) !important;
        text-align: center !important;
        line-height: 1.2 !important;
    }
    @media print {
        header, footer, div[data-testid="stSidebar"], .stButton, div[data-testid="stSlider"], div[data-testid="stSelectbox"], .no-print {
            display: none !important;
        }
        div[data-testid="stMainBlockContainer"] {
            max-width: 100% !important;
            padding: 0 !important;
        }
        .calendar-table {
            page-break-inside: avoid;
        }
    }
    .pdf-btn-blue {
        width: 100%; 
        height: 42px; 
        background-color: #1c83e1; 
        color: white; 
        border: none; 
        border-radius: 4px; 
        cursor: pointer; 
        font-size: 14px; 
        font-weight: bold;
        margin-bottom: 10px;
    }
    .pdf-btn-green {
        width: 100%; 
        height: 42px; 
        background-color: #4CAF50; 
        color: white; 
        border: none; 
        border-radius: 4px; 
        cursor: pointer; 
        font-size: 14px; 
        font-weight: bold;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; color: #333;'>📅 シフト配置確認カレンダー</h1>", unsafe_allow_html=True)

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

        if target_month in [4, 6, 9, 11]:
            max_days = 30
        elif target_month == 2:
            max_days = 28
        else:
            max_days = 31

        for d_num in range(1, max_days + 1):
            target_col = 2 + (d_num - 1)
            if target_col > ws_main.max_column:
                break

            for r_idx in range(4, 15):
                cell_val = ws_main.cell(row=r_idx, column=target_col).value
                if cell_val:
                    cell_str = str(cell_val).strip()
                    if not cell_str or cell_str in ["┃", "│", "↓", "｜", "〃"]:
                        continue
                    
                    staff_name = ws_main.cell(row=r_idx, column=1).value
                    if not staff_name:
                        continue
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
                        
                        if current_time >= end_time:
                            continue
                        
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

        # 🚀 固定配置の反映＆自動省略（サ1〜サ3対応）
        apply_fixed_service_schedule(calendar_data)

        st.success(f"🎉 シート「{selected_sheet_name}」のデータを読み込みました。全サービス固定配置が完了しています！")
    except Exception as e:
        st.error(f"Excelの読み込みエラー: {e}")

st.write("---")

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

hours_sequence = list(range(5, 24)) + [0]

def get_bg(val, h_type):
    if not val or val in ["┃", "｜", "↓", "〃"]: return "#ffffff"
    if h_type == "h3": return "#ffc0cb"
    if h_type == "h1": return "#ffff00"
    return "#ffffff"

# 💡 全角カッコを半角カッコに自動置換
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

def build_and_upload_static_html():
    html_src = f"""<!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{target_month}月 シフト配置予定表</title>
        <style>
            body {{ font-family: sans-serif; background-color: #f4f4f4; padding: 10px; margin: 0; }}
            .container {{ max-width: 800px; margin: 0 auto; background: #fff; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
            h1 {{ text-align: center; color: #333; font-size: 20px; }}
            .calendar-table {{ table-layout: fixed; width: 100%; border-collapse: collapse; text-align: center; font-size: 11px; }}
            .calendar-table td {{ border: 1px solid #ccc; padding: 4px 1px; vertical-align: middle; }}
            .staff-name-box {{ display: block; white-space: nowrap; text-align: center; font-weight: bold; font-size: 11px; }}
            .day-section {{ margin-bottom: 30px; border: 1px solid #999; padding: 10px; background: #fff; border-radius: 6px; page-break-inside: avoid; }}
            .print-btn {{ padding: 10px 20px; background: #4CAF50; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; font-size: 14px; width: 100%; }}
            @media print {{ .no-print {{ display: none !important; }} body {{ padding: 0; }} .container {{ box-shadow: none; padding: 0; max-width: 100%; }} }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="no-print" style="margin-bottom: 20px;">
                <button onclick="window.print()" class="print-btn">📄 このページをPDFで保存 / 印刷する</button>
            </div>
            <h1>📅 {target_month}月 週間予定表（シフト）</h1>
            <p style="text-align: center; color: #666; font-size: 12px;">最終更新日時: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
    """
    
    for d_pointer in range(1, max_days + 1):
        try:
            t_date = datetime.date(target_year, target_month, d_pointer)
            wd = ["月", "火", "水", "木", "金", "土", "日"][t_date.weekday()]
        except:
            wd = ""
        html_src += f"""
        <div class="day-section">
            <h3 style="margin-top:0; color:#1c83e1;">📅 {target_month}月 {d_pointer}日 ({wd}曜日)</h3>
            {make_html_table_with_time(calendar_data[d_pointer], font_size='11px', padding='3px')}
        </div>
        """
        
    html_src += """
        </div>
    </body>
    </html>
    """
    
    try:
        with FTP(FTP_HOST, FTP_USER, FTP_PASS) as ftp:
            ftp.cwd(FTP_DIR)
            bio = io.BytesIO(html_src.encode('utf-8'))
            ftp.storbinary(f"STOR {OUTPUT_FILE_NAME}", bio)
        return True, None
    except Exception as ex:
        return False, str(ex)

if uploaded_file is not None:
    c1, c2 = st.columns([2, 1])
    with c1:
        if st.button("🚀 この内容をWeb稼働表として公開保存する（確定送信）", use_container_width=True):
            with st.spinner("サーバーのファイルを更新中..."):
                success, err_msg = build_and_upload_static_html()
                if success:
                    st.success(f"🟢 保存完了！テスト公開ページへダイレクトに反映されました。")
                else:
                    st.error(f"🔴 保存に失敗しました。FTP設定内容を再度確認してください: {err_msg}")
    with c2:
        st.markdown('<button onclick="window.print()" style="width:100%; height:40px; background-color:#4CAF50; color:white; border:none; border-radius:4px; cursor:pointer; font-size:14px; font-weight:bold;">📄 この画面全体をPDF保存</button>', unsafe_allow_html=True)

view_mode = st.tabs(["📊 1ヶ月表示（カレンダー）", "📅 1週間表示（時間軸スリム化）", "🔍 1日集中表示", "🌐 公開ページ確認"])

# タブ1：1ヶ月表示
with view_mode[0]:
    st.html("""
    <button onclick="printTabSection('month-view-container', 'landscape')" class="pdf-btn-blue">
        📄 1ヶ月分カレンダー形式でPDF閲覧 / ダウンロードする（A4横）
    </button>
    """)

    st.html("<div id='month-view-container'>")

    weekdays = ["日", "月", "火", "水", "木", "金", "土"]
    header_cols = st.columns(7)
    for i, day in enumerate(weekdays):
        color = '#ff4b4b' if day == '日' else '#1c83e1' if day == '土' else '#333333'
        header_cols[i].markdown(f"<h3 style='text-align: center; color: {color}; margin: 0;'>{day}曜日</h3>", unsafe_allow_html=True)
    st.write("---")

    day_pointer = 1
    for week in range(6):
        if day_pointer > max_days:
            break
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
                        
    st.html("</div>")

# タブ2：1週間表示
with view_mode[1]:
    if 'current_week_idx' not in st.session_state:
        st.session_state.current_week_idx = 0
        
    weeks_list = ["第1週 (1日〜)", "第2週", "第3週", "第4週", "第5週", "第6週"]
    
    b_col1, b_col2, b_col3 = st.columns([1.2, 4, 1.2])
    with b_col1:
        if st.button("← 前の週", use_container_width=True, key="prev_week_btn"):
            if st.session_state.current_week_idx > 0:
                st.session_state.current_week_idx -= 1
    with b_col3:
        if st.button("次の週 →", use_container_width=True, key="next_week_btn"):
            if st.session_state.current_week_idx < len(weeks_list) - 1:
                st.session_state.current_week_idx += 1
    with b_col2:
        week_option = st.selectbox("週選択", options=weeks_list, index=st.session_state.current_week_idx, key="week_select_box")
        st.session_state.current_week_idx = weeks_list.index(week_option)

    st.write("---")
    
    st.html("""
    <button onclick="printTabSection('week-view-container', 'landscape')" class="pdf-btn-blue">
        📄 1週間横並び形式でPDF閲覧 / ダウンロードする（A4横）
    </button>
    """)
    
    start_d = st.session_state.current_week_idx * 7 + 1 - start_offset
    weekdays_labels = ["日", "月", "火", "水", "木", "金", "土"]
    
    html_sheet = []
    html_sheet.append("<div id='week-view-container' style='overflow-x:auto;'>")
    html_sheet.append("<table class='calendar-table' style='width:100%; border-collapse:collapse; text-align:center; font-size:11px; font-family:sans-serif; table-layout:fixed; border:2px solid #333;'>")
    
    html_sheet.append("<tr style='background-color: #f0f0f0; font-weight: bold; position: sticky; top: 0; z-index: 10; border-bottom: 2px solid #ccc;'>")
    html_sheet.append("<td style='border: 1px solid #ccc; width: 2.5%; padding: 8px 0; font-size: 11px;'>時間</td>")
    
    for day_of_week in range(7):
        current_d = start_d + day_of_week
        color = '#ff4b4b' if day_of_week == 0 else '#1c83e1' if day_of_week == 6 else '#333333'
        if 1 <= current_d <= max_days:
            html_sheet.append(f"<td colspan='6' style='border: 1px solid #ccc; color: {color}; font-size: 13px; width: 13.9%;'>{weekdays_labels[day_of_week]} ({current_d}日)</td>")
        else:
            html_sheet.append(f"<td colspan='6' style='border: 1px solid #ccc; color: #aaa; background-color: #fafafa; width: 13.9%;'>{weekdays_labels[day_of_week]} (外)</td>")
    html_sheet.append("</tr>")
    
    html_sheet.append("<tr style='background-color: #f9f9f9; font-size: 9px; border-bottom: 1px solid #ccc;'>")
    html_sheet.append("<td style='border: 1px solid #ccc;'>-</td>")
    for _ in range(7):
        html_sheet.append("<td style='border: 1px solid #eee; width: 1.4%;'>サ1</td><td style='border: 1px solid #ccc; font-weight: bold; width: 2.23%;'>へ1</td>")
        html_sheet.append("<td style='border: 1px solid #eee; width: 1.4%;'>サ2</td><td style='border: 1px solid #ccc; font-weight: bold; width: 2.23%;'>へ2</td>")
        html_sheet.append("<td style='border: 1px solid #eee; width: 1.4%;'>サ3</td><td style='border: 1px solid #ccc; font-weight: bold; width: 2.23%;'>へ3</td>")
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
                rk = "row1" if time_str.endswith(":00") else "row2"
                rd = ds[h_num][rk]
                
                for sk, hk in [("s1", "h1"), ("s2", "h2"), ("s3", "h3")]:
                    val_s = rd[sk]
                    val_h = rd[hk]
                    bgc = get_bg(val_h, hk)
                    
                    html_sheet.append(f"<td style='border-left:1px solid #ccc; border-right:1px solid #ccc; background-color:#fff; padding: 0px !important;'>{wrap_name(val_s, '')}</td>")
                    html_sheet.append(f"<td style='border-left:1px solid #ccc; border-right:1px solid #333; font-weight:bold; background-color:{bgc}; padding: 0px !important; line-height: 1.0 !important;'>{wrap_name(val_h, hk)}</td>")
            else:
                html_sheet.append("<td colspan='6' style='border: 1px solid #eee; background-color: #fdfdfd; opacity:0.2;'></td>")
        html_sheet.append("</tr>")
        
    html_sheet.append("</table></div>")
    st.html(f"<div style='border: 2px solid #999; background-color: #ffffff; border-radius: 6px; padding: 5px;'>{''.join(html_sheet)}</div>")

# タブ3：1日集中表示
with view_mode[2]:
    if 'current_day_val' not in st.session_state:
        st.session_state.current_day_val = 1
        
    d_col1, d_col2, d_col3 = st.columns([1.2, 4, 1.2])
    with d_col1:
        if st.button("← 前の日", use_container_width=True, key="prev_day_btn"):
            if st.session_state.current_day_val > 1:
                st.session_state.current_day_val -= 1
    with d_col3:
        if st.button("次の日 →", use_container_width=True, key="next_day_btn"):
            if st.session_state.current_day_val < max_days:
                st.session_state.current_day_val += 1
    with d_col2:
        select_day = st.slider("日slider", min_value=1, max_value=max_days, value=st.session_state.current_day_val, key="day_slider_bar")
        st.session_state.current_day_val = select_day
    
    st.write("---")
    
    st.html("""
    <button onclick="printTabSection('day-view-container', 'portrait')" class="pdf-btn-green">
        📄 1日詳細スケジュールをPDFで閲覧 / ダウンロードする（A4縦大判）
    </button>
    """)
    
    try:
        this_date = datetime.date(target_year, target_month, st.session_state.current_day_val)
        wd_str = ["月", "火", "水", "木", "金", "土", "日"][this_date.weekday()]
    except:
        wd_str = ""
        
    st.markdown(f"<h2 style='text-align: center; color: #1c83e1;'>🔍 {target_month}月 {st.session_state.current_day_val}日 ({wd_str}曜日) の詳細シフト</h2>", unsafe_allow_html=True)
    
    large_table_html = make_html_table_with_time(calendar_data[st.session_state.current_day_val], font_size="15px", padding="6px", is_large=True)
    st.html(f"""
        <div id='day-view-container' style='max-width: 650px; margin: 0 auto; border: 2px solid #1c83e1; background-color: #ffffff; border-radius: 8px; padding: 10px;'>
            {large_table_html}
        </div>
    """)

# タブ4：公開ページ確認
with view_mode[3]:
    st.markdown("<h2 style='color: #1c83e1;'>🌐 テストアップ公開データの確認</h2>", unsafe_allow_html=True)
    st.write("確定送信ボタンを押した後、実際にインターネット上に書き出された予定表のページです。")
    
    st.link_button("🔗 新しいタブで実際の公開ページを開く", PUBLIC_URL, use_container_width=True)
    st.write("---")
    st.caption("👇 アプリ内で簡易プレビュー")
    st.components.v1.iframe(PUBLIC_URL, height=800, scrolling=True)
