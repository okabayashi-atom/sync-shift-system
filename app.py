code = """import streamlit as st
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
# ⚙️ FTPサーバーの接続設定（テストアップ用アカウント）
# ────────────────────────────────────────────────────────
FTP_HOST = "sv17062.xserver.jp"
FTP_USER = "atom_test@test.atom-makoto.com"
FTP_PASS = "d7CWAQrw"
FTP_DIR  = "/Flower_House/sync_shift_System"
OUTPUT_FILE_NAME = "index.html"

# 🌐 テストサーバー用の公開URL
PUBLIC_URL = "http://test.atom-makoto.com/Flower_House/sync_shift_System/"

# ────────────────────────────────────────────────────────
# 🛠️ PDF閲覧・ダウンロード時に、カラー・罫線・縦棒を完全に維持するCSS & JS
# ────────────────────────────────────────────────────────
st.html(\"\"\"
<script>
function printTabSection(containerId, orientation) {
    // 既存の印刷用動的スタイルがあれば削除
    var oldStyle = document.getElementById('dynamic-print-style');
    if (oldStyle) oldStyle.remove();

    // 印刷・PDF保存時に、背景色や罫線を100%維持し、A4サイズ・向きを完全制御するCSSを注入
    var style = document.createElement('style');
    style.id = 'dynamic-print-style';
    style.innerHTML = `
        @media print {
            /* ⚠️ 【超重要】PDF保存時も背景色（ピンク・黄色）や罫線を絶対に消さない設定 */
            * {
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }
            
            /* Streamlitのメニューやボタン、PDF制御パネル自体をPDF内では非表示にする */
            header, footer, div[data-testid="stSidebar"], div[data-testid="stHeader"], 
            .stButton, div[data-testid="stSelectbox"], div[data-testid="stSlider"], .no-print, .pdf-control-panel {
                display: none !important;
            }
            
            /* メイン容器の余白をゼロにして全面印刷可能にする */
            div[data-testid="stMainBlockContainer"] {
                max-width: 100% !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            div[data-testid="stAppViewContainer"] > div {
                display: block !important;
            }
            
            /* 印刷対象のタブコンテナだけを全面に引き出す */
            #${containerId} {
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
            
            /* カレンダー内のスクロール制限を解除し、下まで全て展開してPDF化する */
            #${containerId} .scrollable-container {
                height: auto !important;
                overflow: visible !important;
            }
            
            /* A4用紙の向きと余白を完全に固定 (landscape=横向き, portrait=縦向き) */
            @page {
                size: A4 ${orientation};
                margin: 6mm 5mm;
            }
            
            /* 表が途中の変な位置でページ切れするのを防止 */
            .calendar-table {
                page-break-inside: avoid !important;
                border-collapse: collapse !important;
            }
            
            /* PDF内でも縦棒が隙間なく綺麗に繋がるようにセルを強制密着 */
            .calendar-table td {
                padding-top: 0px !important;
                padding-bottom: 0px !important;
                line-height: 1.0 !important;
            }
        }
    `;
    document.head.appendChild(style);
    
    // プレビュー画面（＝PDF閲覧・ダウンロードの場所）を起動
    setTimeout(function() {
        window.print();
    }, 250);
}
</script>
\"\"\")

st.markdown(\"\"\"
    <style>
    /* 画面全体の横幅を限界まで引き伸ばす */
    div[data-testid="stMainBlockContainer"] {
        max-width: 97% !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }
    /* プルダウンのガタつき調整 */
    div[data-testid="stSelectbox"] label { display: none !important; }
    div[data-testid="stSelectbox"] { margin-top: 0px !important; padding-top: 0px !important; }
    
    /* 縦棒（┃）を画面上でも隙間なく一本に繋げるためのテーブル設定 */
    .calendar-table {
        border-collapse: collapse !important;
    }
    .calendar-table td {
        padding-top: 0px !important;
        padding-bottom: 0px !important;
        line-height: 1.0 !important;
        height: 18px !important;
    }
    /* 文字サイズ自動フィット */
    .staff-name-box {
        display: block !important;
        white-space: nowrap !important;
        overflow: visible !important;
        font-size: min(12px, 24cqw) !important;
        text-align: center !important;
        line-height: 1.0 !important;
    }
    
    /* 📥 PDF保存・閲覧用の専用コントロールパネルUI */
    .pdf-control-panel {
        background-color: #f0f7ff; 
        border: 1px solid #1c83e1; 
        padding: 12px 15px; 
        border-radius: 6px; 
        margin-bottom: 15px;
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
    }
    </style>
\"\"\", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; color: #333;'>📅 シフト配置確認カレンダー</h1>", unsafe_allow_html=True)

# Excelファイルのアップロード
uploaded_file = st.file_uploader("シフトExcelファイルを選択してください", type=["xlsx", "xlsm"])

# カレンダー表示用のデータ構造（5:00〜0:30）
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

        # ─── Excel原本解析ロジック ───
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

                    if is_ake: assigned_h = "h3"; start_hour, start_min = 5, 0; end_hour, end_min = 8, 0
                    elif is_osor: assigned_h = "h3"; start_hour, start_min = 10, 0; end_hour, end_min = 18, 30
                    elif is_haya: assigned_h = "h1"; start_hour, start_min = 7, 0; end_hour, end_min = 15, 0
                    elif is_nichi: assigned_h = "h2"; start_hour, start_min = 14, 0; end_hour, end_min = 17, 0
                    elif is_yoru: assigned_h = "h1"; start_hour, start_min = 17, 0; end_hour, end_min = 21, 0

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
        st.success(f"🎉 シート「{selected_sheet_name}」のデータを読み込みました。")
    except Exception as e:
        st.error(f"Excelの読み込みエラー: {e}")

st.write("---")

try: start_offset = (datetime.date(target_year, target_month, 1).weekday() + 1) % 7
except: start_offset = 0

if target_month in [4, 6, 9, 11]: max_days = 30
elif target_month == 2: max_days = 28
else: max_days = 31

hours_sequence = list(range(5, 24)) + [0]

def get_bg(val, h_type):
    if not val or val in ["┃", "｜", "↓"]: return "#ffffff"
    if h_type == "h3": return "#ffc0cb"
    if h_type == "h1": return "#ffff00"
    return "#ffffff"

def wrap_name(val, h_type):
    if not val: return ""
    if val in ["┃", "｜", "↓"]: return "┃"
    return f"<span class='staff-name-box'>{val}</span>"

# ─── 🖨️ 縦棒（┃）が上下綺麗につながる1日分テーブル生成関数 ───
def make_html_table_with_time(day_schedule, font_size="11px", padding="0px", is_large=False):
    html = []
    td_p_style = "padding: 0px 2px !important;" if not is_large else f"padding: {padding} 2px !important;"
    line_h_style = "line-height: 1.0 !important;" if not is_large else "line-height: 1.2 !important;"
    
    html.append(f"<table class='calendar-table' style='border-collapse: collapse; text-align: center; font-size: {font_size}; font-family: sans-serif; width: 100%; table-layout: fixed; border: 1px solid #333;'>")
    html.append("<tr style='background-color: #f0f0f0; font-weight: bold; border-bottom: 2px solid #333;'>")
    html.append(f"<td style='border: 1px solid #ccc; width: 24%; padding: 4px 0; font-size: 9px;'>時間</td>")
    html.append("<td style='border: 1px solid #ccc; width: 12%; font-size: 9px;'>サ1</td><td style='border: 1px solid #ccc; width: 13%; font-weight: bold; font-size: 9px;'>へ1</td>")
    html.append("<td style='border: 1px solid #ccc; width: 12%; font-size: 9px;'>サ2</td><td style='border: 1px solid #ccc; width: 13%; font-weight: bold; font-size: 9px;'>へ2</td>")
    html.append("<td style='border: 1px solid #ccc; width: 12%; font-size: 9px;'>サ3</td><td style='border: 1px solid #ccc; width: 14%; font-weight: bold; font-size: 9px;'>へ3</td>")
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
        
        for h_key in ["h1", "h2", "h3"]:
            val = row_data[h_key]
            bg_color = get_bg(val, h_key)
            
            has_above = (idx > 0 and rows_list[idx-1][1][h_key] != "")
            has_below = (idx < len(rows_list) - 1 and rows_list[idx+1][1][h_key] != "")
            
            top_border = "none" if (has_above and val != "") else "1px solid #ccc"
            bottom_border = "none" if (has_below and val != "") else "1px solid #ccc"
            
            # 縦棒密着結合用のスタイルをインラインで強制
            c_style = f"border-left: 1px solid #ccc; border-right: 1px solid #ccc; border-top: {top_border}; border-bottom: {bottom_border}; background-color: #ffffff; {td_p_style} {line_h_style}"
            html.append(f"<td style='{c_style}'></td>")
            
            s_style = f"border-left: 1px solid #ccc; border-right: 1px solid #333; border-top: {top_border}; border-bottom: {bottom_border}; font-weight: bold; background-color: {bg_color}; {td_p_style} {line_h_style}"
            cell_content = wrap_name(val, h_key)
            html.append(f"<td style='{s_style}'>{cell_content}</td>")
        html.append("</tr>")
    html.append("</table>")
    return "".join(html)

# ─── 💾 サーバー保存用HTML出力関数 ───
def build_and_upload_static_html():
    html_src = f"<html><head><meta charset='utf-8'><title>{target_month}月シフト</title></head><body>"
    for d_pointer in range(1, max_days + 1):
        html_src += f"<h3>📅 {target_month}月 {d_pointer}日</h3>{make_html_table_with_time(calendar_data[d_pointer])}"
    html_src += "</body></html>"
    try:
        with FTP(FTP_HOST, FTP_USER, FTP_PASS) as ftp:
            ftp.cwd(FTP_DIR); bio = io.BytesIO(html_src.encode('utf-8'))
            ftp.storbinary(f"STOR {OUTPUT_FILE_NAME}", bio)
        return True, None
    except Exception as ex: return False, str(ex)

if uploaded_file is not None:
    if st.button("🚀 この内容をサーバーのWebシステムへ公開保存する（確定送信）", use_container_width=True):
        with st.spinner("サーバーを更新中..."):
            success, err = build_and_upload_static_html()
            if success: st.success("🟢 サーバー側の公開データを更新しました。")
            else: st.error(f"🔴 送信失敗: {err}")

# ────────────────────────────────────────────────────────
# 📊 表示切り替えタブ（カレンダー・週間横並び・1日大判を完全復元）
# ────────────────────────────────────────────────────────
view_mode = st.tabs(["📊 1ヶ月表示（カレンダー形式）", "📅 1週間表示（A4横並び形式）", "🔍 1日集中表示（A4縦大判）", "🌐 公開ページ確認"])

# ─── 🛠️ タブ1：1ヶ月カレンダー表示 ───
with view_mode[0]:
    # 🌟 [PDF閲覧・ダウンロード専用の案内＆ボタンパネル]
    st.html(f\"\"\"
    <div class='pdf-control-panel'>
        <h4 style='margin: 0 0 5px 0; color: #1c83e1; font-size: 14px;'>📄 1ヶ月分カレンダー形式のPDF閲覧・ダウンロード</h4>
        <p style='margin: 0 0 10px 0; font-size: 12px; color: #555;'>下のボタンを押すと専用の<b>PDF閲覧・保存画面（印刷プレビュー）</b>が開きます。送信先の選択で「PDFとして保存」を選ぶことでダウンロードできます。（A4横向きに完全最適化され、縦棒も繋がります）</p>
        <button onclick="printTabSection('month-view-container', 'landscape')" class='pdf-btn-blue'>
            🔍 1ヶ月カレンダーをPDFで閲覧・ダウンロードする
        </button>
    </div>
    \"\"\")
    
    # グリッドカレンダーの構築
    month_html = []
    month_html.append("<div id='month-view-container'>")
    month_html.append("<table style='width:100%; border-collapse:collapse; table-layout:fixed; background-color:#fff; border:2px solid #333;'>")
    
    weekdays_labels = ["日", "月", "火", "水", "木", "金", "土"]
    month_html.append("<tr style='background-color:#f0f0f0; font-weight:bold; border-bottom:2px solid #333;'>")
    for idx, wl in enumerate(weekdays_labels):
        color = '#ff4b4b' if idx == 0 else '#1c83e1' if idx == 6 else '#333333'
        month_html.append(f"<th style='border:1px solid #ccc; padding:6px; text-align:center; color:{color}; font-size:14px;'>{wl}曜日</th>")
    month_html.append("</tr>")
    
    day_pointer = 1
    for week in range(6):
        if day_pointer > max_days: break
        month_html.append("<tr>")
        for day_of_week in range(7):
            current_cell_idx = week * 7 + day_of_week
            if start_offset <= current_cell_idx and day_pointer <= max_days:
                month_html.append("<td style='border:1px solid #999; vertical-align:top; padding:4px; width:14.2%;'>")
                month_html.append(f"<div style='font-weight:bold; font-size:12px; margin-bottom:3px;'>📅 {day_pointer}日</div>")
                # 画面上は400px、PDF時は全展開
                month_html.append("<div class='scrollable-container' style='height:400px; overflow-y:auto; border:1px solid #eee; background-color:#fff;'>")
                month_html.append(make_html_table_with_time(calendar_data[day_pointer], font_size="9px", padding="0px", is_large=False))
                month_html.append("</div></td>")
                day_pointer += 1
            else:
                month_html.append("<td style='border:1px solid #eee; background-color:#f9f9f9; opacity:0.4;'></td>")
        month_html.append("</tr>")
    month_html.append("</table></div>")
    st.html("".join(month_html))


# ─── 🛠️ タブ2：1週間表示（A4横向き・横並び） ───
with view_mode[1]:
    if 'current_week_idx' not in st.session_state: st.session_state.current_week_idx = 0
    weeks_list = ["第1週 (1日〜)", "第2週", "第3週", "第4週", "第5週", "第6週"]
    
    w_c1, w_c2, w_c3 = st.columns([1, 4, 1])
    with w_c1:
        if st.button("← 前の週", use_container_width=True, key="p_wk"):
            if st.session_state.current_week_idx > 0: st.session_state.current_week_idx -= 1
    with w_c3:
        if st.button("次の週 →", use_container_width=True, key="n_wk"):
            if st.session_state.current_week_idx < len(weeks_list) - 1: st.session_state.current_week_idx += 1
    with w_c2:
        week_option = st.selectbox("週選択", options=weeks_list, index=st.session_state.current_week_idx)
        st.session_state.current_week_idx = weeks_list.index(week_option)

    # 🌟 [PDF閲覧・ダウンロード専用の案内＆ボタンパネル]
    st.html(f\"\"\"
    <div class='pdf-control-panel' style='margin-top:10px;'>
        <h4 style='margin: 0 0 5px 0; color: #1c83e1; font-size: 14px;'>📄 選択中の週間横並び予定表のPDF閲覧・ダウンロード</h4>
        <p style='margin: 0 0 10px 0; font-size: 12px; color: #555;'>下のボタンを押すと、この週の予定表が<b>A4横向きにぴったり収まったPDF閲覧・保存画面（プレビュー）</b>として開きます。「PDFとして保存」でダウンロードしてください。</p>
        <button onclick="printTabSection('week-view-container', 'landscape')" class='pdf-btn-blue'>
            🔍 この週間横並び表をPDFで閲覧・ダウンロードする
        </button>
    </div>
    \"\"\")

    start_d = st.session_state.current_week_idx * 7 + 1 - start_offset
    
    # 週間横並び大テーブルの構築
    week_html = []
    week_html.append("<div id='week-view-container' style='overflow-x:auto;'>")
    week_html.append("<table class='calendar-table' style='width:100%; border-collapse:collapse; text-align:center; font-size:11px; font-family:sans-serif; table-layout:fixed; border:2px solid #333;'>")
    
    # 曜日・日付ヘッダー
    week_html.append("<tr style='background-color:#f0f0f0; font-weight:bold; border-bottom:2px solid #333;'>")
    week_html.append("<td style='border:1px solid #333; width:4%; padding:6px 0; font-size:10px;'>時間</td>")
    for day_of_week in range(7):
        current_d = start_d + day_of_week
        color = '#ff4b4b' if day_of_week == 0 else '#1c83e1' if day_of_week == 6 else '#333333'
        if 1 <= current_d <= max_days:
            week_html.append(f"<td colspan='6' style='border:1px solid #333; color:{color}; font-size:12px; width:13.7%;'>{weekdays_labels[day_of_week]} ({current_d}日)</td>")
        else:
            week_html.append(f"<td colspan='6' style='border:1px solid #333; color:#aaa; background-color:#fafafa; width:13.7%;'>{weekdays_labels[day_of_week]} (外)</td>")
    week_html.append("</tr><tr style='background-color:#f9f9f9; font-size:9px; border-bottom:1px solid #333;'><td>-</td>")
    for _ in range(7):
        week_html.append("<td style='border:1px solid #ccc; width:1.4%;'>サ1</td><td style='border:1px solid #333; font-weight:bold; width:2.5%;'>へ1</td>")
        week_html.append("<td style='border:1px solid #ccc; width:1.4%;'>サ2</td><td style='border:1px solid #333; font-weight:bold; width:2.5%;'>へ2</td>")
        week_html.append("<td style='border:1px solid #ccc; width:1.4%;'>サ3</td><td style='border:1px solid #333; font-weight:bold; width:2.5%;'>へ3</td>")
    week_html.append("</tr>")
    
    # 30分ごとの縦軸ループ
    rows_list = []
    for hour in hours_sequence:
        rows_list.append((f"{hour}:00", True))
        rows_list.append(("0:30" if hour == 0 else f"{hour}:30", False))
        
    for idx, (time_str, is_even) in enumerate(rows_list):
        bbs = "border-bottom:1px dashed #bbb;" if is_even else "border-bottom:1px solid #333;"
        if idx == len(rows_list) - 1: bbs = ""
        week_html.append(f"<tr style='{bbs} height: 18px;'><td style='border-right:1px solid #333; border-left:1px solid #ccc; background-color:#f2f2f2; font-weight:bold; padding:2px 0;'>{time_str}</td>")
        
        for day_of_week in range(7):
            current_d = start_d + day_of_week
            if 1 <= current_d <= max_days:
                ds = calendar_data[current_d]
                h_num = int(time_str.split(":")[0])
                rk = "row1" if time_str.endswith(":00") else "row2"
                rd = ds[h_num][rk]
                
                for hk in ["h1", "h2", "h3"]:
                    val = rd[hk]; bgc = get_bg(val, hk)
                    
                    has_above = False
                    if idx > 0:
                        pt, _ = rows_list[idx-1]; ph = int(pt.split(":")[0]); prk = "row1" if pt.endswith(":00") else "row2"
                        if ds[ph][prk][hk] != "": has_above = True
                    has_below = False
                    if idx < len(rows_list) - 1:
                        nt, _ = rows_list[idx+1]; nh = int(nt.split(":")[0]); nrk = "row1" if nt.endswith(":00") else "row2"
                        if ds[nh][nrk][hk] != "": has_below = True
                        
                    tb = "none" if (has_above and val != "") else "1px solid #ccc"
                    bb = "none" if (has_below and val != "") else "1px solid #ccc"
                    
                    week_html.append(f"<td style='border-left:1px solid #ccc; border-right:1px solid #ccc; border-top:{tb}; border-bottom:{bb}; background-color:#fff; padding: 0px !important;'></td>")
                    week_html.append(f"<td style='border-left:1px solid #ccc; border-right:1px solid #333; border-top:{tb}; border-bottom:{bb}; font-weight:bold; background-color:{bgc}; padding: 0px !important; line-height: 1.0 !important;'>{wrap_name(val, hk)}</td>")
            else:
                week_html.append("<td colspan='6' style='border:1px solid #eee; background-color:#fdfdfd; opacity:0.1;'></td>")
        week_html.append("</tr>")
    week_html.append("</table></div>")
    st.html("".join(week_html))


# ─── 🛠️ タブ3：1日集中表示（A4縦・拡大） ───
with view_mode[2]:
    if 'current_day_val' not in st.session_state: st.session_state.current_day_val = 1
    d_c1, d_c2, d_c3 = st.columns([1, 4, 1])
    with d_c1:
        if st.button("← 前の日", use_container_width=True, key="p_dy"):
            if st.session_state.current_day_val > 1: st.session_state.current_day_val -= 1
    with d_c3:
        if st.button("次の日 →", use_container_width=True, key="n_dy"):
            if st.session_state.current_day_val < max_days: st.session_state.current_day_val += 1
    with d_c2:
        select_day = st.slider("日選択", min_value=1, max_value=max_days, value=st.session_state.current_day_val)
        st.session_state.current_day_val = select_day
        
    # 🌟 [PDF閲覧・ダウンロード専用の案内＆ボタンパネル]
    st.html(f\"\"\"
    <div class='pdf-control-panel' style='margin-top:10px;'>
        <h4 style='margin: 0 0 5px 0; color: #4CAF50; font-size: 14px;'>📄 選択中の1日集中予定表のPDF閲覧・ダウンロード</h4>
        <p style='margin: 0 0 10px 0; font-size: 12px; color: #555;'>下のボタンを押すと、この日の詳細スケジュールが<b>A4縦向きに1枚おっきく拡大されたPDF閲覧・保存画面（プレビュー）</b>として開きます。「PDFとして保存」でダウンロードしてください。</p>
        <button onclick="printTabSection('day-view-container', 'portrait')" class='pdf-btn-green'>
            🔍 この1日スケジュールをPDFで閲覧・ダウンロードする
        </button>
    </div>
    \"\"\")

    try:
        this_date = datetime.date(target_year, target_month, st.session_state.current_day_val)
        wd_str = ["月", "火", "水", "木", "金", "土", "日"][this_date.weekday()]
    except: wd_str = ""
        
    st.markdown(f"<h2 style='text-align: center; color: #1c83e1;'>🔍 {target_month}月 {st.session_state.current_day_val}日 ({wd_str}曜日)</h2>", unsafe_allow_html=True)
    
    # 1日用は大判サイズ(font_size=15px, padding=6px, is_large=True)で大きく引き伸ばす
    large_table_html = make_html_table_with_time(calendar_data[st.session_state.current_day_val], font_size="15px", padding="6px", is_large=True)
    st.html(f\"\"\"
        <div id='day-view-container' style='max-width: 650px; margin: 0 auto; background-color: #ffffff; padding: 10px;'>
            {large_table_html}
        </div>
    \"\"\")


# ─── 🛠️ タブ4：公開ページ確認 ───
with view_mode[3]:
    st.markdown("<h2 style='color: #1c83e1;'>🌐 テストアップ公開データの確認</h2>", unsafe_allow_html=True)
    st.link_button("🔗 新しいタブで実際の公開ページを開く", PUBLIC_URL, use_container_width=True)
    st.write("---")
    st.components.v1.iframe(PUBLIC_URL, height=800, scrolling=True)
"""

with open("app.py", "w", encoding="utf-8") as f:
    f.write(code)

print("SUCCESS")
