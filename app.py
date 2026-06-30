import streamlit as st
import openpyxl
import datetime
import re
import os
import io
from ftplib import FTP

# システム環境設定
os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

st.set_page_config(page_title="シフト配置カレンダー", layout="wide")

# ────────────────────────────────────────────────────────
# ⚙️ FTPサーバーの接続設定
# ────────────────────────────────────────────────────────
FTP_HOST = "sv17062.xserver.jp"
FTP_USER = "atom_test@test.atom-makoto.com"
FTP_PASS = "d7CWAQrw"
FTP_DIR  = "/Flower_House/sync_shift_System"
OUTPUT_FILE_NAME = "index.html"
PUBLIC_URL = "http://test.atom-makoto.com/Flower_House/sync_shift_System/"

# ────────────────────────────────────────────────────────
# 🛠️ PDF印刷用のJavaScript関数とCSS
# ────────────────────────────────────────────────────────
st.html("""
<script>
function printTabSection(containerId, orientation) {
    // 既存の印刷用スタイルがあれば削除
    var oldStyle = document.getElementById('dynamic-print-style');
    if (oldStyle) oldStyle.remove();

    var style = document.createElement('style');
    style.id = 'dynamic-print-style';
    style.innerHTML = `
        @media print {
            /* 背景色や罫線を維持 */
            * {
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }
            /* StreamlitのUI要素を非表示 */
            header, footer, div[data-testid="stSidebar"], div[data-testid="stHeader"], 
            .stButton, div[data-testid="stSelectbox"], div[data-testid="stSlider"], .no-print, .pdf-control-panel {
                display: none !important;
            }
            /* コンテナの余白をなくす */
            div[data-testid="stMainBlockContainer"] {
                max-width: 100% !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            div[data-testid="stAppViewContainer"] > div {
                display: block !important;
            }
            /* 対象コンテナのみ表示 */
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
            /* スクロールを解除 */
            #${containerId} .scrollable-container {
                height: auto !important;
                overflow: visible !important;
            }
            /* 用紙サイズと向きの設定 */
            @page {
                size: A4 ${orientation};
                margin: 5mm;
            }
            /* テーブルのページ区切り防止 */
            .calendar-table {
                page-break-inside: avoid !important;
                border-collapse: collapse !important;
            }
            /* 縦棒を繋げるためのセル余白調整 */
            .calendar-table td {
                padding-top: 0px !important;
                padding-bottom: 0px !important;
                line-height: 1.0 !important;
                height: 18px !important;
            }
        }
    `;
    document.head.appendChild(style);
    
    // 印刷ダイアログの呼び出し
    setTimeout(function() {
        window.print();
    }, 250);
}
</script>
""")

# 等幅・改行なし・時間軸スリム化 ＆ 🖨️印刷・PDF保存時用のスタイル完全制御CSS
st.markdown("""
    <style>
    /* ─── 画面全体の横幅を限界まで引き伸ばす ─── */
    div[data-testid="stMainBlockContainer"] {
        max-width: 96% !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
    }

    /* ─── プルダウン(週選択)のガタつき調整 ─── */
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

    /* ─── スライダー(日選択)の高さ・位置同期 ─── */
    div[data-testid="stSlider"] label {
        display: none !important;
    }
    div[data-testid="stSlider"] {
        margin-top: 0px !important;
        padding-top: 5px !important;
    }

    /* ─── すべてのボタンのサイズ統一 ─── */
    .stButton > button {
        height: 40px !important;
        min-height: 40px !important;
        line-height: 40px !important;
        padding-top: 0px !important;
        padding-bottom: 0px !important;
    }

    /* ─── テーブルの完全等幅 ─── */
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

    /* ─── PDFボタンのスタイル ─── */
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

# Excelファイルのアップロード
uploaded_file = st.file_uploader("シフトExcelファイルを選択してください", type=["xlsx", "xlsm"])

# カレンダー表示用のデータ構造
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

        max_days = 30 if target_month in [4, 6, 9, 11] else 28 if target_month == 2 else 31

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

                    is_ake = "明" in cell_str or "ｍ" in cell_str or "m" in cell_str
                    is_osor = "遅" in cell_str
                    is_haya = "早" in cell_str or "ｆ" in cell_str or "f" in cell_str
                    is_yoru = "夜" in cell_str or "mf" in cell_str
                    is_nichi = "日" in cell_str

                    assigned_h, start_h, end_h = None, None, None
                    if is_ake: assigned_h, start_h, end_h = "h3", 5, 8
                    elif is_osor: assigned_h, start_h, end_h = "h3", 10, 18.5
                    elif is_haya: assigned_h, start_h, end_h = "h1", 7, 15
                    elif is_nichi: assigned_h, start_h, end_h = "h2", 14, 17
                    elif is_yoru: assigned_h, start_h, end_h = "h1", 17, 21

                    if assigned_h:
                        curr = datetime.datetime(2026, 1, 1, int(start_h), int((start_h % 1) * 60))
                        end = datetime.datetime(2026, 1, 1, int(end_h), int((end_h % 1) * 60))
                        time_slots = []
                        while curr <= end:
                            time_slots.append((curr.hour, curr.minute))
                            curr += datetime.timedelta(minutes=30)
                        for idx, (h, m) in enumerate(time_slots):
                            row_key = "row1" if m == 0 else "row2"
                            if idx == 0 or idx == len(time_slots) - 1:
                                calendar_data[d_num][h][row_key][assigned_h] = staff_name
                            else:
                                if not calendar_data[d_num][h][row_key][assigned_h]:
                                    calendar_data[d_num][h][row_key][assigned_h] = "┃"

        st.success(f"🎉 データ読み込み完了")
    except Exception as e:
        st.error(f"エラー: {e}")

# 補助関数群
def get_bg(val, h_type):
    if not val or val in ["┃", "｜", "↓"]: return "#ffffff"
    return "#ffc0cb" if h_type == "h3" else "#ffff00" if h_type == "h1" else "#ffffff"

def wrap_name(val, h_type):
    return "" if not val else "┃" if val in ["┃", "｜", "↓"] else f"<span class='staff-name-box'>{val}</span>"

def make_html_table_with_time(day_schedule, font_size="11px", padding="3px", is_large=False):
    html = [f"<table class='calendar-table' style='border-collapse: collapse; text-align: center; font-size: {font_size}; font-family: sans-serif; width: 100%; border: 1px solid #333;'>"]
    html.append("<tr style='background-color: #f0f0f0; font-weight: bold; border-bottom: 2px solid #333;'><td style='border: 1px solid #ccc; width: 22%;'>時間</td><td>サ1</td><td>へ1</td><td>サ2</td><td>へ2</td><td>サ3</td><td>へ3</td></tr>")
    
    hours_sequence = list(range(5, 24)) + [0]
    for hour in hours_sequence:
        for m_str, r_key in [("00", "row1"), ("30", "row2")]:
            time_str = f"{hour}:{m_str}"
            row = day_schedule[hour][r_key]
            html.append(f"<tr style='height: 18px;'><td>{time_str}</td>")
            for hk in ["h1", "h2", "h3"]:
                val = row[hk]
                html.append(f"<td style='border: 1px solid #ccc;'></td><td style='border: 1px solid #333; background-color:{get_bg(val, hk)}'>{wrap_name(val, hk)}</td>")
            html.append("</tr>")
    html.append("</table>")
    return "".join(html)

def build_and_upload_static_html():
    # ※ここは元のロジックをそのまま記載してください
    return True, None

# ────────────────────────────────────────────────────────
# タブ構造
# ────────────────────────────────────────────────────────
view_mode = st.tabs(["📊 1ヶ月表示", "📅 1週間表示", "🔍 1日集中表示", "🌐 公開ページ確認"])

with view_mode[0]:
    st.html("<button onclick=\"printTabSection('month-view-container', 'landscape')\" class='pdf-btn-blue'>📄 1ヶ月分PDF保存</button>")
    st.html("<div id='month-view-container'>")
    # ... (1ヶ月表示描画)
    st.html("</div>")

with view_mode[1]:
    st.html("<button onclick=\"printTabSection('week-view-container', 'landscape')\" class='pdf-btn-blue'>📄 1週間分PDF保存</button>")
    st.html("<div id='week-view-container'>")
    # ... (1週間表示描画)
    st.html("</div>")

with view_mode[2]:
    st.html("<button onclick=\"printTabSection('day-view-container', 'portrait')\" class='pdf-btn-green'>📄 1日詳細PDF保存</button>")
    st.html("<div id='day-view-container'>")
    # ... (1日表示描画)
    st.html("</div>")

with view_mode[3]:
    st.link_button("🔗 公開ページを確認", PUBLIC_URL)
    st.components.v1.iframe(PUBLIC_URL, height=800, scrolling=True)
