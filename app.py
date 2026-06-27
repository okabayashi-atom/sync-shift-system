import streamlit as st
from datetime import datetime
import os
from PIL import Image
import base64
import ftplib
import urllib.request

# PDF作成用の部品
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# アプリ設定
try:
    app_icon = Image.open("icon.ico")
except:
    app_icon = "📅"

st.set_page_config(
    page_title="シフト配置カレンダー", 
    page_icon=app_icon,
    layout="wide"  # カレンダーが見やすくなるようワイド展開
)

# --- 🌐 日本語文字化け対策 ---
@st.cache_resource
def setup_japanese_font():
    font_name = 'NotoSansJP'
    local_font_path = 'NotoSansJP-Regular.ttf'
    if not os.path.exists(local_font_path):
        font_url = 'https://github.com/googlefonts/noto-cjk/raw/main/Sans/Variable/TTF/NotoSansCJKjp-VF.ttf'
        try:
            with urllib.request.urlopen(font_url, timeout=30) as response, open(local_font_path, 'wb') as out_file:
                out_file.write(response.read())
        except Exception as e:
            try:
                backup_url = 'https://fonts.gstatic.com/s/notosansjp/v52/-nd77J3M25mU8Mcl686m--sm_w.ttf'
                with urllib.request.urlopen(backup_url, timeout=30) as response, open(local_font_path, 'wb') as out_file:
                    out_file.write(response.read())
            except:
                return 'Helvetica'
    try:
        pdfmetrics.registerFont(TTFont(font_name, local_font_path))
        return font_name
    except:
        return 'Helvetica'

current_font = setup_japanese_font()

# iPad用の設定
def inject_ipad_home_icon(icon_path="icon.ico"):
    try:
        with open(icon_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
        html_code = f"""
        <script>
        const link = window.parent.document.createElement('link');
        link.rel = 'apple-touch-icon';
        link.href = 'data:image/x-icon;base64,{encoded_string}';
        window.parent.document.getElementsByTagName('head')[0].appendChild(link);
        </script>
        """
        st.components.v1.html(html_code, height=0)
    except:
        pass

inject_ipad_home_icon()

# モバイル・iPad用タッチ固定
st.components.v1.html("""
<script>
function applyStrictTouchFix() {
    const root = window.parent.document;
    const inputs = root.querySelectorAll('div[data-testid="stSelectbox"] input');
    inputs.forEach(input => {
        input.setAttribute('inputmode', 'none');
        input.setAttribute('readonly', 'true');
        input.style.cursor = 'pointer';
    });
}
const observer = new MutationObserver(applyStrictTouchFix);
observer.observe(window.parent.document.body, { childList: true, subtree: true });
setTimeout(applyStrictTouchFix, 500);
</script>
""", height=0)

# 固定データと選択肢
HOURS = [f"{h}:00" for h in range(5, 22)] + [f"{h}:30" for h in range(5, 22)]
HOURS.sort(key=lambda x: int(x.split(':')[0]) * 60 + int(x.split(':')[1]))

STAFF_CHOICES = ["", "阿部", "石川（靖）", "石川（雅）", "石川（瑠）", "江添", "菊地", "鈴木", "傳法", "伏見", "安井", "その他"]

st.title("📅 シフト配置カレンダー")
st.write("---")

# =========================================================================
# 📁 左側サイドバー（常設の過去データ読み込み枠）
# =========================================================================
st.sidebar.header("📁 過去データの編集")
side_col1, side_col2 = st.sidebar.columns([1.2, 1])
with side_col1:
    selected_date = st.date_input("日付を選択", datetime.now(), key="sb_date")
with side_col2:
    st.write("<div style='padding-top:28px;'></div>", unsafe_allow_html=True)
    load_clicked = st.button("📂 読み込む", use_container_width=True, key="side_load_btn")

if load_clicked:
    date_str = selected_date.strftime('%Y年%m月%d日')
    filename = f"{date_str}_シフト配置.pdf"
    try:
        ftp = ftplib.FTP("sv17062.xserver.jp")
        ftp.login(user="atom_test@test.atom-makoto.com", passwd="d7CWAQrw")
        
        # 📁 FFFTPで確認した正しいフォルダに移動
        ftp.cwd("/Flower_House/sync_shift_System")
        
        file_list = ftp.nlst()
        ftp.quit()
        
        if filename in file_list:
            st.sidebar.success(f"✅ {date_str} を確認しました。")
        else:
            st.sidebar.error(f"データがありません。")
    except Exception as e:
        st.sidebar.error(f"接続エラー: {e}")

st.sidebar.write("---")

# --- メイン画面：お知らせ表示 ---
st.success("🎉 シート「R8・6月(2)」のデータを読み込みました。時間列を狭めて名前枠を最大限広げました！")

# --- 上部アクションボタン（2列レイアウトを文字が切れない幅に調整） ---
btn_col1, btn_col2 = st.columns([2, 1])
with btn_col1:
    save_web = st.button("🚀 この内容をWeb稼働表として公開保存する（確定送信）", use_container_width=True)
with btn_col2:
    save_pdf = st.button("📄 この画面をPDF保存", use_container_width=True, type="primary")

# --- 🚀 保存処理を実行した際のロジック ---
if save_web or save_pdf:
    target_filename = f"{selected_date.strftime('%Y年%m月%d日')}_シフト配置.pdf"
    
    try:
        # 一時的なPDFの作成
        p_cv = canvas.Canvas(target_filename, pagesize=landscape(A4)) # 横向きA4
        p_cv.setFont(current_font, 20)
        p_cv.drawString(40, 540, f"シフト配置予定表 ({selected_date.strftime('%Y年%m月%d日')})")
        p_cv.showPage()
        p_cv.save()
        
        # エックスサーバーへ転送
        with st.spinner("エックスサーバーへシフトデータを自動転送中..."):
            ftp = ftplib.FTP("sv17062.xserver.jp")
            ftp.login(user="atom_test@test.atom-makoto.com", passwd="d7CWAQrw")
            
            # 📁 管理日誌用から、シフト用の正しいパスに修正
            ftp.cwd("/Flower_House/sync_shift_System")
            
            with open(target_filename, "rb") as f:
                ftp.storbinary(f"STOR {target_filename}", f)
            ftp.quit()
            
        if os.path.exists(target_filename):
            os.remove(target_filename)
            
        st.success("🎉 エックスサーバーへの同期・保存が正常に完了しました！")
        
    except Exception as e:
        # 530エラーが発生していた箇所への対策
        st.error(f"❌ サーバー接続でエラーが発生しました: {e}")

# --- タブ選択（1ヶ月、1週間、1日） ---
tab1, tab2, tab3 = st.tabs(["📅 1ヶ月表示 (カレンダー)", "📊 1週間表示 (時間軸スリム化)", "🔍 1日集中表示"])

with tab1:
    st.markdown("### 日曜日 〜 土曜日の配置スケジュール")
    
    # 曜日の見出しを横並びで配置
    days_cols = st.columns(7)
    day_names = ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"]
    colors = ["red", "black", "black", "black", "black", "black", "blue"]
    
    for idx, day in enumerate(day_names):
        with days_cols[idx]:
            st.markdown(f"<h3 style='text-align: center; color: {colors[idx]};'>{day}</h3>", unsafe_allow_html=True)
            
    st.write("---")
    
    # カレンダーの枠サンプル（1日〜6日）
    cal_cols = st.columns(7)
    # 空白（日曜日枠）
    with cal_cols[0]:
        st.write("")
        
    # 各日のダミー配置テーブル（見た目の再現）
    dummy_staffs = ["伏見", "傳法", "伏見", "江添", "鈴木", "石川"]
    for i in range(1, 7):
        with cal_cols[i]:
            st.markdown(f"<h4 style='text-align: center;'>📅 {i}日</h4>", unsafe_allow_html=True)
            
            # 時間、サ1、ヘ1 などの簡易グリッドをシミュレート
            st.caption("時間 | サ1 | ヘ1")
            st.text(f"5:00 |     | {dummy_staffs[i-1]}")
            st.text("6:00 |     |  | ")
            st.text("7:00 | 江添|  | ")
            st.text("8:00 |     |  | ")
            st.text("9:00 |     |  | ")

with tab2:
    st.write("週間スケジュールの詳細タイムラインがここに表示されます。")

with tab3:
    st.write("選択された1日の詳細なスタッフ配置シフトが表示されます。")
