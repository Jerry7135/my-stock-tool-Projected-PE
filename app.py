import streamlit as st
import pandas as pd
from fugle_marketdata import RestClient
import time
from datetime import datetime

# ==========================================
# 1. 基本設定與戰略風格 CSS 注入
# ==========================================
st.set_page_config(page_title="戰術監控終端 | 台股本益比", layout="wide", page_icon="📡")

# 注入 CSS 打造暗黑/螢光綠的戰略雷達風格
st.markdown("""
    <style>
    /* 標題與重點文字螢光綠 */
    h1, h2, h3 {
        color: #00ff41 !important;
        font-family: 'Courier New', Courier, monospace;
    }
    /* 按鈕戰術風格化 */
    .stButton>button {
        background-color: #003300;
        color: #00ff41;
        border: 1px solid #00ff41;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #00ff41;
        color: #000000;
        border: 1px solid #00ff41;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📡 台股即時本益比 | 雲端戰術面板")

# ==========================================
# 2. 雲端環境與 API 設定
# ==========================================
# 富果 API Key
FUGLE_API_KEY = "NGJjYzIxZjgtMGIzNS00MzAzLTk5ZGQtZjNkMTQ3MjI0ZDNiIGYxM2I0ODk5LTYxZDEtNDA3ZS04ZDhjLWVkMDUyNjkzNjc5OA=="

# 🔴 TODO: 請在此填入你的 Google Drive 檔案 ID
# 如何取得：在 Google Drive 對 Excel 檔案按右鍵 ->「共用」-> 權限改為「知道連結的任何人均可檢視」
# 複製連結後，取出 https://drive.google.com/file/d/【這中間的一長串代碼】/view 
DRIVE_FILE_ID = "18vi9hL3PUEivCnUVjpfN6MVoKa2YjTDU"

# ==========================================
# 3. 核心邏輯函式
# ==========================================
@st.cache_data(ttl=300) # 快取雲端 Excel 檔案 5 分鐘，避免每次重整都去下載
def load_cloud_data(file_id):
    # 將 File ID 轉換為 Pandas 可直接讀取的直連網址
    url = f"https://drive.google.com/uc?id={file_id}"
    
    try:
        # 這裡預設你的雲端檔案是 Excel (.xlsx)
        df = pd.read_excel(url, header=[0, 1])
        
        # --- 表頭清洗器 ---
        col_df = df.columns.to_frame(index=False)
        col_df.iloc[:, 0] = col_df.iloc[:, 0].apply(lambda x: pd.NA if 'Unnamed' in str(x) else x).ffill()
        for i in range(col_df.shape[1]):
            col_df.iloc[:, i] = col_df.iloc[:, i].apply(
                lambda x: "" if pd.isna(x) or 'Unnamed' in str(x) else str(x).replace('.0', '')
            )
        df.columns = pd.MultiIndex.from_frame(col_df)
        return df
    except Exception as e:
        st.error(f"❌ 雲端資料擷取失敗，請確認檔案共用權限已開啟！錯誤訊息：{e}")
        return None

def get_fugle_realtime_prices(symbols):
    client = RestClient(api_key=FUGLE_API_KEY)
    stock = client.stock
    prices = {}
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    total = len(symbols)
    
    for i, symbol in enumerate(symbols):
        status_text.text(f"📡 鎖定目標 {symbol} 擷取即時戰況... ({i+1}/{total})")
        try:
            symbol_str = str(symbol).split('.')[0].strip()
            quote = stock.intraday.quote(symbol=symbol_str)
            price = quote.get('lastPrice', quote.get('closePrice', None))
            if price is not None:
                prices[symbol_str] = price
        except Exception:
            pass
            
        progress_bar.progress((i + 1) / total)
        time.sleep(0.05)
        
    status_text.empty()
    progress_bar.empty()
    return prices

# ==========================================
# 4. 前端介面與自動化狀態管理
# ==========================================

# 建立 Session State
if "fetched_df" not in st.session_state:
    st.session_state.fetched_df = None
if "last_update_time" not in st.session_state:
    st.session_state.last_update_time = ""

# 建立上方控制列 (按鈕與更新時間並排)
col1, col2 = st.columns([2, 8])
with col1:
    refresh_btn = st.button("🔄 重新掃描最新報價", use_container_width=True)
with col2:
    if st.session_state.last_update_time:
        st.markdown(f"<p style='color:#00ff41; padding-top: 10px; font-weight: bold;'>⚡ 最後更新時間: {st.session_state.last_update_time}</p>", unsafe_allow_html=True)

# 觸發條件：1. 按下重新掃描按鈕  2. 第一次開啟網頁時 (fetched_df 為空)
if refresh_btn or st.session_state.fetched_df is None:
    if DRIVE_FILE_ID == "在這裡貼上你的_FILE_ID":
        st.warning("⚠️ 請先在程式碼第 39 行填入你的 Google Drive 檔案 ID！")
    else:
        # 1. 讀取雲端基礎資料
        base_df = load_cloud_data(DRIVE_FILE_ID)
        
        if base_df is not None:
            df = base_df.copy()
            
            with st.spinner("啟動富果連線，擷取即時價格中..."):
                code_col = [col for col in df.columns if '代碼' in str(col[0])][0]
                price_col = [col for col in df.columns if '最新收盤價' in str(col[0])][0]
                eps_years = [col[1] for col in df.columns if '財測EPS' in str(col[0]) and col[1] != ""]
                
                # 2. 抓取富果即時報價
                symbols = df[code_col].dropna().astype(str).str.split('.').str[0].unique()
                realtime_prices = get_fugle_realtime_prices(symbols)
                
                # 3. 更新數據
                for idx, row in df.iterrows():
                    sym = str(row[code_col]).split('.')[0]
                    if sym in realtime_prices:
                        real_price = realtime_prices[sym]
                        df.at[idx, price_col] = real_price
                        
                        for year in eps_years:
                            eps_col = ('財測EPS', year)
                            pe_col = ('本益比', year)
                            
                            if eps_col in df.columns and pe_col in df.columns:
                                eps_val = row[eps_col]
                                try:
                                    eps_val = float(eps_val)
                                    if eps_val > 0:
                                        df.at[idx, pe_col] = round(real_price / eps_val, 2)
                                except (ValueError, TypeError):
                                    pass

                # 將結果與當下時間存入 session_state
                st.session_state.fetched_df = df
                st.session_state.last_update_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                # 使用 rerun 刷新頁面
                st.rerun()

# ==========================================
# 5. 戰略風格資料表呈現
# ==========================================
if st.session_state.fetched_df is not None:
    display_df = st.session_state.fetched_df
    
    # 定位最新收盤價的欄位
    price_col = [col for col in display_df.columns if '最新收盤價' in str(col[0])][0]
    
    # 使用 Pandas Styler 進行戰略高亮 (深綠底、亮綠字、加粗)
    styled_df = display_df.style.set_properties(
        subset=[price_col], 
        **{
            'background-color': '#002200', 
            'color': '#00ff41', 
            'font-weight': 'bold', 
            'font-size': '15px'
        }
    ).format(precision=2)
    
    # 顯示最終成果
    st.dataframe(styled_df, use_container_width=True, height=600)
