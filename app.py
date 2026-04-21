import streamlit as st
import pandas as pd
from fugle_marketdata import RestClient
import time
from datetime import datetime

# ==========================================
# 1. 基本設定與極簡黑白 CSS 注入
# ==========================================
st.set_page_config(page_title="戰術監控終端 | 台股本益比", layout="wide", page_icon="📡")

# 注入 CSS：強制全螢幕黑底白字
st.markdown("""
    <style>
    .stApp {
        background-color: #000000;
        color: #ffffff;
    }
    h1, h2, h3, p, span {
        font-family: 'Courier New', Courier, monospace;
        color: #ffffff !important;
    }
    h1 {
        color: #00ff41 !important; 
    }
    .stButton>button {
        background-color: #111111;
        color: #ffffff;
        border: 1px solid #ffffff;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #ffffff;
        color: #000000;
        border: 1px solid #ffffff;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📡 台股即時本益比 | 雲端戰術面板")

# ==========================================
# 2. 雲端環境與金鑰設定
# ==========================================
FUGLE_API_KEY = "NGJjYzIxZjgtMGIzNS00MzAzLTk5ZGQtZjNkMTQ3MjI0ZDNiIGYxM2I0ODk5LTYxZDEtNDA3ZS04ZDhjLWVkMDUyNjkzNjc5OA=="
DRIVE_FILE_ID = "18tsyJeKzyEmWHLUG7OYdpdhWHFeMLdSz"  

# ==========================================
# 3. 核心邏輯函式
# ==========================================
@st.cache_data(ttl=300) 
def load_cloud_data(file_id):
    url = f"https://drive.google.com/uc?id={file_id}"
    try:
        df = pd.read_excel(url, header=[0, 1])
        
        # --- 表頭清洗器 ---
        col_df = df.columns.to_frame(index=False)
        col_df.iloc[:, 0] = col_df.iloc[:, 0].apply(lambda x: pd.NA if 'Unnamed' in str(x) else x).ffill()
        for i in range(col_df.shape[1]):
            col_df.iloc[:, i] = col_df.iloc[:, i].apply(
                lambda x: "" if pd.isna(x) or 'Unnamed' in str(x) else str(x).replace('.0', '')
            )
        df.columns = pd.MultiIndex.from_frame(col_df)
        
        # 💡 --- 資料清洗器：處理「跨欄置中」的空白 --- 💡
        # 尋找包含「產業類別」的欄位，並將空白處「向下填滿」(ffill)
        for col in df.columns:
            if '產業類別' in str(col[0]):
                df[col] = df[col].ffill()
                
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
            
            # 💡 升級版抓價邏輯：優先抓最新成交價，若無成交則抓昨日收盤價
            price = quote.get('lastPrice', quote.get('closePrice', quote.get('previousClose', None)))
            
            if price is not None:
                prices[symbol_str] = price
        except Exception:
            pass 
            
        progress_bar.progress((i + 1) / total)
        time.sleep(0.02) # 稍微抓一點安全間隔，避免被 API 擋下
        
    status_text.empty()
    progress_bar.empty()
    return prices

# ==========================================
# 4. 前端介面與自動化狀態管理
# ==========================================
if "fetched_df" not in st.session_state:
    st.session_state.fetched_df = None
if "last_update_time" not in st.session_state:
    st.session_state.last_update_time = ""

col1, col2 = st.columns([2, 8])
with col1:
    refresh_btn = st.button("🔄 重新掃描最新報價", use_container_width=True)
with col2:
    if st.session_state.last_update_time:
        st.markdown(f"<p style='color:#ffffff; padding-top: 10px; font-weight: bold;'>⚡ 最後更新時間: {st.session_state.last_update_time}</p>", unsafe_allow_html=True)

if refresh_btn or st.session_state.fetched_df is None:
    base_df = load_cloud_data(DRIVE_FILE_ID)
    
    if base_df is not None:
        df = base_df.copy()
        
        with st.spinner("啟動富果連線，擷取即時價格中..."):
            try:
                code_col = [col for col in df.columns if '代碼' in str(col[0])][0]
                price_col = [col for col in df.columns if '最新收盤價' in str(col[0])][0]
                eps_years = [col[1] for col in df.columns if '財測EPS' in str(col[0]) and col[1] != ""]
                
                symbols = df[code_col].dropna().astype(str).str.split('.').str[0].unique()
                realtime_prices = get_fugle_realtime_prices(symbols)
                
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

                st.session_state.fetched_df = df
                st.session_state.last_update_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                st.rerun()
                
            except IndexError:
                st.error("❌ 找不到關鍵欄位，請確認 Excel 表頭第一層包含『代碼』與『最新收盤價』！")

# ==========================================
# 5. 戰略風格資料表呈現 (條件化整行上色)
# ==========================================
if st.session_state.fetched_df is not None:
    display_df = st.session_state.fetched_df
    
    try:
        # 定位最新收盤價的欄位
        price_col = [col for col in display_df.columns if '最新收盤價' in str(col[0])][0]
        
        # 定義整行上色的邏輯
        def tactical_row_highlighter(row):
            # 檢查第一欄(A欄位)是否有寫註記
            val = str(row.iloc[0]).strip()
            is_annotated = val != "" and val.lower() != "nan" and val.lower() != "<na>"
            
            styles = []
            for col_name in row.index:
                if is_annotated:
                    # 【有註記的股票】：整行變成警戒色 (深紅底、亮黃字)
                    if col_name == price_col:
                        styles.append('background-color: #330000; color: #00ff41; font-weight: bold; border-bottom: 1px solid #550000;')
                    else:
                        styles.append('background-color: #330000; color: #ffdd00; font-weight: bold; border-bottom: 1px solid #550000;')
                else:
                    # 【沒有註記的股票】：一般的黑底白字
                    if col_name == price_col:
                        styles.append('background-color: #002200; color: #00ff41; font-weight: bold;')
                    else:
                        styles.append('background-color: #000000; color: #ffffff;')
            return styles

        # 應用樣式到 DataFrame
        styled_df = display_df.style.apply(tactical_row_highlighter, axis=1).format(precision=2)
        
        st.dataframe(styled_df, use_container_width=True, height=700)
    except Exception as e:
        st.dataframe(display_df, use_container_width=True, height=700)
