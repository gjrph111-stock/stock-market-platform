"""
台股多因子研究面板 v3

功能：
1. 讀取 output 資料夾內最新的掃描 Excel
2. 股票搜尋、觀察名單儲存、點選查看
3. 技術面：分級、主力突破、MA、MACD、RSI、技術圖
4. 基本面：即時抓取 yfinance 公司估值與財務摘要
5. 籌碼面：量價代理、主力突破分數、成交量觀察
6. 總體市場：加權、OTC、美股、VIX、美元台幣
7. K線圖（Candlestick）
8. CSV 下載

安裝：
py -m pip install streamlit pandas openpyxl pillow yfinance requests mplfinance matplotlib

執行：
py -m streamlit run tw_stock_dashboard.py
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import yfinance as yf
import mplfinance as mpf
import matplotlib.pyplot as plt
from PIL import Image


OUTPUT_DIR = Path("output")
WATCHLIST_FILE = Path("watchlist.csv")


st.set_page_config(
    page_title="台股多因子研究面板",
    page_icon="📈",
    layout="wide",
)


@st.cache_data(show_spinner=False)
def find_latest_excel() -> Path | None:
    files = sorted(
        OUTPUT_DIR.glob("台股全上市櫃盤後掃描_*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return files[0] if files else None


@st.cache_data(show_spinner=False)
def load_excel(file_path: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    result_df = pd.read_excel(file_path, sheet_name="掃描結果")

    try:
        failed_df = pd.read_excel(file_path, sheet_name="無資料或失敗")
    except Exception:
        failed_df = pd.DataFrame()

    try:
        chart_df = pd.read_excel(file_path, sheet_name="圖表索引")
    except Exception:
        chart_df = pd.DataFrame()

    return result_df, failed_df, chart_df


@st.cache_data(show_spinner=False)
def load_watchlist() -> pd.DataFrame:
    if WATCHLIST_FILE.exists():
        try:
            df = pd.read_csv(WATCHLIST_FILE, dtype={"股票代號": str})
            if "股票代號" not in df.columns:
                return pd.DataFrame(columns=["股票代號", "股票名稱", "市場", "加入時間"])
            return df
        except Exception:
            return pd.DataFrame(columns=["股票代號", "股票名稱", "市場", "加入時間"])
    return pd.DataFrame(columns=["股票代號", "股票名稱", "市場", "加入時間"])


def save_watchlist(df: pd.DataFrame) -> None:
    df = df.drop_duplicates(subset=["股票代號"], keep="last")
    df.to_csv(WATCHLIST_FILE, index=False, encoding="utf-8-sig")
    st.cache_data.clear()


def add_to_watchlist(stock_id: str, stock_name: str, market: str) -> None:
    watchlist = load_watchlist()
    new_row = pd.DataFrame(
        [
            {
                "股票代號": str(stock_id),
                "股票名稱": str(stock_name),
                "市場": str(market),
                "加入時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        ]
    )
    watchlist = pd.concat([watchlist, new_row], ignore_index=True)
    save_watchlist(watchlist)


def remove_from_watchlist(stock_id: str) -> None:
    watchlist = load_watchlist()
    if not watchlist.empty:
        watchlist = watchlist[watchlist["股票代號"].astype(str) != str(stock_id)]
        save_watchlist(watchlist)


def level_badge(level: str) -> str:
    if level == "A級":
        return "🟥 A級 強勢觀察"
    if level == "B級":
        return "🟧 B級 多頭續攻"
    if level == "C級":
        return "🟨 C級 初步轉強"
    if level == "排除":
        return "⬛ 排除"
    return str(level)


def format_number(value, percent: bool = False):
    try:
        if pd.isna(value):
            return "-"
        if value is None:
            return "-"
        if percent:
            return f"{float(value) * 100:.2f}%"
        if isinstance(value, float):
            return f"{value:,.2f}"
        if isinstance(value, int):
            return f"{value:,}"
        return value
    except Exception:
        return value


def guess_suffix(stock_id: str, market: str | None = None) -> str:
    if market == "上櫃":
        return ".TWO"
    return ".TW"


def get_yf_symbol(stock_id: str, market: str | None = None) -> str:
    stock_id = str(stock_id).strip()
    if stock_id.endswith(".TW") or stock_id.endswith(".TWO"):
        return stock_id
    return f"{stock_id}{guess_suffix(stock_id, market)}"


@st.cache_data(show_spinner=False, ttl=3600)
def get_fundamental_data(stock_id: str, market: str | None = None) -> dict:
    """基本面資料：使用 Yahoo Finance / yfinance。

    若 .TW 查不到，上櫃股票會再嘗試 .TWO。
    """
    candidates = [get_yf_symbol(stock_id, market)]
    if not candidates[0].endswith(".TWO"):
        candidates.append(f"{stock_id}.TWO")

    for symbol in candidates:
        try:
            ticker = yf.Ticker(symbol)
            info = ticker.info or {}
            if info and (info.get("shortName") or info.get("longName") or info.get("regularMarketPrice")):
                return {
                    "symbol": symbol,
                    "公司名稱": info.get("longName") or info.get("shortName") or "-",
                    "產業": info.get("industry") or "-",
                    "市值": info.get("marketCap"),
                    "本益比": info.get("trailingPE"),
                    "預估本益比": info.get("forwardPE"),
                    "EPS": info.get("trailingEps"),
                    "股價淨值比": info.get("priceToBook"),
                    "ROE": info.get("returnOnEquity"),
                    "毛利率": info.get("grossMargins"),
                    "營益率": info.get("operatingMargins"),
                    "殖利率": info.get("dividendYield"),
                    "52週高點": info.get("fiftyTwoWeekHigh"),
                    "52週低點": info.get("fiftyTwoWeekLow"),
                    "目前價格": info.get("regularMarketPrice") or info.get("currentPrice"),
                    "資料狀態": "成功",
                }
        except Exception:
            continue

    return {"symbol": candidates[0], "資料狀態": "查無資料"}


@st.cache_data(show_spinner=False, ttl=1800)
def get_market_dashboard_data() -> pd.DataFrame:
    symbols = {
        "加權指數": "^TWII",
        "OTC櫃買指數": "^TWOII",
        "S&P 500": "^GSPC",
        "Nasdaq": "^IXIC",
        "VIX": "^VIX",
        "美元台幣": "USDTWD=X",
    }

    rows = []
    for name, symbol in symbols.items():
        try:
            df = yf.download(symbol, period="3mo", interval="1d", progress=False, auto_adjust=False)
            if df.empty or len(df) < 2:
                rows.append({"項目": name, "代碼": symbol, "收盤": "-", "日漲跌%": "-", "20日趨勢": "-"})
                continue

            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(0)

            close = df["Close"].dropna()
            latest = float(close.iloc[-1])
            prev = float(close.iloc[-2])
            pct = (latest / prev - 1) * 100
            ma20 = close.rolling(20).mean().iloc[-1]
            trend = "偏多" if latest > ma20 else "偏弱"

            rows.append(
                {
                    "項目": name,
                    "代碼": symbol,
                    "收盤": round(latest, 2),
                    "日漲跌%": round(pct, 2),
                    "20日趨勢": trend,
                }
            )
        except Exception:
            rows.append({"項目": name, "代碼": symbol, "收盤": "-", "日漲跌%": "-", "20日趨勢": "讀取失敗"})

    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False, ttl=1800)
def generate_candlestick_chart(stock_id: str, market: str | None = None) -> Path | None:
    try:
        symbol = get_yf_symbol(stock_id, market)
        df = yf.download(symbol, period="6mo", interval="1d", progress=False, auto_adjust=False)

        if df.empty:
            return None

        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)

        required_cols = ["Open", "High", "Low", "Close", "Volume"]
        for col in required_cols:
            if col not in df.columns:
                return None

        df = df.dropna(subset=required_cols)

        df["MA5"] = df["Close"].rolling(5).mean()
        df["MA20"] = df["Close"].rolling(20).mean()
        df["MA60"] = df["Close"].rolling(60).mean()

        delta = df["Close"].diff()
        gain = delta.clip(lower=0)
        loss = -delta.clip(upper=0)
        avg_gain = gain.rolling(14).mean()
        avg_loss = loss.rolling(14).mean()
        rs = avg_gain / avg_loss
        df["RSI"] = 100 - (100 / (1 + rs))

        exp1 = df["Close"].ewm(span=12, adjust=False).mean()
        exp2 = df["Close"].ewm(span=26, adjust=False).mean()
        df["MACD"] = exp1 - exp2
        df["Signal"] = df["MACD"].ewm(span=9, adjust=False).mean()

        chart_dir = OUTPUT_DIR / "candlestick_charts"
        chart_dir.mkdir(parents=True, exist_ok=True)

        save_path = chart_dir / f"{stock_id}_candlestick.png"

        add_plots = [
            mpf.make_addplot(df["MA5"]),
            mpf.make_addplot(df["MA20"]),
            mpf.make_addplot(df["MA60"]),
            mpf.make_addplot(df["RSI"], panel=2, ylabel="RSI"),
            mpf.make_addplot(df["MACD"], panel=3, ylabel="MACD"),
            mpf.make_addplot(df["Signal"], panel=3),
        ]

        mpf.plot(
            df.tail(120),
            type="candle",
            mav=(5, 20, 60),
            volume=True,
            addplot=add_plots,
            figsize=(14, 10),
            style="yahoo",
            title=f"{stock_id} Candlestick Chart",
            savefig=dict(fname=str(save_path), dpi=120, bbox_inches="tight"),
            panel_ratios=(6, 2, 2, 2),
        )

        plt.close("all")

        return save_path

    except Exception:
        return None


def find_chart_path(stock_id: str, chart_df: pd.DataFrame) -> Path | None:
    if chart_df.empty:
        return None
    if "股票代號" not in chart_df.columns or "圖表路徑" not in chart_df.columns:
        return None

    matched = chart_df[chart_df["股票代號"].astype(str) == str(stock_id)]
    if matched.empty:
        return None

    path = Path(str(matched.iloc[0]["圖表路徑"]))
    return path if path.exists() else None


def find_stock_row(stock_id: str, result_df: pd.DataFrame, watchlist_df: pd.DataFrame) -> pd.Series | None:
    stock_id = str(stock_id)
    matched = result_df[result_df["股票代號"].astype(str) == stock_id]
    if not matched.empty:
        return matched.iloc[0]

    matched_watch = watchlist_df[watchlist_df["股票代號"].astype(str) == stock_id]
    if not matched_watch.empty:
        base = matched_watch.iloc[0].to_dict()
        base.setdefault("買進觀察分級", "未入選")
        base.setdefault("分級理由", "此股票目前不在最新掃描入選清單中，但仍保留於觀察名單。")
        base.setdefault("主力突破分數", "-")
        base.setdefault("主力突破訊號", "-")
        base.setdefault("條件明細", "-")
        return pd.Series(base)

    return None


def show_fundamental_section(stock_id: str, stock_name: str, market: str | None) -> None:
    st.write("### 基本面：企業體質與內在價值")
    data = get_fundamental_data(stock_id, market)

    if data.get("資料狀態") != "成功":
        st.warning("目前抓不到此股票的基本面資料。可能是 Yahoo Finance 暫時無資料，或市場代碼需要調整。")
        return

    st.caption(f"資料代碼：{data.get('symbol')}｜公司：{data.get('公司名稱', '-')}")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("目前價格", format_number(data.get("目前價格")))
    c2.metric("本益比", format_number(data.get("本益比")))
    c3.metric("EPS", format_number(data.get("EPS")))
    c4.metric("股價淨值比", format_number(data.get("股價淨值比")))
    c5.metric("殖利率", format_number(data.get("殖利率"), percent=True))

    c6, c7, c8, c9, c10 = st.columns(5)
    c6.metric("ROE", format_number(data.get("ROE"), percent=True))
    c7.metric("毛利率", format_number(data.get("毛利率"), percent=True))
    c8.metric("營益率", format_number(data.get("營益率"), percent=True))
    c9.metric("52週高點", format_number(data.get("52週高點")))
    c10.metric("52週低點", format_number(data.get("52週低點")))

    st.write(f"**產業**：{data.get('產業', '-')}")
    st.info("基本面資料主要用於觀察估值與體質，不應單獨作為買賣依據。台股部分欄位可能因資料源限制而缺漏。")


def show_chip_section(row: pd.Series) -> None:
    st.write("### 籌碼面：資金動向與主力意圖")
    st.caption("v1 先使用量價與主力突破代理指標；三大法人、融資融券與分點集中度可於下一版正式介接。")

    score = row.get("主力突破分數", "-")
    volume = row.get("成交量_張", "-")
    breakout = row.get("主力突破訊號", "-")
    level = row.get("買進觀察分級", "-")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("主力突破分數", format_number(score))
    c2.metric("成交量", f"{format_number(volume)} 張")
    c3.metric("分級", level_badge(str(level)))
    c4.metric("籌碼代理判斷", "偏強" if isinstance(score, (int, float)) and score >= 3 else "觀察")

    st.write(f"**量價訊號**：{breakout}")

    chip_notes = []
    try:
        if float(score) >= 4:
            chip_notes.append("量價結構偏積極，可能有資金集中跡象。")
        elif float(score) >= 2:
            chip_notes.append("有部分資金進場跡象，但仍需搭配法人與分點資料確認。")
        else:
            chip_notes.append("目前主力突破訊號不明顯，適合等待更明確量價結構。")
    except Exception:
        chip_notes.append("此股票不在最新掃描清單中，籌碼代理資料不足。")

    for note in chip_notes:
        st.write(f"- {note}")

    st.info("下一版建議接入：外資、投信、自營商買賣超、融資融券、主力分點。")


def show_macro_section() -> None:
    st.write("### 總體經濟與市場風向")
    market_df = get_market_dashboard_data()
    st.dataframe(market_df, use_container_width=True, hide_index=True)

    try:
        twii = market_df[market_df["項目"] == "加權指數"].iloc[0]
        otc = market_df[market_df["項目"] == "OTC櫃買指數"].iloc[0]
        vix = market_df[market_df["項目"] == "VIX"].iloc[0]

        notes = []
        notes.append(f"加權指數 20 日趨勢：{twii['20日趨勢']}")
        notes.append(f"OTC 20 日趨勢：{otc['20日趨勢']}")
        notes.append(f"VIX 狀態：{vix['收盤']}")

        st.write("**市場解讀**")
        for note in notes:
            st.write(f"- {note}")
    except Exception:
        st.caption("市場解讀暫時無法產生。")

    st.info("總體市場資料用於判斷順風/逆風環境。若大盤、OTC 同步偏弱，個股訊號需要降低部位與提高確認門檻。")


def show_stock_card(row: pd.Series, chart_df: pd.DataFrame) -> None:
    stock_id = str(row.get("股票代號", ""))
    stock_name = str(row.get("股票名稱", ""))
    market = str(row.get("市場", "上市"))
    level = str(row.get("買進觀察分級", ""))

    with st.container(border=True):
        top_cols = st.columns([1.1, 1.1, 1.1, 1.1, 2.5])
        top_cols[0].metric("股票", f"{stock_id} {stock_name}")
        top_cols[1].metric("分級", level_badge(level))
        top_cols[2].metric("主力突破分數", format_number(row.get("主力突破分數", "-")))
        top_cols[3].metric("收盤價", format_number(row.get("收盤價", "-")), f"{format_number(row.get('漲跌幅%', '-'))}%")
        top_cols[4].write("**分級理由**")
        top_cols[4].write(str(row.get("分級理由", "-")))

        detail_cols = st.columns([1, 1, 1, 1, 1, 2])
        detail_cols[0].write(f"**成交量**：{format_number(row.get('成交量_張', '-'))} 張")
        detail_cols[1].write(f"**RSI14**：{format_number(row.get('RSI14', '-'))}")
        detail_cols[2].write(f"**MACD**：{format_number(row.get('MACD', '-'))}")
        detail_cols[3].write(f"**20MA**：{format_number(row.get('20MA', '-'))}")
        detail_cols[4].write(f"**60MA**：{format_number(row.get('60MA', '-'))}")
        detail_cols[5].write(f"**主力突破訊號**：{row.get('主力突破訊號', '-')}")

        st.markdown("---")

        analysis_tabs = st.tabs(["技術面", "基本面", "籌碼面", "市場風向"])

        with analysis_tabs[0]:
            st.write("### 技術面：進出場時機與趨勢")
            st.write(f"- 技術條件：{row.get('條件明細', '-')}")
            st.write(f"- AI 技術評估：{row.get('分級理由', '-')}")
            st.write(f"- 主力突破訊號：{row.get('主力突破訊號', '-')}")
            st.write(f"- RSI14：{row.get('RSI14', '-')}")
            st.write(f"- MACD：{row.get('MACD', '-')}")
            st.write(f"- MA5 / MA20 / MA60：{row.get('5MA', '-')} / {row.get('20MA', '-')} / {row.get('60MA', '-')}")

            chart_path = find_chart_path(stock_id, chart_df)

            if chart_path:
                st.write("#### 掃描器技術圖")
                image = Image.open(chart_path)
                st.image(image, use_container_width=True)

            candle_path = generate_candlestick_chart(stock_id, market)

            if candle_path:
                st.write("#### K線圖（Candlestick）")
                candle_img = Image.open(candle_path)
                st.image(candle_img, use_container_width=True)
            else:
                st.caption("目前無法產生 K 線圖。")

        with analysis_tabs[1]:
            show_fundamental_section(stock_id, stock_name, market)

        with analysis_tabs[2]:
            show_chip_section(row)

        with analysis_tabs[3]:
            show_macro_section()


latest_excel = find_latest_excel()

st.title("📈 台股多因子研究面板")
st.caption("整合技術面、基本面、籌碼面與市場風向的盤後研究平台。")

if latest_excel is None:
    st.error("找不到掃描 Excel。請先執行 tw_stock_after_market_scanner.py。")
    st.stop()

result_df, failed_df, chart_df = load_excel(latest_excel)

if result_df.empty or "訊息" in result_df.columns:
    st.warning("目前沒有符合條件的股票。")
    st.stop()

result_df["股票代號"] = result_df["股票代號"].astype(str)
watchlist_df = load_watchlist()

file_time = datetime.fromtimestamp(latest_excel.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
st.sidebar.success(f"目前讀取：{latest_excel.name}")
st.sidebar.caption(f"檔案時間：{file_time}")

st.sidebar.header("篩選條件")

available_levels = [
    level for level in ["A級", "B級", "C級", "排除"]
    if level in result_df["買進觀察分級"].astype(str).unique()
]
selected_levels = st.sidebar.multiselect("買進觀察分級", options=available_levels, default=available_levels)

market_options = sorted(result_df["市場"].dropna().astype(str).unique().tolist()) if "市場" in result_df.columns else []
selected_markets = st.sidebar.multiselect("市場", options=market_options, default=market_options)

if "主力突破分數" in result_df.columns:
    min_breakout = int(result_df["主力突破分數"].fillna(0).min())
    max_breakout = int(result_df["主力突破分數"].fillna(0).max())
else:
    min_breakout = 0
    max_breakout = 0

breakout_range = st.sidebar.slider("主力突破分數", min_value=min_breakout, max_value=max_breakout, value=(min_breakout, max_breakout))

keyword = st.sidebar.text_input("搜尋股票代號 / 股票名稱")
show_cards = st.sidebar.checkbox("用卡片模式顯示", value=True)
show_only_with_charts = st.sidebar.checkbox("只顯示有圖表的股票", value=False)

st.sidebar.divider()
st.sidebar.header("自選觀察名單")
search_pick = st.sidebar.text_input("輸入股票代號加入觀察", placeholder="例如：2330")

if search_pick:
    pick = search_pick.strip()
    matched_pick = result_df[result_df["股票代號"].astype(str) == pick]
    if not matched_pick.empty:
        picked_row = matched_pick.iloc[0]
        st.sidebar.success(f"找到：{picked_row['股票代號']} {picked_row['股票名稱']}")
        if st.sidebar.button("加入觀察名單", key=f"add_{pick}"):
            add_to_watchlist(picked_row["股票代號"], picked_row["股票名稱"], picked_row.get("市場", "-"))
            st.sidebar.success("已加入觀察名單，請按 R 重新整理。")
    else:
        st.sidebar.warning("最新掃描清單中找不到此股票。可先等下一版加入任意股票查詢。")

watchlist_df = load_watchlist()
if not watchlist_df.empty:
    watchlist_options = [f"{row.股票代號} {row.股票名稱}" for row in watchlist_df.itertuples(index=False)]
    selected_watch = st.sidebar.selectbox("點選觀察股票", options=["未選擇"] + watchlist_options)
    if selected_watch != "未選擇":
        selected_watch_id = selected_watch.split()[0]
        st.session_state["selected_stock_id"] = selected_watch_id
        if st.sidebar.button("從觀察名單移除", key=f"remove_{selected_watch_id}"):
            remove_from_watchlist(selected_watch_id)
            st.sidebar.warning("已移除，請按 R 重新整理。")
else:
    st.sidebar.caption("目前尚未建立觀察名單。")

filtered = result_df.copy()

if selected_levels:
    filtered = filtered[filtered["買進觀察分級"].isin(selected_levels)]

if selected_markets and "市場" in filtered.columns:
    filtered = filtered[filtered["市場"].isin(selected_markets)]

if "主力突破分數" in filtered.columns:
    filtered = filtered[
        (filtered["主力突破分數"].fillna(0) >= breakout_range[0])
        & (filtered["主力突破分數"].fillna(0) <= breakout_range[1])
    ]

if keyword:
    keyword = keyword.strip()
    filtered = filtered[
        filtered["股票代號"].str.contains(keyword, case=False, na=False)
        | filtered["股票名稱"].astype(str).str.contains(keyword, case=False, na=False)
    ]

if show_only_with_charts and not chart_df.empty and "股票代號" in chart_df.columns:
    chart_ids = chart_df["股票代號"].astype(str).tolist()
    filtered = filtered[filtered["股票代號"].isin(chart_ids)]

level_counts = result_df["買進觀察分級"].value_counts().to_dict()

with st.container(border=True):
    st.subheader("🌏 市場總覽")
    market_df = get_market_dashboard_data()
    overview_cols = st.columns(8)
    overview_cols[0].metric("入選總數", len(result_df))
    overview_cols[1].metric("A級", level_counts.get("A級", 0))
    overview_cols[2].metric("B級", level_counts.get("B級", 0))
    overview_cols[3].metric("C級", level_counts.get("C級", 0))
    overview_cols[4].metric("平均突破分數", round(result_df["主力突破分數"].mean(), 2) if "主力突破分數" in result_df.columns else "-")
    overview_cols[5].metric("平均RSI", round(result_df["RSI14"].mean(), 2) if "RSI14" in result_df.columns else "-")
    overview_cols[6].metric("平均漲跌幅", f"{round(result_df['漲跌幅%'].mean(), 2)}%" if "漲跌幅%" in result_df.columns else "-")
    overview_cols[7].metric("圖表數", len(chart_df) if not chart_df.empty and "圖表路徑" in chart_df.columns else 0)
    with st.expander("查看市場風向資料", expanded=False):
        st.dataframe(market_df, use_container_width=True, hide_index=True)

st.divider()

left, right = st.columns([1.35, 1])

with left:
    st.subheader("📋 多因子掃描清單")
    display_columns = [
        col for col in [
            "日期", "市場", "股票代號", "股票名稱", "買進觀察分級", "主力突破分數",
            "收盤價", "漲跌幅%", "成交量_張", "RSI14", "分級理由", "主力突破訊號", "條件明細",
        ]
        if col in filtered.columns
    ]
    st.dataframe(filtered[display_columns], use_container_width=True, hide_index=True)

    csv = filtered.to_csv(index=False).encode("utf-8-sig")
    st.download_button("下載目前篩選清單 CSV", data=csv, file_name="台股觀察清單.csv", mime="text/csv")

with right:
    st.subheader("📊 市場結構分析")
    stat_df = result_df["買進觀察分級"].value_counts().reset_index()
    stat_df.columns = ["分級", "數量"]
    st.bar_chart(stat_df, x="分級", y="數量")

    st.subheader("📈 技術面摘要")
    tech_summary = pd.DataFrame(
        {
            "指標": ["平均 RSI", "平均 MACD", "平均突破分數", "平均成交量(張)"],
            "數值": [
                round(result_df["RSI14"].mean(), 2) if "RSI14" in result_df.columns else "-",
                round(result_df["MACD"].mean(), 2) if "MACD" in result_df.columns else "-",
                round(result_df["主力突破分數"].mean(), 2) if "主力突破分數" in result_df.columns else "-",
                int(result_df["成交量_張"].mean()) if "成交量_張" in result_df.columns else "-",
            ],
        }
    )
    st.dataframe(tech_summary, use_container_width=True, hide_index=True)

    st.subheader("⚠️ 無資料或失敗")
    if failed_df.empty or "訊息" in failed_df.columns:
        st.caption("無失敗資料。")
    else:
        st.metric("失敗檔數", len(failed_df))
        with st.expander("查看失敗清單"):
            st.dataframe(failed_df, use_container_width=True, hide_index=True)

st.divider()

st.subheader("⭐ 我的觀察名單")
watchlist_df = load_watchlist()
if watchlist_df.empty:
    st.info("尚未加入觀察股票。可從左側輸入股票代號加入。")
else:
    st.dataframe(watchlist_df, use_container_width=True, hide_index=True)

st.divider()

st.subheader("🧭 股票深度觀察")

selected_stock_id = st.session_state.get("selected_stock_id")

if selected_stock_id:
    selected_row = find_stock_row(selected_stock_id, result_df, watchlist_df)
    if selected_row is not None:
        st.success(f"目前查看：{selected_stock_id}")
        show_stock_card(selected_row, chart_df)
    else:
        st.warning("找不到這檔股票的資料。")
else:
    st.caption("尚未從觀察名單點選股票，下面先顯示目前篩選清單。")
    if filtered.empty:
        st.info("目前篩選條件下沒有股票。")
    else:
        if show_cards:
            max_cards = st.slider("顯示卡片數量", 1, min(len(filtered), 30), min(len(filtered), 10))
            for _, row in filtered.head(max_cards).iterrows():
                c1, c2 = st.columns([6, 1])
                with c1:
                    show_stock_card(row, chart_df)
                with c2:
                    if st.button("加入觀察", key=f"quick_add_{row['股票代號']}"):
                        add_to_watchlist(row["股票代號"], row["股票名稱"], row.get("市場", "-"))
                        st.success(f"已加入 {row['股票代號']} {row['股票名稱']}，請按 R 重新整理。")
        else:
            st.caption("卡片模式已關閉。")
