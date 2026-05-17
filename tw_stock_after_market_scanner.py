"""
台股盤後掃描器 v4：全上市櫃 + 買進觀察分級 + 主力突破偵測 + 自動產生圖表

功能：
1. 自動抓取上市 .TW、上櫃 .TWO 股票清單
2. 抓取日線資料
3. 計算 5MA / 10MA / 20MA / 60MA
4. 計算成交量均量
5. 計算 MACD、RSI
6. 篩選符合條件的股票
7. 加入買進觀察分級：A級 / B級 / C級 / 排除
8. 加入主力突破偵測
9. 自動產生 A/B 級股票技術圖 PNG
10. 匯出 Excel，並把圖表放入「圖表索引」工作表

安裝：
py -m pip install yfinance pandas openpyxl ta tqdm requests matplotlib

執行：
py tw_stock_after_market_scanner.py
"""

from __future__ import annotations

import time
import warnings
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
import requests
import yfinance as yf
from openpyxl.drawing.image import Image as ExcelImage
from ta.momentum import RSIIndicator
from ta.trend import MACD
from tqdm import tqdm

warnings.filterwarnings("ignore")


@dataclass
class ScannerConfig:
    period: str = "6mo"
    interval: str = "1d"
    min_volume: int = 2_000_000  # 2000張 = 2,000,000股
    output_dir: str = "output"
    chart_dir: str = "charts"
    chart_levels: tuple[str, ...] = ("A級", "B級")  # 只幫 A/B 級產圖，避免圖片太多
    max_charts: int = 30  # 最多產生幾張圖，避免 Excel 過大
    sleep_seconds: float = 0.05
    max_stocks: int | None = None  # 測試可改成 50；正式掃描用 None


CONFIG = ScannerConfig()


FALLBACK_STOCKS = [
    {"stock_id": "2330", "stock_name": "台積電", "market": "上市", "suffix": ".TW"},
    {"stock_id": "2317", "stock_name": "鴻海", "market": "上市", "suffix": ".TW"},
    {"stock_id": "2454", "stock_name": "聯發科", "market": "上市", "suffix": ".TW"},
    {"stock_id": "2303", "stock_name": "聯電", "market": "上市", "suffix": ".TW"},
    {"stock_id": "2881", "stock_name": "富邦金", "market": "上市", "suffix": ".TW"},
    {"stock_id": "2882", "stock_name": "國泰金", "market": "上市", "suffix": ".TW"},
    {"stock_id": "2891", "stock_name": "中信金", "market": "上市", "suffix": ".TW"},
    {"stock_id": "2412", "stock_name": "中華電", "market": "上市", "suffix": ".TW"},
    {"stock_id": "6488", "stock_name": "環球晶", "market": "上櫃", "suffix": ".TWO"},
    {"stock_id": "5347", "stock_name": "世界", "market": "上櫃", "suffix": ".TWO"},
]


def is_common_stock_id(stock_id: str) -> bool:
    stock_id = str(stock_id).strip()
    return stock_id.isdigit() and len(stock_id) == 4


def fetch_twse_list() -> pd.DataFrame:
    url = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L"
    response = requests.get(url, timeout=20)
    response.raise_for_status()
    data = response.json()

    rows = []
    for item in data:
        stock_id = str(item.get("Code") or item.get("公司代號") or item.get("有價證券代號") or "").strip()
        stock_name = str(item.get("Name") or item.get("公司簡稱") or item.get("有價證券名稱") or "").strip()

        if is_common_stock_id(stock_id) and stock_name:
            rows.append({"stock_id": stock_id, "stock_name": stock_name, "market": "上市", "suffix": ".TW"})

    return pd.DataFrame(rows)


def fetch_tpex_list() -> pd.DataFrame:
    urls = [
        "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_peratio_analysis",
        "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O",
    ]

    last_error: Exception | None = None

    for url in urls:
        try:
            response = requests.get(url, timeout=20)
            response.raise_for_status()
            data = response.json()

            rows = []
            for item in data:
                stock_id = str(
                    item.get("SecuritiesCompanyCode")
                    or item.get("CompanyCode")
                    or item.get("Code")
                    or item.get("公司代號")
                    or item.get("股票代號")
                    or item.get("有價證券代號")
                    or ""
                ).strip()

                stock_name = str(
                    item.get("CompanyName")
                    or item.get("Name")
                    or item.get("公司簡稱")
                    or item.get("股票名稱")
                    or item.get("有價證券名稱")
                    or ""
                ).strip()

                if is_common_stock_id(stock_id) and stock_name:
                    rows.append({"stock_id": stock_id, "stock_name": stock_name, "market": "上櫃", "suffix": ".TWO"})

            if rows:
                return pd.DataFrame(rows)

        except Exception as exc:
            last_error = exc
            continue

    raise RuntimeError(f"無法取得上櫃清單：{last_error}")


def get_all_stock_list() -> pd.DataFrame:
    frames = []

    try:
        twse = fetch_twse_list()
        if not twse.empty:
            frames.append(twse)
            print(f"上市股票清單：{len(twse)} 檔")
    except Exception as exc:
        print(f"上市清單抓取失敗：{exc}")

    try:
        tpex = fetch_tpex_list()
        if not tpex.empty:
            frames.append(tpex)
            print(f"上櫃股票清單：{len(tpex)} 檔")
    except Exception as exc:
        print(f"上櫃清單抓取失敗：{exc}")

    if not frames:
        print("官方清單抓取失敗，改用備援清單。")
        return pd.DataFrame(FALLBACK_STOCKS)

    stocks = pd.concat(frames, ignore_index=True)
    stocks = stocks.drop_duplicates(subset=["stock_id", "market"])
    stocks = stocks.sort_values(by=["market", "stock_id"]).reset_index(drop=True)

    if CONFIG.max_stocks is not None:
        stocks = stocks.head(CONFIG.max_stocks)

    return stocks


def get_yfinance_symbol(stock_id: str, suffix: str) -> str:
    return f"{stock_id}{suffix}"


def download_daily_data(stock_id: str, suffix: str) -> pd.DataFrame:
    symbol = get_yfinance_symbol(stock_id, suffix)
    df = yf.download(
        symbol,
        period=CONFIG.period,
        interval=CONFIG.interval,
        auto_adjust=False,
        progress=False,
        threads=False,
    )

    if df.empty:
        return pd.DataFrame()

    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)

    df = df.reset_index()
    df = df.rename(
        columns={
            "Date": "date",
            "Open": "open",
            "High": "high",
            "Low": "low",
            "Close": "close",
            "Adj Close": "adj_close",
            "Volume": "volume",
        }
    )

    required_columns = {"date", "open", "high", "low", "close", "volume"}
    if not required_columns.issubset(df.columns):
        return pd.DataFrame()

    df = df.dropna(subset=["open", "high", "low", "close", "volume"])
    return df


def add_indicators(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df["ma5"] = df["close"].rolling(5).mean()
    df["ma10"] = df["close"].rolling(10).mean()
    df["ma20"] = df["close"].rolling(20).mean()
    df["ma60"] = df["close"].rolling(60).mean()

    df["vol_ma5"] = df["volume"].rolling(5).mean()
    df["vol_ma20"] = df["volume"].rolling(20).mean()

    macd = MACD(close=df["close"], window_slow=26, window_fast=12, window_sign=9)
    df["macd"] = macd.macd()
    df["macd_signal"] = macd.macd_signal()
    df["macd_hist"] = macd.macd_diff()

    rsi = RSIIndicator(close=df["close"], window=14)
    df["rsi14"] = rsi.rsi()

    df["pct_change"] = df["close"].pct_change() * 100

    df["high_20"] = df["high"].rolling(20).max()
    df["high_60"] = df["high"].rolling(60).max()
    df["prev_high_20"] = df["high"].shift(1).rolling(20).max()
    df["prev_high_60"] = df["high"].shift(1).rolling(60).max()
    df["volume_up_3days"] = (df["volume"] > df["volume"].shift(1)) & (df["volume"].shift(1) > df["volume"].shift(2))
    df["ma_converge"] = (
        (abs(df["ma5"] - df["ma20"]) / df["ma20"] <= 0.03)
        & (abs(df["ma10"] - df["ma20"]) / df["ma20"] <= 0.03)
    )

    return df


def detect_major_breakout(latest: pd.Series, prev: pd.Series) -> tuple[int, str]:
    signals = []

    close = latest["close"]
    volume = latest["volume"]
    vol_ma5 = latest["vol_ma5"]
    vol_ma20 = latest["vol_ma20"]
    ma20 = latest["ma20"]
    prev_high_20 = latest.get("prev_high_20")
    prev_high_60 = latest.get("prev_high_60")

    if pd.notna(prev_high_20) and close > prev_high_20:
        signals.append("收盤突破20日高點")

    if pd.notna(prev_high_60) and close > prev_high_60:
        signals.append("收盤突破60日高點")

    if volume > vol_ma20 * 1.8:
        signals.append("成交量大於20日均量1.8倍")
    elif volume > vol_ma20 * 1.4:
        signals.append("成交量大於20日均量1.4倍")

    if volume > vol_ma5 * 1.5:
        signals.append("成交量大於5日均量1.5倍")

    if bool(latest.get("volume_up_3days")):
        signals.append("連續3日增量")

    if bool(prev.get("ma_converge")) and close > ma20 and volume > vol_ma20 * 1.4:
        signals.append("均線糾結後放量轉強")

    if close > latest["open"] and close >= latest["low"] + (latest["high"] - latest["low"]) * 0.75:
        signals.append("紅K且收盤接近高點")

    score = len(signals)
    return score, "、".join(signals) if signals else "無明顯主力突破訊號"


def classify_watch_level(latest: pd.Series, prev: pd.Series, conditions: dict, breakout_score: int, breakout_reason: str) -> tuple[str, str]:
    close = latest["close"]
    open_ = latest["open"]
    high = latest["high"]
    low = latest["low"]
    ma5 = latest["ma5"]
    ma10 = latest["ma10"]
    ma20 = latest["ma20"]
    ma60 = latest["ma60"]
    volume = latest["volume"]
    vol_ma5 = latest["vol_ma5"]
    vol_ma20 = latest["vol_ma20"]
    rsi = latest["rsi14"]
    macd = latest["macd"]
    macd_signal = latest["macd_signal"]
    pct_change = latest["pct_change"]

    body = abs(close - open_)
    full_range = max(high - low, 0.01)
    upper_shadow = high - max(close, open_)

    above_all_ma = close > ma5 > ma10 > ma20 > ma60
    ma_bullish = ma5 > ma10 > ma20
    volume_expansion = volume > vol_ma5 and volume > vol_ma20
    macd_bullish = macd > macd_signal and macd > 0
    healthy_rsi = 50 <= rsi <= 75
    strong_close = close > open_ and close >= low + full_range * 0.65
    near_high_close = close >= low + full_range * 0.75

    too_hot = rsi > 82 or close > ma20 * 1.25
    long_black = close < open_ and abs(pct_change) >= 4 and body / full_range >= 0.55
    long_upper_shadow = upper_shadow / full_range >= 0.45 and close < high * 0.97
    break_ma5 = close < ma5
    weak_macd = macd < macd_signal and latest["macd_hist"] < prev["macd_hist"]

    risk_reasons = []
    if too_hot:
        risk_reasons.append("乖離或RSI過熱")
    if long_black:
        risk_reasons.append("長黑K風險")
    if long_upper_shadow:
        risk_reasons.append("上影線偏長")
    if break_ma5:
        risk_reasons.append("跌破5MA")
    if weak_macd:
        risk_reasons.append("MACD動能轉弱")

    if risk_reasons and (too_hot or long_black or break_ma5):
        return "排除", "、".join(risk_reasons)

    if above_all_ma and volume_expansion and macd_bullish and healthy_rsi and strong_close and breakout_score >= 3:
        return "A級", f"強勢多頭排列、量能放大、MACD偏多、RSI健康、主力突破訊號：{breakout_reason}"

    if above_all_ma and volume_expansion and healthy_rsi and breakout_score >= 4:
        return "A級", f"量價突破強，主力突破訊號：{breakout_reason}"

    if ma_bullish and close > ma20 and conditions["20MA向上"] and conditions["成交量大於2000張"] and rsi > 50:
        reasons = ["均線趨勢偏多", "收盤站上20MA", "量能達標", "RSI站上50"]
        if macd_bullish:
            reasons.append("MACD偏多")
        if near_high_close:
            reasons.append("收盤接近高點")
        if breakout_score >= 2:
            reasons.append("主力突破訊號：" + breakout_reason)
        if risk_reasons:
            reasons.append("但" + "、".join(risk_reasons))
        return "B級", "、".join(reasons)

    return "C級", "初步轉強，可列入觀察，但尚未形成強勢突破"


def check_signal(df: pd.DataFrame) -> dict | None:
    if len(df) < 60:
        return None

    latest = df.iloc[-1]
    prev = df.iloc[-2]

    if pd.isna(latest["ma20"]) or pd.isna(latest["ma60"]) or pd.isna(latest["rsi14"]):
        return None

    conditions = {
        "收盤站上20MA": latest["close"] > latest["ma20"],
        "20MA向上": latest["ma20"] > prev["ma20"],
        "成交量大於2000張": latest["volume"] >= CONFIG.min_volume,
        "量大於5日均量": latest["volume"] > latest["vol_ma5"],
        "MACD黃金交叉": prev["macd"] <= prev["macd_signal"] and latest["macd"] > latest["macd_signal"],
        "RSI大於50": latest["rsi14"] > 50,
        "收紅K": latest["close"] > latest["open"],
    }

    score = sum(conditions.values())

    if not (conditions["收盤站上20MA"] and conditions["成交量大於2000張"] and score >= 4):
        return None

    breakout_score, breakout_reason = detect_major_breakout(latest, prev)
    watch_level, watch_reason = classify_watch_level(latest, prev, conditions, breakout_score, breakout_reason)

    return {
        "日期": latest["date"].strftime("%Y-%m-%d") if hasattr(latest["date"], "strftime") else latest["date"],
        "收盤價": round(float(latest["close"]), 2),
        "漲跌幅%": round(float(latest["pct_change"]), 2),
        "成交量_張": int(latest["volume"] // 1000),
        "5MA": round(float(latest["ma5"]), 2),
        "10MA": round(float(latest["ma10"]), 2),
        "20MA": round(float(latest["ma20"]), 2),
        "60MA": round(float(latest["ma60"]), 2),
        "MACD": round(float(latest["macd"]), 3),
        "MACD_signal": round(float(latest["macd_signal"]), 3),
        "RSI14": round(float(latest["rsi14"]), 2),
        "符合條件數": score,
        "主力突破分數": breakout_score,
        "主力突破訊號": breakout_reason,
        "買進觀察分級": watch_level,
        "分級理由": watch_reason,
        "條件明細": "、".join([k for k, v in conditions.items() if v]),
    }


def scan_market() -> tuple[pd.DataFrame, pd.DataFrame]:
    stocks = get_all_stock_list()
    print(f"本次預計掃描：{len(stocks)} 檔")

    results = []
    failed = []

    for row in tqdm(stocks.itertuples(index=False), total=len(stocks), desc="掃描上市櫃"):
        stock_id = row.stock_id
        stock_name = row.stock_name
        market = row.market
        suffix = row.suffix

        try:
            df = download_daily_data(stock_id, suffix)
            if df.empty:
                failed.append({"股票代號": stock_id, "股票名稱": stock_name, "市場": market, "原因": "無資料"})
                continue

            df = add_indicators(df)
            signal = check_signal(df)

            if signal:
                signal["股票代號"] = stock_id
                signal["股票名稱"] = stock_name
                signal["市場"] = market
                signal["yfinance代碼"] = get_yfinance_symbol(stock_id, suffix)
                results.append(signal)

            time.sleep(CONFIG.sleep_seconds)

        except Exception as exc:
            failed.append({"股票代號": stock_id, "股票名稱": stock_name, "市場": market, "原因": str(exc)})

    result_df = pd.DataFrame(results)
    failed_df = pd.DataFrame(failed)

    if not result_df.empty:
        columns = [
            "日期", "市場", "股票代號", "股票名稱", "yfinance代碼",
            "買進觀察分級", "分級理由", "主力突破分數", "主力突破訊號",
            "收盤價", "漲跌幅%", "成交量_張", "5MA", "10MA", "20MA", "60MA",
            "MACD", "MACD_signal", "RSI14", "符合條件數", "條件明細",
        ]
        result_df = result_df[columns]
        level_order = {"A級": 1, "B級": 2, "C級": 3, "排除": 9}
        result_df["分級排序"] = result_df["買進觀察分級"].map(level_order).fillna(8)
        result_df = result_df.sort_values(
            by=["分級排序", "主力突破分數", "符合條件數", "成交量_張"],
            ascending=[True, False, False, False],
        )
        result_df = result_df.drop(columns=["分級排序"])

    return result_df, failed_df


def save_stock_chart(stock_id: str, stock_name: str, market: str, suffix: str, output_dir: Path) -> Path | None:
    try:
        df = download_daily_data(stock_id, suffix)
        if df.empty or len(df) < 60:
            return None

        df = add_indicators(df).tail(90).reset_index(drop=True)
        df["date_label"] = pd.to_datetime(df["date"]).dt.strftime("%m-%d")
        x = range(len(df))

        output_dir.mkdir(parents=True, exist_ok=True)
        chart_path = output_dir / f"{stock_id}_{stock_name}.png"

        fig = plt.figure(figsize=(11, 8))
        ax_price = fig.add_axes([0.08, 0.58, 0.86, 0.34])
        ax_volume = fig.add_axes([0.08, 0.43, 0.86, 0.12], sharex=ax_price)
        ax_macd = fig.add_axes([0.08, 0.24, 0.86, 0.14], sharex=ax_price)
        ax_rsi = fig.add_axes([0.08, 0.08, 0.86, 0.11], sharex=ax_price)

        ax_price.plot(x, df["close"], label="Close", linewidth=1.5)
        ax_price.plot(x, df["ma5"], label="MA5", linewidth=1)
        ax_price.plot(x, df["ma20"], label="MA20", linewidth=1)
        ax_price.plot(x, df["ma60"], label="MA60", linewidth=1)
        ax_price.set_title(f"{stock_id} {stock_name} | {market} | 技術圖")
        ax_price.legend(loc="upper left", fontsize=9)
        ax_price.grid(True, alpha=0.25)

        ax_volume.bar(x, df["volume"] / 1000, label="Volume(張)")
        ax_volume.plot(x, df["vol_ma20"] / 1000, label="Vol MA20", linewidth=1)
        ax_volume.legend(loc="upper left", fontsize=8)
        ax_volume.grid(True, alpha=0.25)

        ax_macd.plot(x, df["macd"], label="MACD", linewidth=1)
        ax_macd.plot(x, df["macd_signal"], label="Signal", linewidth=1)
        ax_macd.bar(x, df["macd_hist"], label="Hist")
        ax_macd.axhline(0, linewidth=0.8)
        ax_macd.legend(loc="upper left", fontsize=8)
        ax_macd.grid(True, alpha=0.25)

        ax_rsi.plot(x, df["rsi14"], label="RSI14", linewidth=1)
        ax_rsi.axhline(70, linewidth=0.8, linestyle="--")
        ax_rsi.axhline(50, linewidth=0.8, linestyle="--")
        ax_rsi.axhline(30, linewidth=0.8, linestyle="--")
        ax_rsi.set_ylim(0, 100)
        ax_rsi.legend(loc="upper left", fontsize=8)
        ax_rsi.grid(True, alpha=0.25)

        tick_step = max(len(df) // 8, 1)
        tick_positions = list(range(0, len(df), tick_step))
        ax_rsi.set_xticks(tick_positions)
        ax_rsi.set_xticklabels(df.loc[tick_positions, "date_label"], rotation=45)

        plt.setp(ax_price.get_xticklabels(), visible=False)
        plt.setp(ax_volume.get_xticklabels(), visible=False)
        plt.setp(ax_macd.get_xticklabels(), visible=False)

        fig.savefig(chart_path, dpi=130, bbox_inches="tight")
        plt.close(fig)
        return chart_path

    except Exception as exc:
        print(f"圖表產生失敗：{stock_id} {stock_name}，原因：{exc}")
        return None


def generate_charts(result_df: pd.DataFrame) -> pd.DataFrame:
    if result_df.empty:
        return pd.DataFrame()

    chart_output_dir = Path(CONFIG.output_dir) / CONFIG.chart_dir / datetime.now().strftime("%Y%m%d")
    targets = result_df[result_df["買進觀察分級"].isin(CONFIG.chart_levels)].head(CONFIG.max_charts)

    chart_rows = []
    print(f"\n開始產生圖表：{len(targets)} 張")

    for row in tqdm(targets.itertuples(index=False), total=len(targets), desc="產生圖表"):
        suffix = ".TW" if row.市場 == "上市" else ".TWO"
        chart_path = save_stock_chart(row.股票代號, row.股票名稱, row.市場, suffix, chart_output_dir)
        if chart_path:
            chart_rows.append(
                {
                    "股票代號": row.股票代號,
                    "股票名稱": row.股票名稱,
                    "市場": row.市場,
                    "買進觀察分級": row.買進觀察分級,
                    "主力突破分數": row.主力突破分數,
                    "圖表路徑": str(chart_path),
                }
            )

    return pd.DataFrame(chart_rows)


def save_to_excel(result_df: pd.DataFrame, failed_df: pd.DataFrame, chart_df: pd.DataFrame | None = None) -> Path:
    output_dir = Path(CONFIG.output_dir)
    output_dir.mkdir(exist_ok=True)

    now_text = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = output_dir / f"台股全上市櫃盤後掃描_{now_text}.xlsx"

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        if result_df.empty:
            pd.DataFrame([{"訊息": "今日沒有股票符合條件"}]).to_excel(writer, index=False, sheet_name="掃描結果")
        else:
            result_df.to_excel(writer, index=False, sheet_name="掃描結果")

        if failed_df.empty:
            pd.DataFrame([{"訊息": "無失敗資料"}]).to_excel(writer, index=False, sheet_name="無資料或失敗")
        else:
            failed_df.to_excel(writer, index=False, sheet_name="無資料或失敗")

        if chart_df is not None and not chart_df.empty:
            chart_df.to_excel(writer, index=False, sheet_name="圖表索引")
            chart_ws = writer.sheets["圖表索引"]
            chart_ws.column_dimensions["F"].width = 90

            start_row = len(chart_df) + 4
            chart_ws.cell(row=start_row - 1, column=1, value="技術圖表")

            for idx, chart_row in enumerate(chart_df.itertuples(index=False), start=0):
                image_path = Path(chart_row.圖表路徑)
                if image_path.exists():
                    img = ExcelImage(str(image_path))
                    img.width = 760
                    img.height = 520
                    anchor_cell = f"A{start_row + idx * 28}"
                    chart_ws.add_image(img, anchor_cell)
                    chart_ws.cell(
                        row=start_row + idx * 28 - 1,
                        column=1,
                        value=f"{chart_row.股票代號} {chart_row.股票名稱}｜{chart_row.買進觀察分級}",
                    )
        else:
            pd.DataFrame([{"訊息": "本次沒有產生圖表"}]).to_excel(writer, index=False, sheet_name="圖表索引")

        for worksheet in writer.sheets.values():
            for column_cells in worksheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter
                for cell in column_cells:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                worksheet.column_dimensions[column_letter].width = min(max_length + 2, 60)

    return file_path


def main() -> None:
    print("開始執行台股全上市櫃盤後掃描器...")
    result_df, failed_df = scan_market()

    if result_df.empty:
        print("今日沒有股票符合條件。")
    else:
        print("\n掃描結果：")
        print(result_df.to_string(index=False))

    print(f"\n無資料或失敗檔數：{len(failed_df)}")

    chart_df = generate_charts(result_df)
    file_path = save_to_excel(result_df, failed_df, chart_df)

    print(f"\n已輸出 Excel：{file_path}")
    if not chart_df.empty:
        print(f"已產生圖表：{len(chart_df)} 張")


if __name__ == "__main__":
    main()
