#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
B站汽车赛道内容数据分析：数据清洗 + 分析 + 图表输出

输入：市场规模+作者规模+比亚迪+竞品（原版+清洗版）.xlsx
      6 个 sheet：市场规模 / 作者规模 / 比亚迪 / 自身比亚迪 / 竞品特斯拉 / 竞品蔚来

输出：
  output/清洗数据.xlsx        每个 sheet 的清洗后数据
  output/分析结果.xlsx        各类汇总分析表
  output/charts/*.png         图表

用法：
  python analyze_bilibili.py [输入文件路径]
  默认读取同目录下的 xlsx 文件，结果写到同目录 output/ 下。

依赖：pandas openpyxl matplotlib
      pip install pandas openpyxl matplotlib
"""

import os
import re
import sys
import warnings

import pandas as pd
import numpy as np

warnings.filterwarnings("ignore")

# 中文/西文混排字体，避免图表中文变方框
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams["font.sans-serif"] = [
    "Noto Sans CJK JP", "Noto Sans CJK SC", "WenQuanYi Zen Hei",
    "SimHei", "Microsoft YaHei", "Arial Unicode MS", "DejaVu Sans",
]
plt.rcParams["axes.unicode_minus"] = False

DEFAULT_FILE = "市场规模+作者规模+比亚迪+竞品（原版+清洗版）.xlsx"


# ============================================================
# 一、通用清洗工具
# ============================================================

# 单位换算表：匹配「数字 + 单位」，支持 万/亿/w/k
_UNIT_MAP = {
    "亿": 1e8, "萬": 1e4, "万": 1e4, "w": 1e4, "W": 1e4,
    "k": 1e3, "K": 1e3,
}
_NUM_UNIT_RE = re.compile(r"(-?\d+(?:\.\d+)?)\s*(亿|万|萬|w|W|k|K)?")


def to_num(value):
    """
    把 "4.3万" / "1.1亿" / "43000" / "1,234" / "-" / "" 转成数值。
    无法解析返回 np.nan。

    >>> to_num("4.3万")
    43000.0
    >>> to_num("1.1亿")
    110000000.0
    >>> to_num("1,234")
    1234.0
    >>> to_num("-")
    nan
    """
    if value is None:
        return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        return np.nan if pd.isna(value) else float(value)

    text = str(value).strip()
    if text in ("", "-", "--", "—", "/", "暂无", "nan", "None", "NaN"):
        return np.nan

    text = text.replace(",", "").replace("，", "").replace(" ", "")
    match = _NUM_UNIT_RE.search(text)
    if not match:
        return np.nan
    number, unit = match.group(1), match.group(2)
    try:
        result = float(number)
    except ValueError:
        return np.nan
    if unit:
        result *= _UNIT_MAP.get(unit, 1)
    return result


def duration_to_seconds(value):
    """
    把 "0:22:13" / "22:13" / "1333" 转成秒数。

    >>> duration_to_seconds("0:22:13")
    1333
    >>> duration_to_seconds("22:13")
    1333
    >>> duration_to_seconds(1333)
    1333
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        return int(value)

    text = str(value).strip()
    if not text or text in ("-", "--"):
        return np.nan

    # 纯数字字符串，直接当秒
    if re.fullmatch(r"\d+(\.\d+)?", text):
        return int(float(text))

    parts = text.split(":")
    try:
        parts = [int(float(p)) for p in parts]
    except ValueError:
        return np.nan
    seconds = 0
    for part in parts:            # 从高位到低位累加：时、分、秒
        seconds = seconds * 60 + part
    return seconds


def parse_datetime(series):
    """统一时间格式：优先按字符串解析，失败再按 Excel 序列号处理。"""
    parsed = pd.to_datetime(series, errors="coerce", format="mixed")
    if parsed.isna().mean() > 0.5:
        parsed = pd.to_datetime(series, errors="coerce", unit="D", origin="1899-12-30")
    return parsed


def clean_text(series):
    """文本列：去首尾空格、连续空白压成单空格、去换行。"""
    return (
        series.astype("string")
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )


# ============================================================
# 二、各 sheet 的清洗逻辑
# ============================================================

def clean_market(df: pd.DataFrame, log: dict) -> pd.DataFrame:
    """市场规模：视频级数据（播放量、弹幕量、时长、发布时间）。"""
    before = len(df)
    df = df.copy()

    # 原始列可能是 "4.3万" 文本，统一转数值；若已有新*数值列则优先用
    df["播放量_num"] = df["播放量"].map(to_num) if "播放量" in df else np.nan
    df["新播放量_num"] = df["新播放量"].map(to_num) if "新播放量" in df else np.nan
    df["播放量_最终"] = df["新播放量_num"].fillna(df["播放量_num"])

    df["弹幕量_num"] = df["弹幕量"].map(to_num) if "弹幕量" in df else np.nan
    df["新弹幕量_num"] = df["新弹幕量"].map(to_num) if "新弹幕量" in df else np.nan
    df["弹幕量_最终"] = df["新弹幕量_num"].fillna(df["弹幕量_num"])

    df["时长_秒"] = df["新视频时长"].map(duration_to_seconds) if "新视频时长" in df else np.nan
    df["时长_秒"] = df["时长_秒"].fillna(
        df["视频时长"].map(duration_to_seconds) if "视频时长" in df else np.nan
    )
    df["时长_分钟"] = (df["时长_秒"] / 60).round(1)

    df["发布时间_dt"] = parse_datetime(df["发布时间"]) if "发布时间" in df else pd.NaT
    df["发布年月"] = df["发布时间_dt"].dt.to_period("M").astype("string")
    df["发布时间_日期"] = df["发布时间_dt"].dt.date.astype("string")

    df["标题"] = clean_text(df["标题"]) if "标题" in df else df.get("标题")
    df["作者"] = clean_text(df["作者"]) if "作者" in df else df.get("作者")

    # 互动率（弹幕/播放）与完播热度指标
    df["互动率"] = np.where(
        df["播放量_最终"] > 0, df["弹幕量_最终"] / df["播放量_最终"], np.nan
    )
    df["互动率"] = df["互动率"].round(5)

    # 去重：同一标题+同一作者视为重复（保留播放量最高的一条）
    dup_mask = df.duplicated(subset=["标题", "作者"], keep=False)
    dup_count = int(df.duplicated(subset=["标题", "作者"], keep="first").sum())
    if dup_count:
        df = (
            df.sort_values("播放量_最终", ascending=False, na_position="last")
            .drop_duplicates(subset=["标题", "作者"], keep="first")
        )

    # 丢弃核心指标全空的行
    df = df.dropna(subset=["播放量_最终"], how="all")

    log["市场规模"] = {
        "原始行数": before,
        "去重删除": dup_count,
        "重复组涉及行数": int(dup_mask.sum()),
        "清洗后行数": len(df),
        "播放量缺失": int(df["播放量_最终"].isna().sum()),
        "时长缺失": int(df["时长_秒"].isna().sum()),
        "发布时间解析失败": int(df["发布时间_dt"].isna().sum()),
        "作者数": int(df["作者"].nunique()),
    }
    return df.reset_index(drop=True)


def clean_author(df: pd.DataFrame, log: dict) -> pd.DataFrame:
    """作者规模：作者级数据（粉丝、获赞、播放）。"""
    before = len(df)
    df = df.copy()

    for src, dst in [
        ("新粉丝数", "粉丝数"),
        ("新播放数", "播放总数"),
        ("新点赞数", "获赞数"),
        ("新视频数", "视频数"),
        ("新关注数", "关注数"),
    ]:
        raw = df[src].map(to_num) if src in df else np.nan
        backup = df[dst].map(to_num) if dst in df else np.nan
        df[dst + "_最终"] = pd.Series(raw, index=df.index).fillna(pd.Series(backup, index=df.index))

    df["会员等级"] = pd.to_numeric(df.get("会员等级"), errors="coerce")
    df["作者名"] = clean_text(df["作者名"]) if "作者名" in df else df.get("作者名")

    dup = int(df.duplicated(subset=["作者名"], keep="first").sum())
    if dup:
        df = df.sort_values("粉丝数_最终", ascending=False, na_position="last")
        df = df.drop_duplicates(subset=["作者名"], keep="first")

    df["粉丝播放比"] = np.where(
        df["粉丝数_最终"] > 0, df["播放总数_最终"] / df["粉丝数_最终"], np.nan
    ).round(2)

    # 粉丝分层
    bins = [-1, 1e4, 1e5, 5e5, 1e6, 1e7, np.inf]
    labels = ["1万以下", "1-10万", "10-50万", "50-100万", "100-1000万", "1000万以上"]
    df["粉丝层级"] = pd.cut(df["粉丝数_最终"], bins=bins, labels=labels)

    log["作者规模"] = {
        "原始行数": before,
        "去重删除": dup,
        "清洗后行数": len(df),
        "粉丝数缺失": int(df["粉丝数_最终"].isna().sum()),
    }
    return df.reset_index(drop=True)


BRAND_SHEET_COLUMNS = {
    "比亚迪": {"play": "新播放数", "danmu": "新弹幕数", "like": "点赞数",
               "coin": "金币数", "fav": "收藏量", "share": "分享数",
               "pub": "发布时间", "author": "发布者", "time": "时长"},
    "自身比亚迪": {"play": "新播放量", "danmu": "新弹幕数", "like": "新点赞数",
                   "coin": "新投币数", "fav": "新收藏量", "share": "新分享数",
                   "pub": "发布时间", "author": "发布者", "time": "视频时长"},
    "竞品特斯拉": {"play": "新播放数", "danmu": "新弹幕数", "like": "点赞数",
                   "coin": "金币数", "fav": "收藏量", "share": "分享数",
                   "pub": "发布时间", "author": "发布者", "time": "时长"},
    "竞品蔚来": {"play": "新播放数", "danmu": "新弹幕数", "like": "点赞数",
                 "coin": "金币数", "fav": "收藏量", "share": "分享数",
                 "pub": "发布时间", "author": "发布者", "time": "时长"},
}


def clean_brand(df: pd.DataFrame, sheet: str, log: dict) -> pd.DataFrame:
    """品牌 sheet：统一成 play / danmu / like / coin / fav / share / pub / author 标准列。"""
    before = len(df)
    df = df.copy()
    cols = BRAND_SHEET_COLUMNS[sheet]
    out = pd.DataFrame(index=df.index)

    out["品牌"] = sheet.replace("自身", "").replace("竞品", "")
    out["标题"] = clean_text(df["标题"]) if "标题" in df else ""

    for dst, src in [("play", "play"), ("danmu", "danmu"), ("like", "like"),
                     ("coin", "coin"), ("fav", "fav"), ("share", "share")]:
        out[dst] = df[cols[src]].map(to_num) if cols[src] in df else np.nan

    out["时长_秒"] = df[cols["time"]].map(duration_to_seconds) if cols["time"] in df else np.nan
    out["时长_分钟"] = (out["时长_秒"] / 60).round(1)
    out["发布时间_dt"] = parse_datetime(df[cols["pub"]]) if cols["pub"] in df else pd.NaT
    out["发布年月"] = out["发布时间_dt"].dt.to_period("M").astype("string")
    out["发布者"] = clean_text(df[cols["author"]]) if cols["author"] in df else ""

    # 互动率：非播放类互动 / 播放
    out["互动率"] = np.where(
        out["play"] > 0, (out["like"] + out["coin"] + out["fav"] + out["share"]) / out["play"], np.nan
    ).round(5)
    out["弹幕率"] = np.where(out["play"] > 0, out["danmu"] / out["play"], np.nan).round(5)

    dup = int(out.duplicated(subset=["标题"], keep="first").sum())
    if dup:
        out = (
            out.sort_values("play", ascending=False, na_position="last")
            .drop_duplicates(subset=["标题"], keep="first")
        )

    log[sheet] = {
        "原始行数": before,
        "去重删除": dup,
        "清洗后行数": len(out),
        "播放量缺失": int(out["play"].isna().sum()),
        "发布时间解析失败": int(out["发布时间_dt"].isna().sum()),
    }
    return out.reset_index(drop=True)


# ============================================================
# 三、分析
# ============================================================

def analyze(market: pd.DataFrame, authors: pd.DataFrame, brands: pd.DataFrame) -> dict:
    """产出各维度分析表，返回 {sheet名: DataFrame}。"""
    res = {}

    # 1. 市场大盘概况
    res["01_市场大盘"] = pd.DataFrame(
        {
            "指标": ["视频数", "作者数", "总播放量", "总弹幕量", "平均播放量", "播放量中位数",
                     "平均时长(分钟)", "平均互动率", "时间范围"],
            "数值": [
                len(market),
                market["作者"].nunique(),
                round(market["播放量_最终"].sum(), 0),
                round(market["弹幕量_最终"].sum(), 0),
                round(market["播放量_最终"].mean(), 0),
                round(market["播放量_最终"].median(), 0),
                round(market["时长_分钟"].mean(), 1),
                round(market["互动率"].mean(), 4),
                f"{market['发布时间_dt'].min():%Y-%m-%d} ~ {market['发布时间_dt'].max():%Y-%m-%d}",
            ],
        }
    )

    # 2. 播放量分布（分层）
    bins = [-1, 1e4, 5e4, 1e5, 5e5, 1e6, 1e7, np.inf]
    labels = ["1万以下", "1-5万", "5-10万", "10-50万", "50-100万", "100-1000万", "1000万以上"]
    market["播放层级"] = pd.cut(market["播放量_最终"], bins=bins, labels=labels)
    dist = market["播放层级"].value_counts().reindex(labels).fillna(0).astype(int)
    res["02_播放量分布"] = pd.DataFrame(
        {"播放区间": labels, "视频数": dist.values,
         "占比": (dist.values / len(market)).round(4)}
    )

    # 3. 时长与播放关系（时长分段）
    tbins = [0, 60, 180, 300, 600, 900, 1800, np.inf]
    tlabels = ["1分钟内", "1-3分钟", "3-5分钟", "5-10分钟", "10-15分钟", "15-30分钟", "30分钟以上"]
    market["时长分层"] = pd.cut(market["时长_秒"], bins=tbins, labels=tlabels)
    tstat = (
        market.groupby("时长分层", observed=True)
        .agg(视频数=("播放量_最终", "size"),
             平均播放=("播放量_最终", "mean"),
             播放中位数=("播放量_最终", "median"),
             平均互动率=("互动率", "mean"))
        .round(2).reset_index()
        .rename(columns={"时长分层": "时长区间"})
    )
    res["03_时长效应"] = tstat

    # 4. 月度趋势
    monthly = (
        market.dropna(subset=["发布年月"])
        .groupby("发布年月")
        .agg(视频数=("播放量_最终", "size"),
             总播放=("播放量_最终", "sum"),
             平均播放=("播放量_最终", "mean"))
        .round(0).reset_index()
    )
    res["04_月度趋势"] = monthly

    # 5. TOP20 视频
    top20 = market.nlargest(20, "播放量_最终")[
        ["标题", "作者", "播放量_最终", "弹幕量_最终", "互动率", "时长_分钟", "发布时间_日期"]
    ].rename(columns={"播放量_最终": "播放量", "弹幕量_最终": "弹幕量", "发布时间_日期": "发布时间"})
    res["05_TOP20视频"] = top20

    # 6. TOP20 作者（按视频总播放）
    top_author = (
        market.groupby("作者")
        .agg(视频数=("播放量_最终", "size"),
             总播放=("播放量_最终", "sum"),
             平均播放=("播放量_最终", "mean"))
        .sort_values("总播放", ascending=False)
        .head(20).round(0).reset_index()
    )
    res["06_TOP20作者"] = top_author

    # 7. 品牌对比（比亚迪 / 特斯拉 / 蔚来）
    if not brands.empty:
        brand_cmp = (
            brands.groupby("品牌")
            .agg(视频数=("play", "size"),
                 总播放=("play", "sum"),
                 平均播放=("play", "mean"),
                 播放中位数=("play", "median"),
                 平均点赞=("like", "mean"),
                 平均投币=("coin", "mean"),
                 平均收藏=("fav", "mean"),
                 平均分享=("share", "mean"),
                 平均弹幕率=("弹幕率", "mean"),
                 平均互动率=("互动率", "mean"),
                 平均时长分钟=("时长_分钟", "mean"))
            .round(3).reset_index()
            .sort_values("总播放", ascending=False)
        )
        res["07_品牌对比"] = brand_cmp

        brand_top = (
            brands.sort_values("play", ascending=False)
            .groupby("品牌").head(5)[["品牌", "标题", "play", "like", "互动率", "发布年月"]]
            .rename(columns={"play": "播放量"})
        )
        res["08_品牌TOP5视频"] = brand_top

        # 品牌月度声量
        if brands["发布年月"].notna().any():
            res["09_品牌月度趋势"] = (
                brands.dropna(subset=["发布年月"])
                .groupby(["发布年月", "品牌"])
                .agg(视频数=("play", "size"), 总播放=("play", "sum"))
                .round(0).reset_index()
            )

    # 8. 作者规模分布
    if not authors.empty:
        res["10_作者粉丝分层"] = (
            authors["粉丝层级"].value_counts().reindex(
                ["1万以下", "1-10万", "10-50万", "50-100万", "100-1000万", "1000万以上"]
            ).fillna(0).astype(int)
            .rename_axis("粉丝层级").reset_index(name="作者数")
        )
        res["11_头部作者TOP20"] = (
            authors.nlargest(20, "粉丝数_最终")[
                ["作者名", "粉丝数_最终", "获赞数_最终", "视频数_最终", "播放总数_最终", "会员等级"]
            ].rename(columns={"粉丝数_最终": "粉丝数", "获赞数_最终": "获赞数",
                              "视频数_最终": "视频数", "播放总数_最终": "播放总数"})
        )

    return res


# ============================================================
# 四、图表
# ============================================================

def plot_all(market, authors, brands, results, chart_dir):
    os.makedirs(chart_dir, exist_ok=True)
    saved = []

    def _save(fig, name):
        path = os.path.join(chart_dir, name)
        fig.tight_layout()
        fig.savefig(path, dpi=150, bbox_inches="tight")
        plt.close(fig)
        saved.append(path)

    # 图1 播放量分布
    d = results["02_播放量分布"]
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.bar(d["播放区间"], d["视频数"], color="#4C78A8")
    ax.set_title("视频播放量分层分布")
    ax.set_ylabel("视频数")
    ax.tick_params(axis="x", rotation=30)
    for i, v in enumerate(d["视频数"]):
        ax.text(i, v, str(v), ha="center", va="bottom", fontsize=8)
    _save(fig, "01_播放量分布.png")

    # 图2 时长与平均播放
    d = results["03_时长效应"].dropna(subset=["平均播放"])
    if not d.empty:
        fig, ax = plt.subplots(figsize=(8, 4.5))
        ax.bar(d["时长区间"], d["平均播放"] / 1e4, color="#F58518")
        ax.set_title("不同时长区间的平均播放量")
        ax.set_ylabel("平均播放（万）")
        ax.tick_params(axis="x", rotation=30)
        _save(fig, "02_时长效应.png")

    # 图3 月度趋势
    d = results["04_月度趋势"]
    if not d.empty:
        fig, ax = plt.subplots(figsize=(14, 5.5))
        x = np.arange(len(d))                      # 用连续序号画图，避免月份字符串挤在一起
        ax.plot(x, d["视频数"], marker="o", markersize=4, label="视频数")
        ax2 = ax.twinx()
        ax2.plot(x, d["平均播放"] / 1e4, marker="s", markersize=4,
                 color="#E45756", label="平均播放（万）")
        ax.set_title("月度发布量与平均播放趋势")
        ax.set_ylabel("视频数")
        ax2.set_ylabel("平均播放（万）")
        # X 轴最多显示 12 个月份标签，间隔均匀、斜 45°，避免上百个标签重叠
        step = max(1, len(d) // 12)
        ax.set_xticks(x[::step])
        ax.set_xticklabels(d["发布年月"].astype(str)[::step], rotation=45, ha="right")
        ax.legend(loc="upper left", fontsize=9)
        ax2.legend(loc="upper right", fontsize=9)
        _save(fig, "03_月度趋势.png")

    # 图4 品牌对比
    if "07_品牌对比" in results:
        d = results["07_品牌对比"]
        fig, axes = plt.subplots(1, 2, figsize=(11, 4.5))
        axes[0].bar(d["品牌"], d["总播放"] / 1e4, color="#54A24B")
        axes[0].set_title("各品牌总播放量（万）")
        axes[0].set_ylabel("万")
        axes[1].bar(d["品牌"], d["平均播放"] / 1e4, color="#B279A2")
        axes[1].set_title("各品牌平均播放量（万）")
        axes[1].set_ylabel("万")
        _save(fig, "04_品牌对比.png")

        fig, ax = plt.subplots(figsize=(8, 4.5))
        x = np.arange(len(d))
        w = 0.2
        ax.bar(x - 1.5 * w, d["平均点赞"], w, label="点赞")
        ax.bar(x - 0.5 * w, d["平均投币"], w, label="投币")
        ax.bar(x + 0.5 * w, d["平均收藏"], w, label="收藏")
        ax.bar(x + 1.5 * w, d["平均分享"], w, label="分享")
        ax.set_xticks(x)
        ax.set_xticklabels(d["品牌"])
        ax.set_title("各品牌平均互动量对比")
        ax.legend(fontsize=8)
        _save(fig, "05_品牌互动对比.png")

    # 图5 作者粉丝分层
    if "10_作者粉丝分层" in results:
        d = results["10_作者粉丝分层"]
        fig, ax = plt.subplots(figsize=(8, 4.5))
        ax.bar(d["粉丝层级"], d["作者数"], color="#72B7B2")
        ax.set_title("作者粉丝规模分层分布")
        ax.set_ylabel("作者数")
        ax.tick_params(axis="x", rotation=30)
        _save(fig, "06_作者粉丝分层.png")

    return saved


# ============================================================
# 五、主流程
# ============================================================

def main():
    src = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_FILE
    if not os.path.exists(src):
        sys.exit(f"[错误] 找不到文件：{src}\n请把 xlsx 放在脚本同目录，或用 python {os.path.basename(__file__)} <文件路径>")

    base_dir = os.path.dirname(os.path.abspath(src))
    out_dir = os.path.join(base_dir, "output")
    chart_dir = os.path.join(out_dir, "charts")
    os.makedirs(chart_dir, exist_ok=True)

    print(f"[1/4] 读取：{os.path.basename(src)}")
    sheets = pd.read_excel(src, sheet_name=None)

    log = {}
    print("[2/4] 清洗中……")
    market = clean_market(sheets["市场规模"], log) if "市场规模" in sheets else pd.DataFrame()
    authors = clean_author(sheets["作者规模"], log) if "作者规模" in sheets else pd.DataFrame()

    brand_frames = []
    for name in ["比亚迪", "自身比亚迪", "竞品特斯拉", "竞品蔚来"]:
        if name in sheets:
            brand_frames.append(clean_brand(sheets[name], name, log))
    brands = pd.concat(brand_frames, ignore_index=True) if brand_frames else pd.DataFrame()

    # 关键修正：「比亚迪」与「自身比亚迪」是同一账号的重叠抓取（约 97% 标题重复），
    # 直接拼接会把比亚迪的量算两遍。按「标题+品牌」跨表去重，保留播放量最高的一条。
    if not brands.empty:
        before_cross = len(brands)
        brands = (
            brands.sort_values("play", ascending=False, na_position="last")
            .drop_duplicates(subset=["品牌", "标题"], keep="first")
            .reset_index(drop=True)
        )
        overlap = before_cross - len(brands)
        log["品牌跨表去重"] = {
            "拼接后行数": before_cross,
            "跨表重复删除": overlap,
            "去重后行数": len(brands),
        }

    # 输出清洗数据
    clean_path = os.path.join(out_dir, "清洗数据.xlsx")
    with pd.ExcelWriter(clean_path, engine="openpyxl") as writer:
        if not market.empty:
            market.to_excel(writer, sheet_name="市场规模_清洗", index=False)
        if not authors.empty:
            authors.to_excel(writer, sheet_name="作者规模_清洗", index=False)
        if not brands.empty:
            brands.to_excel(writer, sheet_name="品牌视频_清洗", index=False)

        pd.DataFrame(
            [{"表名": k, **v} for k, v in log.items()]
        ).to_excel(writer, sheet_name="清洗日志", index=False)

    print(f"       已输出：{clean_path}")

    print("[3/4] 分析中……")
    results = analyze(market, authors, brands)

    analysis_path = os.path.join(out_dir, "分析结果.xlsx")
    with pd.ExcelWriter(analysis_path, engine="openpyxl") as writer:
        for name, table in results.items():
            table.to_excel(writer, sheet_name=name[:31], index=False)
    print(f"       已输出：{analysis_path}")

    print("[4/4] 绘图……")
    charts = plot_all(market, authors, brands, results, chart_dir)
    for p in charts:
        print(f"       {p}")

    # 控制台摘要
    print("\n" + "=" * 52)
    print("清洗与统计摘要")
    print("=" * 52)
    for name, info in log.items():
        print(f"[{name}] " + " | ".join(f"{k}={v}" for k, v in info.items()))
    print("\n关键结论：")
    if not market.empty:
        top_share = market["播放量_最终"].head(0)
        print(f"  · 样本视频 {len(market)} 条，覆盖 {market['作者'].nunique()} 位作者")
        print(f"  · 总播放 {market['播放量_最终'].sum()/1e8:.2f} 亿，"
              f"单视频播放中位数 {market['播放量_最终'].median()/1e4:.1f} 万")
        skew = market["播放量_最终"].mean() / market["播放量_最终"].median()
        print(f"  · 均值/中位数 = {skew:.2f}（>2 说明播放量高度头部集中）")
    if "07_品牌对比" in results:
        d = results["07_品牌对比"]
        for _, r in d.iterrows():
            print(f"  · {r['品牌']}：{int(r['视频数'])} 条视频，总播放 {r['总播放']/1e4:.0f} 万，"
                  f"平均 {r['平均播放']/1e4:.1f} 万，平均互动率 {r['平均互动率']:.3f}")
    print("\n完成，结果在 output/ 目录。")

    return results


if __name__ == "__main__":
    main()
