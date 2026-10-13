"""사용자 작업 이력(시각화 로그) 집계 프로그램.

입력 CSV 형식 (헤더 없음, 구분자 ','):
    Job종류,사용자명,"총데이터수",차트수,시작일시,종료일시,경과시간
    View,fharenheit,"887,576",18,2026-10-10 0:11:11,2026-10-10 0:14:16,3:05

데이터수는 천 단위 comma 때문에 "" 로 감싸여 있으며, 모든 값은 앞뒤 공백을 trim 한다.

사용법:
    python usage_stats.py <log.csv> [-o output] [--top 20] [--sla 60] [--no-charts]

출력:
    <output>/usage_stats.xlsx   집계 시트 모음
    <output>/charts/*.png       차트 이미지
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd

COLUMNS = ["job_type", "user", "rows", "charts", "start", "end", "elapsed_raw"]
DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"
WEEKDAYS_KO = ["월", "화", "수", "목", "금", "토", "일"]

# 컬럼 한글 표시명 (엑셀 출력용)
LABELS = {
    "job_type": "Job종류", "user": "사용자", "rows": "데이터수", "charts": "차트수",
    "start": "시작일시", "end": "종료일시", "elapsed_raw": "경과시간(원본)",
    "elapsed_sec": "경과초", "duration_sec": "실제경과초(종료-시작)",
    "rows_per_sec": "처리량(건/초)", "rows_per_chart": "차트당데이터수",
    "sec_per_chart": "차트당소요초", "date": "일자", "week": "주차", "month": "월",
    "hour": "시간대", "weekday": "요일", "is_weekend": "주말여부",
}

ROW_BINS = [0, 10_000, 100_000, 1_000_000, 10_000_000, np.inf]
ROW_BIN_LABELS = ["~1만", "1만~10만", "10만~100만", "100만~1000만", "1000만~"]
SEC_BINS = [-0.1, 10, 30, 60, 300, 600, np.inf]
SEC_BIN_LABELS = ["~10초", "10~30초", "30초~1분", "1~5분", "5~10분", "10분~"]
CHART_BINS = [-1, 0, 1, 5, 10, 20, np.inf]
CHART_BIN_LABELS = ["0개", "1개", "2~5개", "6~10개", "11~20개", "21개~"]


# --------------------------------------------------------------------------- #
# 로딩 & 전처리
# --------------------------------------------------------------------------- #
def parse_elapsed(value: str) -> float:
    """'M:ss' 또는 'H:MM:ss' 형식을 초로 변환한다. 파싱 실패 시 NaN."""
    try:
        parts = [int(p) for p in str(value).strip().split(":")]
    except ValueError:
        return np.nan
    if len(parts) == 2:
        m, s = parts
        return m * 60 + s
    if len(parts) == 3:
        h, m, s = parts
        return h * 3600 + m * 60 + s
    return np.nan


def load(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    """CSV를 읽어 (정상 데이터, 파싱 오류 행)을 반환한다."""
    raw = pd.read_csv(
        path, sep=",", quotechar='"', header=None, names=COLUMNS, dtype=str,
        skipinitialspace=True, on_bad_lines="warn", encoding="utf-8-sig",
    )
    # "" 안쪽 공백 포함, 모든 값 trim
    raw = raw.apply(lambda s: s.str.strip())
    raw.insert(0, "line_no", raw.index + 1)

    df = raw.copy()
    df["rows"] = pd.to_numeric(df["rows"].str.replace(",", "", regex=False), errors="coerce")
    df["charts"] = pd.to_numeric(df["charts"].str.replace(",", "", regex=False), errors="coerce")
    df["start"] = pd.to_datetime(df["start"], format=DATETIME_FORMAT, errors="coerce")
    df["end"] = pd.to_datetime(df["end"], format=DATETIME_FORMAT, errors="coerce")
    df["elapsed_sec"] = df["elapsed_raw"].map(parse_elapsed)

    required = ["job_type", "user", "rows", "charts", "start", "end", "elapsed_sec"]
    bad_mask = df[required].isna().any(axis=1)
    errors = raw[bad_mask].copy()
    errors["오류컬럼"] = [
        ", ".join(LABELS.get(c, c) for c in r.index[r])
        for _, r in df.loc[bad_mask, required].isna().iterrows()
    ]
    return df[~bad_mask].reset_index(drop=True), errors


def enrich(df: pd.DataFrame) -> pd.DataFrame:
    """집계에 쓰는 파생 컬럼을 추가한다."""
    df = df.copy()
    df["rows"] = df["rows"].astype("int64")
    df["charts"] = df["charts"].astype("int64")
    df["duration_sec"] = (df["end"] - df["start"]).dt.total_seconds()
    safe_sec = df["elapsed_sec"].where(df["elapsed_sec"] > 0)
    safe_charts = df["charts"].where(df["charts"] > 0)
    df["rows_per_sec"] = df["rows"] / safe_sec
    df["rows_per_chart"] = df["rows"] / safe_charts
    df["sec_per_chart"] = df["elapsed_sec"] / safe_charts

    df["date"] = df["start"].dt.normalize()
    df["week"] = df["start"].dt.to_period("W-SUN").dt.start_time
    df["month"] = df["start"].dt.to_period("M").astype(str)
    df["hour"] = df["start"].dt.hour
    df["weekday_no"] = df["start"].dt.weekday
    df["weekday"] = df["weekday_no"].map(lambda i: WEEKDAYS_KO[i])
    df["is_weekend"] = df["weekday_no"] >= 5

    df["rows_bin"] = pd.cut(df["rows"], ROW_BINS, labels=ROW_BIN_LABELS, right=False)
    df["sec_bin"] = pd.cut(df["elapsed_sec"], SEC_BINS, labels=SEC_BIN_LABELS)
    df["chart_bin"] = pd.cut(df["charts"], CHART_BINS, labels=CHART_BIN_LABELS)
    return df


# --------------------------------------------------------------------------- #
# 집계
# --------------------------------------------------------------------------- #
def p(q: float):
    def _p(s: pd.Series) -> float:
        return s.quantile(q)
    _p.__name__ = f"P{int(q * 100)}"
    return _p


def basic_agg(df: pd.DataFrame, by, extra_users: bool = True) -> pd.DataFrame:
    """공통 집계: 건수, 데이터/차트 합계·평균, 소요시간 통계, 처리량."""
    g = df.groupby(by, observed=False)
    out = pd.DataFrame({
        "Job수": g.size(),
        "데이터수 합계": g["rows"].sum(),
        "데이터수 평균": g["rows"].mean(),
        "차트수 합계": g["charts"].sum(),
        "차트수 평균": g["charts"].mean(),
        "소요초 합계": g["elapsed_sec"].sum(),
        "소요초 평균": g["elapsed_sec"].mean(),
        "소요초 중앙값": g["elapsed_sec"].median(),
        "소요초 P95": g["elapsed_sec"].quantile(0.95),
        "소요초 최대": g["elapsed_sec"].max(),
        "처리량 평균(건/초)": g["rows_per_sec"].mean(),
    })
    if extra_users:
        out.insert(1, "사용자수", g["user"].nunique())
    return out


def summary(df: pd.DataFrame, sla: int, errors: pd.DataFrame, max_conc: int) -> pd.DataFrame:
    e = df["elapsed_sec"]
    items = [
        ("집계 기간", f"{df['start'].min():%Y-%m-%d %H:%M:%S} ~ {df['end'].max():%Y-%m-%d %H:%M:%S}"),
        ("총 Job 수", len(df)),
        *[(f"  {t} Job 수 (비율)", f"{n:,} ({n / len(df):.1%})")
          for t, n in df["job_type"].value_counts().items()],
        ("고유 사용자 수", df["user"].nunique()),
        ("활동 일수", df["date"].nunique()),
        ("데이터수 합계", df["rows"].sum()),
        ("데이터수 평균", round(df["rows"].mean(), 1)),
        ("차트수 합계", df["charts"].sum()),
        ("차트수 평균", round(df["charts"].mean(), 2)),
        ("소요시간 합계(시간)", round(e.sum() / 3600, 2)),
        ("소요초 평균", round(e.mean(), 1)),
        ("소요초 중앙값(P50)", e.median()),
        ("소요초 P90", round(e.quantile(0.90), 1)),
        ("소요초 P95", round(e.quantile(0.95), 1)),
        ("소요초 P99", round(e.quantile(0.99), 1)),
        ("소요초 최소", e.min()),
        ("소요초 최대", e.max()),
        ("소요초 표준편차", round(e.std(), 1)),
        ("평균 처리량(건/초)", round(df["rows_per_sec"].mean(), 1)),
        (f"SLA({sla}초) 초과 건수", int((e > sla).sum())),
        (f"SLA({sla}초) 초과 비율", f"{(e > sla).mean():.1%}"),
        ("최대 동시 실행 Job 수", max_conc),
        ("데이터수-소요시간 상관계수", round(df["rows"].corr(e), 3)),
        ("차트수-소요시간 상관계수", round(df["charts"].corr(e), 3)),
        ("파싱 오류로 제외된 행", len(errors)),
    ]
    return pd.DataFrame(items, columns=["항목", "값"])


def user_stats(df: pd.DataFrame) -> pd.DataFrame:
    out = basic_agg(df, "user", extra_users=False)
    g = df.groupby("user")
    type_counts = df.pivot_table(index="user", columns="job_type", values="rows",
                                 aggfunc="size", fill_value=0)
    type_counts.columns = [f"{c} 건수" for c in type_counts.columns]
    out = out.join(type_counts)
    out["첫 사용일시"] = g["start"].min()
    out["마지막 사용일시"] = g["start"].max()
    out["활동 일수"] = g["date"].nunique()
    out = out.sort_values("Job수", ascending=False)

    # 사용량 등급: Job수 기준 상위 20% Heavy, 다음 30% Medium, 나머지 Light
    rank_pct = out["Job수"].rank(pct=True, ascending=False, method="first")
    out["등급"] = np.select([rank_pct <= 0.2, rank_pct <= 0.5], ["Heavy", "Medium"], "Light")
    out["누적 Job 비율"] = out["Job수"].cumsum() / out["Job수"].sum()
    return out


def pareto(users: pd.DataFrame) -> pd.DataFrame:
    n = len(users)
    rows = []
    for pct in (0.1, 0.2, 0.5):
        k = max(1, int(np.ceil(n * pct)))
        rows.append({
            "상위 사용자 비율": f"{pct:.0%}", "사용자 수": k,
            "Job 점유율": users["Job수"].iloc[:k].sum() / users["Job수"].sum(),
            "데이터 점유율": users["데이터수 합계"].nlargest(k).sum() / users["데이터수 합계"].sum(),
            "소요시간 점유율": users["소요초 합계"].nlargest(k).sum() / users["소요초 합계"].sum(),
        })
    return pd.DataFrame(rows)


def time_series(df: pd.DataFrame, key: str) -> pd.DataFrame:
    out = basic_agg(df, key)
    by_type = df.pivot_table(index=key, columns="job_type", values="rows",
                             aggfunc="size", fill_value=0)
    by_type.columns = [f"{c} 건수" for c in by_type.columns]
    out = out.join(by_type)
    out["Job수 증감률"] = out["Job수"].pct_change()
    out["누적 Job수"] = out["Job수"].cumsum()
    return out


def new_users_by_month(df: pd.DataFrame) -> pd.DataFrame:
    first = df.groupby("user")["start"].min().dt.to_period("M").astype(str)
    out = first.value_counts().sort_index().rename("신규 사용자수").to_frame()
    out["누적 사용자수"] = out["신규 사용자수"].cumsum()
    out.index.name = "월"
    return out


def concurrency(df: pd.DataFrame) -> tuple[pd.Series, int]:
    """시점별 동시 실행 Job 수 (시작 +1, 종료 -1 스윕)."""
    events = pd.concat([
        pd.Series(1, index=df["start"]),
        pd.Series(-1, index=df["end"]),
    ]).sort_index(kind="stable")
    # 같은 시각에 종료와 시작이 겹치면 종료를 먼저 처리
    events = events.groupby(level=0).sum()
    running = events.cumsum()
    return running, int(running.max()) if len(running) else 0


def outliers(df: pd.DataFrame) -> pd.DataFrame:
    """소요시간 IQR 이상치 + 데이터 규모 대비 느린 Job(로그 회귀 잔차)."""
    e = df["elapsed_sec"]
    q1, q3 = e.quantile([0.25, 0.75])
    upper = q3 + 1.5 * (q3 - q1)

    out = df.copy()
    out["IQR 이상치"] = e > upper
    valid = (out["rows"] > 0) & (e > 0)
    out["규모대비 느림"] = False
    if valid.sum() >= 3:
        x = np.log10(out.loc[valid, "rows"])
        y = np.log10(out.loc[valid, "elapsed_sec"])
        slope, intercept = np.polyfit(x, y, 1)
        resid = y - (slope * x + intercept)
        out.loc[valid, "예상소요초"] = 10 ** (slope * x + intercept)
        out.loc[valid, "규모대비 느림"] = resid > resid.mean() + 2 * resid.std()
    out = out[out["IQR 이상치"] | out["규모대비 느림"]]
    cols = ["line_no", "job_type", "user", "rows", "charts", "start", "end",
            "elapsed_raw", "elapsed_sec", "예상소요초", "rows_per_sec", "IQR 이상치", "규모대비 느림"]
    cols = [c for c in cols if c in out.columns]
    return out[cols].sort_values("elapsed_sec", ascending=False)


def quality(df: pd.DataFrame) -> pd.DataFrame:
    checks = {
        "경과시간≠(종료-시작)": (df["elapsed_sec"] - df["duration_sec"]).abs() > 1,
        "종료일시<시작일시": df["end"] < df["start"],
        "데이터수 0": df["rows"] == 0,
        "차트수 0": df["charts"] == 0,
        "중복(사용자+시작일시)": df.duplicated(["user", "start"], keep=False),
    }
    frames = []
    for name, mask in checks.items():
        if mask.any():
            f = df.loc[mask, ["line_no", "job_type", "user", "rows", "charts",
                              "start", "end", "elapsed_raw", "elapsed_sec", "duration_sec"]].copy()
            f.insert(0, "검증항목", name)
            frames.append(f)
    if not frames:
        return pd.DataFrame({"검증항목": ["이상 없음"]})
    return pd.concat(frames, ignore_index=True)


def build_tables(df: pd.DataFrame, errors: pd.DataFrame, top: int, sla: int) -> dict[str, pd.DataFrame]:
    running, max_conc = concurrency(df)
    users = user_stats(df)

    weekday_hour = df.pivot_table(index="weekday_no", columns="hour", values="rows",
                                  aggfunc="size", fill_value=0).reindex(index=range(7), columns=range(24), fill_value=0)
    weekday_hour.index = WEEKDAYS_KO

    top_users = users.head(top).index
    user_month = df[df["user"].isin(top_users)].pivot_table(
        index="user", columns="month", values="rows", aggfunc="size", fill_value=0
    ).reindex(top_users)

    user_type_sec = df.pivot_table(index="user", columns="job_type", values="elapsed_sec",
                                   aggfunc="mean").reindex(users.index)
    user_type_sec.columns = [f"{c} 평균소요초" for c in user_type_sec.columns]

    hourly = basic_agg(df, "hour").reindex(range(24))
    hourly["Job수"] = hourly["Job수"].fillna(0).astype(int)

    weekday = basic_agg(df, "weekday_no").reindex(range(7))
    weekday.index = WEEKDAYS_KO

    conc_daily = running.groupby(running.index.normalize()).max().rename("최대 동시 실행수").to_frame()

    corr = df[["rows", "charts", "elapsed_sec", "rows_per_sec"]].corr().round(3)
    corr.index = corr.columns = ["데이터수", "차트수", "소요초", "처리량"]

    return {
        "요약": summary(df, sla, errors, max_conc),
        "Job종류별": basic_agg(df, "job_type"),
        "사용자별": users,
        "사용자 파레토": pareto(users),
        "사용자xJob종류": user_type_sec,
        f"사용자x월(Top{top})": user_month,
        "일별": time_series(df, "date"),
        "주별": time_series(df, "week"),
        "월별": time_series(df, "month").join(new_users_by_month(df)),
        "시간대별": hourly,
        "요일별": weekday,
        "평일vs주말": basic_agg(df, "is_weekend").rename(index={False: "평일", True: "주말"}),
        "요일x시간대(건수)": weekday_hour,
        "Job종류x시간대": df.pivot_table(index="hour", columns="job_type", values="rows",
                                      aggfunc="size", fill_value=0),
        "데이터규모 구간": basic_agg(df, "rows_bin"),
        "Job종류x데이터규모": df.pivot_table(index="rows_bin", columns="job_type", values="elapsed_sec",
                                        aggfunc=["size", "mean", p(0.95)], observed=False),
        "소요시간 구간": basic_agg(df, "sec_bin"),
        "차트수 구간": basic_agg(df, "chart_bin"),
        "상관계수": corr,
        "일별 최대동시실행": conc_daily,
        "이상치": outliers(df),
        "품질검증": quality(df),
        "파싱오류": errors if len(errors) else pd.DataFrame({"결과": ["오류 없음"]}),
        "원본(파생포함)": df.drop(columns=["weekday_no"]).rename(columns=LABELS),
    }


# --------------------------------------------------------------------------- #
# 엑셀 출력
# --------------------------------------------------------------------------- #
def write_excel(tables: dict[str, pd.DataFrame], path: Path) -> None:
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="E8EEF7")
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, table in tables.items():
            sheet = name[:31].replace("/", "-")
            t = table.copy()
            if isinstance(t.columns, pd.MultiIndex):
                t.columns = [" ".join(map(str, c)).strip() for c in t.columns]
            t.index.name = LABELS.get(t.index.name, t.index.name)
            keep_index = not isinstance(t.index, pd.RangeIndex)
            t.to_excel(writer, sheet_name=sheet, index=keep_index)

            ws = writer.sheets[sheet]
            ws.freeze_panes = "B2" if keep_index else "A2"
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.fill = header_fill
            for i, col in enumerate(ws.columns, start=1):
                values = [c.value for c in col[:200]]
                width = max(len(str(v)) if v is not None else 0 for v in values)
                ws.column_dimensions[get_column_letter(i)].width = min(max(10, width * 1.3 + 2), 40)
                for c in col[1:]:
                    if isinstance(c.value, float):
                        header = str(col[0].value)
                        c.number_format = "0.0%" if ("비율" in header or "점유율" in header
                                                     or "증감률" in header) else "#,##0.0"
                    elif isinstance(c.value, int) and not isinstance(c.value, bool):
                        c.number_format = "#,##0"
                    elif hasattr(c.value, "year"):
                        c.number_format = "yyyy-mm-dd hh:mm:ss"


# --------------------------------------------------------------------------- #
# 차트
# --------------------------------------------------------------------------- #
# dataviz 기본 팔레트 (라이트 모드)
SURFACE = "#fcfcfb"
TEXT = "#0b0b0b"
TEXT_2 = "#52514e"
GRID = "#e4e3df"
SERIES = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
SEQ_BLUE = ["#fcfcfb", "#cde2fb", "#9ec5f4", "#6da7ec", "#3987e5", "#256abf", "#184f95", "#0d366b"]


def setup_matplotlib():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager

    for name in ["Noto Sans CJK KR", "NanumGothic", "Malgun Gothic", "AppleGothic"]:
        if any(f.name == name for f in font_manager.fontManager.ttflist):
            plt.rcParams["font.family"] = name
            break
    plt.rcParams.update({
        "axes.unicode_minus": False,
        "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
        "text.color": TEXT, "axes.labelcolor": TEXT_2, "xtick.color": TEXT_2, "ytick.color": TEXT_2,
        "axes.edgecolor": GRID, "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.8,
        "axes.spines.top": False, "axes.spines.right": False, "axes.axisbelow": True,
        "axes.titlesize": 13, "axes.titleweight": "bold", "axes.titlelocation": "left",
        "font.size": 10, "lines.linewidth": 2, "legend.frameon": False,
    })
    return plt


def type_colors(df: pd.DataFrame) -> dict[str, str]:
    # 색은 순위가 아니라 엔티티에 고정: View=1번, Batch=2번 슬롯, 그 외는 다음 슬롯
    order = ["View", "Batch"] + sorted(t for t in df["job_type"].unique() if t not in ("View", "Batch"))
    return {t: SERIES[i % len(SERIES)] for i, t in enumerate(order)}


def draw_charts(df: pd.DataFrame, tables: dict[str, pd.DataFrame], out_dir: Path, top: int) -> list[Path]:
    plt = setup_matplotlib()
    from matplotlib.colors import LinearSegmentedColormap
    from matplotlib.ticker import FuncFormatter

    out_dir.mkdir(parents=True, exist_ok=True)
    colors = type_colors(df)
    types = [t for t in colors if t in set(df["job_type"])]
    thousands = FuncFormatter(lambda v, _: f"{v:,.0f}")
    saved: list[Path] = []

    def save(fig, name):
        path = out_dir / name
        fig.tight_layout()
        fig.savefig(path, dpi=144)
        plt.close(fig)
        saved.append(path)

    # 1. 일별 Job 수 추이 (Job 종류별 누적 막대)
    daily = df.pivot_table(index="date", columns="job_type", values="rows", aggfunc="size", fill_value=0)
    fig, ax = plt.subplots(figsize=(12, 4.5))
    bottom = np.zeros(len(daily))
    for t in types:
        if t in daily:
            ax.bar(daily.index, daily[t], bottom=bottom, color=colors[t], label=t,
                   edgecolor=SURFACE, linewidth=1, width=0.8)
            bottom += daily[t].to_numpy()
    ax.set_title("일별 Job 수")
    ax.set_ylabel("Job 수")
    ax.legend(loc="upper left", ncols=len(types))
    fig.autofmt_xdate()
    save(fig, "01_daily_jobs.png")

    # 2. 일별 소요시간 평균 / P95 추이
    d = df.groupby("date")["elapsed_sec"]
    fig, ax = plt.subplots(figsize=(12, 4.5))
    ax.plot(d.mean().index, d.mean(), color=SERIES[0], label="평균", marker="o", markersize=4)
    ax.plot(d.quantile(0.95).index, d.quantile(0.95), color=SERIES[1], label="P95", marker="o", markersize=4)
    ax.set_title("일별 소요시간 추이")
    ax.set_ylabel("소요시간(초)")
    ax.legend(loc="upper left", ncols=2)
    fig.autofmt_xdate()
    save(fig, "02_daily_elapsed.png")

    # 3. 월별 처리 데이터 / 차트 수 (단위가 달라 2개 패널로 분리)
    monthly = df.groupby("month").agg(rows=("rows", "sum"), charts=("charts", "sum"),
                                      users=("user", "nunique"))
    fig, axes = plt.subplots(1, 3, figsize=(14, 4.2))
    for ax, col, title in zip(axes, ["rows", "charts", "users"],
                              ["월별 처리 데이터수", "월별 생성 차트수", "월별 활성 사용자수"]):
        ax.bar(monthly.index, monthly[col], color=SERIES[0], width=0.6)
        ax.set_title(title)
        ax.yaxis.set_major_formatter(thousands)
        ax.tick_params(axis="x", rotation=45)
        ax.grid(axis="x", visible=False)
    save(fig, "03_monthly_volume.png")

    # 4. 시간대별 Job 수 (Job 종류별 누적)
    hourly = df.pivot_table(index="hour", columns="job_type", values="rows", aggfunc="size",
                            fill_value=0).reindex(range(24), fill_value=0)
    fig, ax = plt.subplots(figsize=(12, 4.5))
    bottom = np.zeros(24)
    for t in types:
        if t in hourly:
            ax.bar(hourly.index, hourly[t], bottom=bottom, color=colors[t], label=t,
                   edgecolor=SURFACE, linewidth=1, width=0.8)
            bottom += hourly[t].to_numpy()
    ax.set_xticks(range(24))
    ax.set_title("시간대별 Job 수")
    ax.set_xlabel("시작 시각(시)")
    ax.set_ylabel("Job 수")
    ax.grid(axis="x", visible=False)
    ax.legend(loc="upper left", ncols=len(types))
    save(fig, "04_hourly_jobs.png")

    # 5. 요일 x 시간대 히트맵
    heat = tables["요일x시간대(건수)"]
    cmap = LinearSegmentedColormap.from_list("seq_blue", SEQ_BLUE)
    fig, ax = plt.subplots(figsize=(13, 4.2))
    im = ax.imshow(heat.to_numpy(), cmap=cmap, aspect="auto")
    ax.set_xticks(range(24), labels=range(24))
    ax.set_yticks(range(7), labels=heat.index)
    ax.grid(False)
    vmax = heat.to_numpy().max()
    for (i, j), v in np.ndenumerate(heat.to_numpy()):
        if v:
            ax.text(j, i, v, ha="center", va="center", fontsize=8,
                    color="#ffffff" if v > vmax * 0.55 else TEXT)
    ax.set_title("요일 × 시간대 Job 수")
    ax.set_xlabel("시작 시각(시)")
    fig.colorbar(im, ax=ax, fraction=0.025, pad=0.01)
    save(fig, "05_weekday_hour_heatmap.png")

    # 6. Top N 사용자 (Job 수, Job 종류별 누적 가로 막대)
    users = tables["사용자별"].head(top)
    fig, ax = plt.subplots(figsize=(10, max(4, 0.35 * len(users) + 1.5)))
    left = np.zeros(len(users))
    for t in types:
        col = f"{t} 건수"
        if col in users:
            ax.barh(users.index, users[col], left=left, color=colors[t], label=t,
                    edgecolor=SURFACE, linewidth=1, height=0.7)
            left += users[col].to_numpy()
    for y, total in enumerate(users["Job수"]):
        ax.text(total, y, f" {total:,}", va="center", color=TEXT_2, fontsize=9)
    ax.invert_yaxis()
    ax.set_title(f"사용자별 Job 수 Top {len(users)}")
    ax.set_xlabel("Job 수")
    ax.grid(axis="y", visible=False)
    ax.legend(loc="lower right")
    save(fig, "06_top_users.png")

    # 7. Top N 사용자 처리 데이터량
    by_rows = tables["사용자별"]["데이터수 합계"].nlargest(top)
    fig, ax = plt.subplots(figsize=(10, max(4, 0.35 * len(by_rows) + 1.5)))
    ax.barh(by_rows.index, by_rows, color=SERIES[0], height=0.7)
    for y, v in enumerate(by_rows):
        ax.text(v, y, f" {v:,.0f}", va="center", color=TEXT_2, fontsize=9)
    ax.invert_yaxis()
    ax.xaxis.set_major_formatter(thousands)
    ax.set_title(f"사용자별 처리 데이터수 Top {len(by_rows)}")
    ax.grid(axis="y", visible=False)
    save(fig, "07_top_users_rows.png")

    # 8. 데이터수 vs 소요시간 산점도 (로그 스케일)
    fig, ax = plt.subplots(figsize=(10, 6))
    for t in types:
        s = df[(df["job_type"] == t) & (df["rows"] > 0) & (df["elapsed_sec"] > 0)]
        ax.scatter(s["rows"], s["elapsed_sec"], s=22, color=colors[t], alpha=0.65,
                   edgecolors=SURFACE, linewidths=0.8, label=t)
    ax.set_xscale("log")
    ax.set_yscale("log")
    ax.xaxis.set_major_formatter(thousands)
    ax.set_xlabel("데이터수 (log)")
    ax.set_ylabel("소요시간 초 (log)")
    r = df["rows"].corr(df["elapsed_sec"])
    ax.set_title(f"데이터수 vs 소요시간  (상관계수 r = {r:.2f})")
    ax.legend(loc="upper left")
    save(fig, "08_rows_vs_elapsed.png")

    # 9. 차트수 vs 소요시간
    fig, ax = plt.subplots(figsize=(10, 6))
    for t in types:
        s = df[df["job_type"] == t]
        ax.scatter(s["charts"], s["elapsed_sec"], s=22, color=colors[t], alpha=0.65,
                   edgecolors=SURFACE, linewidths=0.8, label=t)
    ax.set_xlabel("차트수")
    ax.set_ylabel("소요시간(초)")
    r = df["charts"].corr(df["elapsed_sec"])
    ax.set_title(f"차트수 vs 소요시간  (상관계수 r = {r:.2f})")
    ax.legend(loc="upper left")
    save(fig, "09_charts_vs_elapsed.png")

    # 10. Job 종류별 소요시간 분포 (박스플롯)
    fig, ax = plt.subplots(figsize=(8, 5))
    data = [df.loc[df["job_type"] == t, "elapsed_sec"] for t in types]
    bp = ax.boxplot(data, tick_labels=types, patch_artist=True, widths=0.5,
                    medianprops={"color": TEXT, "linewidth": 2},
                    flierprops={"marker": "o", "markersize": 4, "markerfacecolor": TEXT_2,
                                "markeredgecolor": SURFACE, "alpha": 0.6})
    for patch, t in zip(bp["boxes"], types):
        patch.set_facecolor(colors[t])
        patch.set_alpha(0.85)
        patch.set_edgecolor(SURFACE)
    ax.set_title("Job 종류별 소요시간 분포")
    ax.set_ylabel("소요시간(초)")
    ax.grid(axis="x", visible=False)
    save(fig, "10_elapsed_box_by_type.png")

    # 11. 소요시간 구간 분포
    sec_bins = df.groupby(["sec_bin", "job_type"], observed=False).size().unstack(fill_value=0)
    fig, ax = plt.subplots(figsize=(10, 4.5))
    x = np.arange(len(sec_bins))
    width = 0.8 / max(1, len(types))
    for i, t in enumerate(types):
        if t in sec_bins:
            ax.bar(x + (i - (len(types) - 1) / 2) * width, sec_bins[t], width=width - 0.03,
                   color=colors[t], label=t)
    ax.set_xticks(x, labels=sec_bins.index.astype(str))
    ax.set_title("소요시간 구간별 Job 수")
    ax.set_ylabel("Job 수")
    ax.grid(axis="x", visible=False)
    ax.legend(loc="upper right")
    save(fig, "11_elapsed_bins.png")

    # 12. 데이터 규모 구간별 평균 / P95 소요시간
    rb = tables["데이터규모 구간"]
    fig, ax = plt.subplots(figsize=(10, 4.5))
    x = np.arange(len(rb))
    ax.bar(x - 0.2, rb["소요초 평균"], width=0.37, color=SERIES[0], label="평균")
    ax.bar(x + 0.2, rb["소요초 P95"], width=0.37, color=SERIES[1], label="P95")
    ax.set_xticks(x, labels=[f"{i}\n({n:,}건)" for i, n in zip(rb.index.astype(str), rb["Job수"])])
    ax.set_title("데이터 규모 구간별 소요시간")
    ax.set_ylabel("소요시간(초)")
    ax.grid(axis="x", visible=False)
    ax.legend(loc="upper left")
    save(fig, "12_rowbin_elapsed.png")

    return saved


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="사용자 작업 이력(시각화 로그) 집계")
    ap.add_argument("csv", type=Path, help="입력 로그 파일 (헤더 없음, ',' 구분)")
    ap.add_argument("-o", "--output", type=Path, default=Path("output"), help="출력 디렉토리 (기본: output)")
    ap.add_argument("--top", type=int, default=20, help="Top N 사용자 수 (기본: 20)")
    ap.add_argument("--sla", type=int, default=60, help="SLA 기준 소요초 (기본: 60)")
    ap.add_argument("--no-charts", action="store_true", help="차트 생성 생략")
    args = ap.parse_args(argv)

    if not args.csv.exists():
        print(f"파일이 없습니다: {args.csv}", file=sys.stderr)
        return 1

    df, errors = load(args.csv)
    if df.empty:
        print("유효한 데이터 행이 없습니다.", file=sys.stderr)
        return 1
    df = enrich(df)

    tables = build_tables(df, errors, args.top, args.sla)
    args.output.mkdir(parents=True, exist_ok=True)
    xlsx = args.output / "usage_stats.xlsx"
    write_excel(tables, xlsx)

    print(tables["요약"].to_string(index=False))
    print(f"\n엑셀: {xlsx} ({len(tables)}개 시트)")
    if not args.no_charts:
        charts = draw_charts(df, tables, args.output / "charts", args.top)
        print(f"차트: {args.output / 'charts'} ({len(charts)}개)")
    if len(errors):
        print(f"경고: 파싱 오류 {len(errors)}행 제외 ('파싱오류' 시트 참고)", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
