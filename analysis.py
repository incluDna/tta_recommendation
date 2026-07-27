# analysis.py — pipeline functions (ไม่มี Streamlit dependency)
# ─────────────────────────────────────────────────────────────────
import os
import warnings
import re
from datetime import datetime
from pathlib import Path
from typing import Optional
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import OrdinalEncoder
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score

import numpy as np
import pandas as pd

warnings.filterwarnings('ignore')

from config import (
    AREA_INSURANCE_MAP,
    CAUSE_THEME,
    COL_4M1E,
    COL_AGE,
    COL_AREA,
    COL_CAUSE,
    COL_DRIVER_EXP,
    COL_EMP_ID,
    COL_LEAVE,
    COL_MONTH,
    COL_PERIOD,
    COL_QUARTER,
    COL_RIDER_EXP,
    COL_ROAD_SURFACE,
    COL_ROAD_TYPE,
    COL_SEVERITY,
    COL_SPEED,
    COL_SPEED_GROUP,
    COL_TIME,
    COL_TRAFFIC,
    COL_VEHICLE,
    COL_VISIBILITY,
    COL_YEAR,
    COLUMN_RENAME,
    DROP_COLS,
    FUZZY_THRESHOLD,
    MONTH_ABBR_ORDER,
    MONTH_ORDER,
    MONTH_TO_QUARTER,
    NORMALIZE_DICT,
    OUT_AREA_FILE,
    OUT_JOINED_FILE,
    OUT_RIDER_FILE,
    OUT_UNMAPPED_FILE,
    OUTPUT_DIR,
    PRIORITY_BINS,
    PRIORITY_LABELS,
    RIDER_INSURANCE_MAP,
    RISK_WEIGHTS,
    THEMES,
    RENAME_DICT_69, 
    CAUSE_THEME, 
    MONTH_MAP, 
    MODEL_FEATURES, 
    DEFAULT_CAUSE_COL, 
    DEFAULT_THEME
)

try:
    from rapidfuzz import process as fuzz_process
    _RAPIDFUZZ_AVAILABLE = True
except ImportError:
    _RAPIDFUZZ_AVAILABLE = False


# ═══════════════════════════════════════════════════════════════════
# 1. FILE UTILITIES
# ═══════════════════════════════════════════════════════════════════

def list_excel_files(folder: str) -> list[str]:
    """คืน list ชื่อไฟล์ .xlsx / .xls ใน folder (เรียงตามชื่อ)"""
    p = Path(folder)
    if not p.exists():
        return []
    return sorted(f.name for f in p.glob("*.xls*") if not f.name.startswith("~$"))


def get_sheet_names(folder: str, filename: str) -> list[str]:
    """คืน list ชื่อ sheet ในไฟล์ Excel"""
    path = Path(folder) / filename
    with pd.ExcelFile(path, engine="openpyxl") as xf:
        return xf.sheet_names


def load_raw(folder: str, filename: str, sheet) -> pd.DataFrame:
    """โหลด sheet จาก Excel → DataFrame ดิบ"""
    path = Path(folder) / filename
    df_raw = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    # หากพบว่าคอลัมน์แรกๆ มีคำว่า Unnamed เยอะ ให้ลองเลื่อน header ลงมาหาแถวที่มีคำว่า "พื้นที่" หรือ "สาเหตุ"
    if any("Unnamed" in str(c) for c in df_raw.columns[:5]):
        for idx, row in df_raw.iterrows():
            row_str = row.astype(str).values
            if any("พื้นที่" in s or "สาเหตุ" in s for s in row_str):
                df_raw = pd.read_excel(path, sheet_name=sheet, skiprows=idx+1, engine="openpyxl")
                break
                
    return df_raw


# ═══════════════════════════════════════════════════════════════════
# 2. COLUMN RESOLVER & SORTING UTILITIES
# ═══════════════════════════════════════════════════════════════════

_REQUIRED_COLS = [COL_AREA, COL_CAUSE, COL_SEVERITY]
_OPTIONAL_COLS = [
    COL_AGE, COL_SPEED, COL_LEAVE, COL_TIME, COL_MONTH, COL_RIDER_EXP, COL_DRIVER_EXP
]
_FUZZY_COL_THRESHOLD = 80


def _col_key(text: str) -> str:
    """ลบ whitespace ทั้งหมด → เปรียบเทียบแบบ whitespace-agnostic"""
    return re.sub(r"\s+", "", str(text))


def natural_sort_key(text) -> tuple:
    """Key สำหรับเรียงชื่อพื้นที่แบบธรรมชาติ (e.g. เขต 2 ก่อน เขต 10)"""
    text = str(text)
    parts = re.split(r"(\d+)", text)
    return tuple(int(p) if p.isdigit() else p for p in parts)


def _normalize_month_name(text) -> str:
    if pd.isna(text):
        return text
    return str(text).strip().title()


def month_sort_key(month_name) -> int:
    """Key สำหรับเรียงเดือนตามปฏิทิน (Jan → Dec)"""
    name = _normalize_month_name(month_name)
    if name in MONTH_ORDER:
        return MONTH_ORDER.index(name)
    if name in MONTH_ABBR_ORDER:
        return MONTH_ABBR_ORDER.index(name)
    name3 = name[:3]
    if name3 in MONTH_ABBR_ORDER:
        return MONTH_ABBR_ORDER.index(name3)
    return 999


def get_month_order(month_series: pd.Series) -> list[str]:
    """คืน list ชื่อเดือนที่มีอยู่จริงใน series เรียงตามปฏิทินแล้ว"""
    unique_months = month_series.dropna().unique().tolist()
    return sorted(unique_months, key=month_sort_key)


def resolve_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    1. Strip whitespace/\n จากชื่อ column
    2. Fuzzy match column
    3. Raise ValueError หาก missing required columns
    """
    df = df.copy()
    df.columns = pd.Index([re.sub(r"\s+", " ", str(c)).strip() for c in df.columns])

    all_targets = _REQUIRED_COLS + _OPTIONAL_COLS + DROP_COLS
    rename_map: dict[str, str] = {}
    col_report: dict[str, str] = {}
    missing_required: list[str] = []
    actual_key_map = {_col_key(c): c for c in df.columns}

    for target in all_targets:
        target_key = _col_key(target)

        if target_key in actual_key_map:
            actual = actual_key_map[target_key]
            if actual != target:
                rename_map[actual] = target
                col_report[target] = f"normalized '{actual}'"
            else:
                col_report[target] = "exact"
            continue

        if _RAPIDFUZZ_AVAILABLE:
            result = fuzz_process.extractOne(target, df.columns.tolist())
            if result and result[1] >= _FUZZY_COL_THRESHOLD:
                actual, score, _ = result
                rename_map[actual] = target
                col_report[target] = f"fuzzy {score:.0f}% '{actual}'"
                continue

        col_report[target] = "NOT FOUND"
        if target in _REQUIRED_COLS:
            missing_required.append(target)

    if missing_required:
        best = lambda c: (
            fuzz_process.extractOne(c, df.columns.tolist())
            if _RAPIDFUZZ_AVAILABLE else None
        )
        lines = "\n".join(
            f"  ✗ {c}  →  best match: {best(c)}"
            for c in missing_required
        )
        raise ValueError(
            f"ไม่พบ column ที่จำเป็น {len(missing_required)} รายการ:\n{lines}\n\n"
            f"column ในไฟล์:\n" + "\n".join(f"  • {c}" for c in df.columns)
        )

    df = df.rename(columns=rename_map)
    return df, col_report


# ═══════════════════════════════════════════════════════════════════
# 3. CLEANING & FEATURE ENGINEERING
# ═══════════════════════════════════════════════════════════════════

def clean(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """ทำความสะอาดข้อมูลเบื้องต้นและแปลงชนิดข้อมูลให้ถูกต้อง"""
    df = df.copy()
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed", na=False)].dropna(axis=1, how="all")
    df = df.rename(columns=COLUMN_RENAME)
    df, col_report = resolve_columns(df)

    drop_normalized = {_col_key(c) for c in DROP_COLS}
    actual_drop = [c for c in df.columns if _col_key(c) in drop_normalized]
    df = df.drop(columns=actual_drop, errors="ignore")

    str_cols = df.select_dtypes("object").columns
    for col in str_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.strip()
            .replace({"": np.nan, "nan": np.nan, "<NA>": np.nan, "None": np.nan})
        )

    for col in [COL_AGE, COL_SPEED, COL_LEAVE]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if COL_MONTH in df.columns:
        df[COL_MONTH] = df[COL_MONTH].apply(_normalize_month_name)

    return df, col_report


def _exp_to_months(text) -> float:
    if pd.isna(text):
        return np.nan
    text = str(text)
    y = int(m.group(1)) if (m := re.search(r"(\d+)ปี", text)) else 0
    mo = int(m.group(1)) if (m := re.search(r"(\d+)เดือน", text)) else 0
    d = int(m.group(1)) if (m := re.search(r"(\d+)วัน", text)) else 0
    w = int(m.group(1)) if (m := re.search(r"(\d+)สัปดาห์", text)) else 0
    return y * 12 + mo + d / 30 + w / 4


def _extract_hour(text) -> float:
    if pd.isna(text):
        return np.nan
    match = re.search(r"(\d{1,2})[:\.](\d{1,2})", str(text).strip())
    return int(match.group(1)) if match else np.nan


def _get_period(hour) -> str | float:
    if pd.isna(hour):
        return np.nan
    if hour < 6:
        return "Late Night"
    if hour < 12:
        return "Morning"
    if hour < 18:
        return "Afternoon"
    return "Night"


def feature_engineer(df: pd.DataFrame) -> pd.DataFrame:
    """เพิ่ม Features: ประสบการณ์, กลุ่มอายุ, ช่วงเวลา, Year, Quarter"""
    df = df.copy()

    if COL_RIDER_EXP in df.columns:
        df["rider_exp_month"] = df[COL_RIDER_EXP].apply(_exp_to_months)
        df["rider_exp_group"] = (
            pd.cut(
                df["rider_exp_month"],
                bins=[-1, 3, 12, 24, 999],
                labels=["<3m", "3-12m", "1-2y", ">2y"],
            ).astype(str)
        )

    if COL_DRIVER_EXP in df.columns:
        df["driver_exp_month"] = df[COL_DRIVER_EXP].apply(_exp_to_months)

    if COL_AGE in df.columns:
        df["age_group"] = pd.cut(
            df[COL_AGE],
            bins=[0, 20, 25, 30, 100],
            labels=["<20", "20-25", "26-30", ">30"],
        )

    if COL_TIME in df.columns:
        df["hour"] = df[COL_TIME].apply(_extract_hour)
        df["period"] = df["hour"].apply(_get_period)

    date_col = "วันที่เกิดอุบัติเหตุ(พ.ศ. เท่านั้น)"
    if date_col in df.columns:
        df["year"] = (
            df[date_col]
            .astype(str)
            .str[:4]
            .pipe(pd.to_numeric, errors="coerce")
            .astype("Int64")
        )

    if COL_MONTH in df.columns:
        df["quarter"] = df[COL_MONTH].map(MONTH_TO_QUARTER)

    return df


# ═══════════════════════════════════════════════════════════════════
# 4. THEME MAPPING
# ═══════════════════════════════════════════════════════════════════

def _normalize_cause(text: str) -> str:
    if pd.isna(text):
        return text
    text = str(text).strip()
    return NORMALIZE_DICT.get(text, text)


def _split_cause_tokens(text: str) -> list[str]:
    text_clean = re.sub(r"\(.*?\)", "", str(text))
    tokens = [t.strip() for t in text_clean.split("/") if t.strip()]
    return tokens if tokens else [text_clean.strip()]


def _token_theme(text: str) -> str | None:
    tokens = _split_cause_tokens(text)
    for token in tokens:
        for key, theme in CAUSE_THEME.items():
            if token in key or key in token:
                return theme
    return None


def _fuzzy_theme(text: str) -> str | None:
    if not _RAPIDFUZZ_AVAILABLE or pd.isna(text):
        return None
    match = fuzz_process.extractOne(text, CAUSE_THEME.keys())
    if match is None:
        return None
    keyword, score, _ = match
    return CAUSE_THEME[keyword] if score >= FUZZY_THRESHOLD else None


def _fuzzy_theme_token(text: str) -> str | None:
    if not _RAPIDFUZZ_AVAILABLE:
        return None
    tokens = _split_cause_tokens(text)
    best_theme, best_score = None, 0
    for token in tokens:
        match = fuzz_process.extractOne(token, CAUSE_THEME.keys())
        if match is None:
            continue
        keyword, score, _ = match
        if score > best_score:
            best_score, best_theme = score, CAUSE_THEME[keyword]
    return best_theme if best_score >= FUZZY_THRESHOLD else None


def map_theme(cause) -> str | None:
    if pd.isna(cause):
        return None
    cause = _normalize_cause(cause)

    if cause in CAUSE_THEME:
        return CAUSE_THEME[cause]

    theme = _token_theme(cause)
    if theme is not None:
        return theme

    theme = _fuzzy_theme(cause)
    if theme is not None:
        return theme

    return _fuzzy_theme_token(cause)


def apply_theme_mapping(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = df.copy()
    if COL_CAUSE not in df.columns:
        df["campaign_theme"] = np.nan
        return df, pd.DataFrame()

    df["campaign_theme"] = df[COL_CAUSE].apply(map_theme)

    if COL_MONTH in df.columns:
        df["quarter"] = df[COL_MONTH].map(MONTH_TO_QUARTER)

    unmapped = (
        df.loc[df["campaign_theme"].isna(), COL_CAUSE]
        .dropna()
        .drop_duplicates()
        .sort_values()
        .reset_index(drop=True)
    )
    return df, unmapped.to_frame("cause")


# ═══════════════════════════════════════════════════════════════════
# 5. RISK SCORE + AREA SUMMARY
# ═══════════════════════════════════════════════════════════════════

def _find_sev_col(severity_table: pd.DataFrame) -> str:
    target = "หยุดงานเกิน3วัน"
    if target in severity_table.columns:
        return target
    candidates = [c for c in severity_table.columns if "เกิน" in c]
    return candidates[0] if candidates else severity_table.columns[-1]


def build_area_summary(df: pd.DataFrame) -> pd.DataFrame:
    if COL_AREA not in df.columns or "campaign_theme" not in df.columns:
        return pd.DataFrame()

    ct = pd.crosstab(df[COL_AREA], df["campaign_theme"])
    for t in THEMES:
        if t not in ct.columns:
            ct[t] = 0
    theme_pct = (ct.div(ct.sum(axis=1), axis=0) * 100).round(1)

    sev_table = (
        pd.crosstab(df[COL_AREA], df[COL_SEVERITY], normalize="index") * 100
    ).round(1)
    sev_col = _find_sev_col(sev_table)
    severity_pct = (
        sev_table[sev_col] if sev_col in sev_table.columns else pd.Series(0, index=ct.index)
    )

    theme_impact_table = (
        pd.crosstab(df["campaign_theme"], df[COL_SEVERITY], normalize="index") * 100
    )
    if sev_col not in theme_impact_table.columns:
        theme_impact_table[sev_col] = 0
    impact_map = theme_impact_table[sev_col].to_dict()

    PCT_MAP = {
        "Defensive Driving": "defensive_pct",
        "Focus & Attention": "focus_pct",
        "Road & Vehicle Safety": "road_pct",
        "Speed Awareness": "speed_pct",
    }

    area = pd.DataFrame({
        "severity_pct": severity_pct,
        "defensive_pct": theme_pct.get("Defensive Driving", 0),
        "focus_pct": theme_pct.get("Focus & Attention", 0),
        "road_pct": theme_pct.get("Road & Vehicle Safety", 0),
        "speed_pct": theme_pct.get("Speed Awareness", 0),
    })

    score_cols = []
    for pct_col, theme in {v: k for k, v in PCT_MAP.items()}.items():
        sc = pct_col.replace("_pct", "_score")
        area[sc] = (area[pct_col] * impact_map.get(theme, 0)).round(2)
        score_cols.append(sc)

    area["risk_score"] = (
        area["severity_pct"] * RISK_WEIGHTS["severity_pct"]
        + area["road_pct"] * RISK_WEIGHTS["road_pct"]
        + area["speed_pct"] * RISK_WEIGHTS["speed_pct"]
    ).round(2)

    area["priority_level"] = pd.cut(
        area["risk_score"],
        bins=PRIORITY_BINS,
        labels=PRIORITY_LABELS,
    )

    pct_to_theme = {v: k for k, v in PCT_MAP.items()}
    score_to_theme = {v.replace("_pct", "_score"): k for k, v in PCT_MAP.items()}

    area["dominant_theme"] = area[list(PCT_MAP.values())].idxmax(axis=1).map(pct_to_theme)

    def _campaign_priority(row):
        ranked = row[score_cols].sort_values(ascending=False)
        return pd.Series({
            "recommended_campaign": score_to_theme[ranked.index[0]],
            "supporting_campaigns": " > ".join(
                score_to_theme[ranked.index[i]] for i in [1, 2]
            ),
        })

    area[["recommended_campaign", "supporting_campaigns"]] = area.apply(
        _campaign_priority, axis=1
    )

    sorted_scores = np.sort(area[score_cols].values, axis=1)
    top1, top2 = sorted_scores[:, -1], sorted_scores[:, -2]
    denom = top1 + top2
    area["confidence"] = np.where(
        denom > 0, (top1 / denom * 100).round(1), np.nan
    )

    def generate_reason(row):
        if row["dominant_theme"] == row["recommended_campaign"]:
            return (
                f"พื้นที่นี้มีความเสี่ยงอยู่ในระดับ {row['priority_level']} "
                f"(Risk Score {row['risk_score']:.1f}) "
                f"สาเหตุที่พบมากที่สุดคือ '{row['dominant_theme']}' "
                f"และเมื่อพิจารณาความถี่ร่วมกับระดับความรุนแรงแล้ว "
                f"'{row['recommended_campaign']}' ก็ยังคงเป็นประเด็นสำคัญที่สุดเช่นกัน "
                f"จึงควรได้รับการรณรงค์เป็นอันดับแรก"
            )
        else:
            return (
                f"พื้นที่นี้มีความเสี่ยงอยู่ในระดับ {row['priority_level']} "
                f"(Risk Score {row['risk_score']:.1f}) "
                f"แม้ว่าสาเหตุที่พบมากที่สุดจะเป็น '{row['dominant_theme']}' "
                f"แต่เมื่อพิจารณาความถี่ร่วมกับระดับความรุนแรง "
                f"พบว่า '{row['recommended_campaign']}' มีความสำคัญมากกว่า "
                f"จึงควรได้รับการรณรงค์เป็นอันดับแรก"
            )

    def generate_insight(row):
        return (
            f"• Risk Level : {row['priority_level']} ({row['risk_score']:.1f})\n"
            f"• Dominant Theme : {row['dominant_theme']}\n"
            f"• Recommended Campaign : {row['recommended_campaign']}\n"
            f"• Supporting Campaign : {row['supporting_campaigns']}\n"
            f"• Recommendation : ควรเริ่มรณรงค์ด้าน '{row['recommended_campaign']}' "
            f"พร้อมดำเนินกิจกรรมด้าน '{row['supporting_campaigns']}' "
            f"ควบคู่กัน เพื่อช่วยลดความเสี่ยงของพื้นที่อย่างครอบคลุม\n"
            f"• Confidence : {row['confidence']:.1f}%"
        )

    area["reason"] = area.apply(generate_reason, axis=1)
    area["insight"] = area.apply(generate_insight, axis=1)

    area["insurance_recommendation"] = area.apply(
        lambda r: AREA_INSURANCE_MAP.get(
            (r["recommended_campaign"], str(r["priority_level"])), "PA พื้นฐาน"
        ),
        axis=1,
    )

    return area.sort_values("risk_score", ascending=False)


def build_cause_detail(df: pd.DataFrame, top_n: int = 3) -> dict:
    if (
        COL_AREA not in df.columns
        or "campaign_theme" not in df.columns
        or COL_CAUSE not in df.columns
    ):
        return {}

    counts = (
        df.dropna(subset=[COL_CAUSE, "campaign_theme"])
        .groupby([COL_AREA, "campaign_theme"])[COL_CAUSE]
        .value_counts()
    )

    detail_map: dict[tuple, str] = {}
    for (area_name, theme), sub in counts.groupby(level=[0, 1]):
        top = sub.sort_values(ascending=False).head(top_n)
        text = " • ".join(
            f"{cause} ({n} ครั้ง)" for (_, _, cause), n in top.items()
        )
        detail_map[(area_name, theme)] = text

    return detail_map


def attach_cause_detail(area: pd.DataFrame, df: pd.DataFrame, top_n: int = 3) -> pd.DataFrame:
    area = area.copy()
    if "recommended_campaign" not in area.columns:
        return area

    detail_map = build_cause_detail(df, top_n=top_n)
    area["recommended_campaign_detail"] = area.apply(
        lambda r: detail_map.get((r.name, r["recommended_campaign"]), "-"),
        axis=1,
    )
    return area


# ═══════════════════════════════════════════════════════════════════
# 6. RIDER-LEVEL INSURANCE & PERSONALIZED OUTPUT
# ═══════════════════════════════════════════════════════════════════

def apply_rider_insurance(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    def _get(row):
        return RIDER_INSURANCE_MAP.get(
            (row.get("campaign_theme"), row.get(COL_SEVERITY)),
            "PA พื้นฐาน (default)",
        )

    df["insurance_recommendation"] = df.apply(_get, axis=1)
    return df


def build_personalized_output(raw: pd.DataFrame, df: pd.DataFrame) -> pd.DataFrame:
    result = raw.copy()
    rename_map = {
        COL_CAUSE: "สาเหตุที่เกิดขึ้นจริง",
        "campaign_theme": "Campaign ที่แนะนำ",
        "insurance_recommendation": "ประกันที่แนะนำ",
    }
    for src_col, out_col in rename_map.items():
        if src_col in df.columns:
            result[out_col] = df[src_col].values

    return result


def join_personal_data(
    df_accident: pd.DataFrame,
    bank_folder: str,
    bank_filename: str,
    bank_sheet,
    join_key: str = COL_EMP_ID,
) -> pd.DataFrame:
    bank_path = Path(bank_folder) / bank_filename
    df_bank = pd.read_excel(bank_path, sheet_name=bank_sheet, engine="openpyxl")

    for d in [df_accident, df_bank]:
        if join_key in d.columns:
            d[join_key] = d[join_key].astype(str).str.strip()

    merged = df_accident.merge(df_bank, on=join_key, how="left", suffixes=("", "_bank"))
    return merged


# ═══════════════════════════════════════════════════════════════════
# 7. EXPORT & PIPELINE EXECUTION
# ═══════════════════════════════════════════════════════════════════

def _ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def export_outputs(
    area_summary: pd.DataFrame,
    df_rider: pd.DataFrame,
    unmapped: pd.DataFrame,
    df_joined: Optional[pd.DataFrame] = None,
    output_dir: str = OUTPUT_DIR,
    tag: str = "",
) -> dict[str, str]:
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    suffix = f"_{tag}_{_ts()}" if tag else f"_{_ts()}"
    saved = {}

    p = Path(output_dir) / OUT_AREA_FILE.replace(".xlsx", f"{suffix}.xlsx")
    area_summary.to_excel(p, index=True)
    saved["area_summary"] = str(p)

    rider_cols = [
        COL_AREA, COL_CAUSE, COL_SEVERITY,
        "campaign_theme", "insurance_recommendation",
        "period", "rider_exp_group", "age_group",
    ]
    out_rider = df_rider[[c for c in rider_cols if c in df_rider.columns]]
    p2 = Path(output_dir) / OUT_RIDER_FILE.replace(".xlsx", f"{suffix}.xlsx")
    out_rider.to_excel(p2, index=False)
    saved["rider_cleaned"] = str(p2)

    if not unmapped.empty:
        p3 = Path(output_dir) / OUT_UNMAPPED_FILE.replace(".csv", f"{suffix}.csv")
        unmapped.to_csv(p3, index=False, encoding="utf-8-sig")
        saved["unmapped"] = str(p3)

    if df_joined is not None and not df_joined.empty:
        p4 = Path(output_dir) / OUT_JOINED_FILE.replace(".xlsx", f"{suffix}.xlsx")
        df_joined.to_excel(p4, index=False)
        saved["joined_pii"] = str(p4)

    return saved


def run_pipeline(
    data_folder: str,
    data_filename: str,
    data_sheet,
    bank_folder: Optional[str] = None,
    bank_filename: Optional[str] = None,
    bank_sheet=0,
    output_dir: str = OUTPUT_DIR,
) -> dict:
    tag = f"{Path(data_filename).stem}_sheet{data_sheet}"

    raw = load_raw(data_folder, data_filename, data_sheet)
    df, col_report = clean(raw)
    df = feature_engineer(df)

    raw_with_id, _ = resolve_columns(raw.copy())

    df, unmapped = apply_theme_mapping(df)
    df = apply_rider_insurance(df)

    area_summary = build_area_summary(df)
    area_summary = attach_cause_detail(area_summary, df, top_n=3)

    df_joined = None
    if bank_folder and bank_filename:
        try:
            df_joined = join_personal_data(
                raw_with_id, bank_folder, bank_filename, bank_sheet
            )
        except Exception as e:
            print(f"⚠️ Join PII failed: {e}")

    saved = export_outputs(
        area_summary, df, unmapped, df_joined,
        output_dir=output_dir, tag=tag,
    )

    return {
        "df": df,
        "raw_with_id": raw_with_id,
        "area_summary": area_summary,
        "unmapped": unmapped,
        "saved_files": saved,
        "col_report": col_report,
        "tag": tag,
    }


def prepare_forecast_dataset(df: pd.DataFrame) -> pd.DataFrame:
    feature_cols = [
        "year",
        "quarter",
        COL_AREA,
        "rider_exp_group",
        COL_VEHICLE,
        COL_PERIOD,
        COL_VISIBILITY,
        COL_ROAD_SURFACE,
        COL_ROAD_TYPE,
        COL_TRAFFIC,
        COL_SPEED_GROUP,
        COL_4M1E,
    ]

    feature_cols = [c for c in feature_cols if c in df.columns]
    train_df = df[feature_cols + ["campaign_theme"]].copy()

    fill_cols = [
        "rider_exp_group",
        COL_VISIBILITY,
        COL_ROAD_SURFACE,
        COL_ROAD_TYPE,
        COL_TRAFFIC,
        COL_SPEED_GROUP,
    ]

    for c in fill_cols:
        if c in train_df.columns:
            train_df[c] = (
                train_df[c]
                .astype(str)
                .replace("nan", "Unknown")
                .replace("<NA>", "Unknown")
            )

    train_df = train_df.dropna(subset=["campaign_theme", COL_4M1E])
    return train_df

def clean_excel_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """ลบ Unnamed และคอลัมน์ที่เป็นค่าว่างทั้งหมด (แปลงชื่อคอลัมน์เป็น String ก่อนดักจับ)"""
    df = df.copy()
    valid_cols = [col for col in df.columns if not str(col).startswith("Unnamed")]
    df = df[valid_cols].dropna(axis=1, how="all")
    return df

def extract_quarter(df: pd.DataFrame) -> pd.DataFrame:
    """ค้นหาและแปลงคอลัมน์เดือนให้เป็น Quarter (Q1-Q4)"""
    df = df.copy()
    month_col = None
    for col in df.columns:
        if "เดือน" in str(col):
            month_col = col
            break

    if month_col is not None:
        clean_month_series = df[month_col].astype(str).str.strip().str.lower()
        df['quarter'] = clean_month_series.map(MONTH_MAP)
    else:
        df['quarter'] = np.nan
    return df

def find_cause_column(df: pd.DataFrame) -> str:
    """ค้นหาคอลัมน์สาเหตุที่แท้จริงแบบยืดหยุ่น"""
    for col in df.columns:
        if "สาเหตุที่แท้จริง" in str(col):
            return col
    return "สาเหตุที่แท้จริงจากการเกิดอุบัติเหตุ (เช่น คุยโทรศัพท์ ขับรถมือเดียว รถตัดหน้า ซ้อนท้าย )"

# ──────────────────────────────────────────────────────────────────────────
# CAUSE_THEME dict คัดลอกตรงจาก forecast_model_final.ipynb (cell รวม pipeline สุดท้าย)
# ใช้เฉพาะใน process_data_and_forecast() เท่านั้น — แยกจาก CAUSE_THEME ที่ import จาก
# config.py (ตัวนั้นใช้ร่วมกับ tab อื่นๆ ในแอป เช่น dashboard/insight/campaign)
# เพื่อการันตีว่าผลลัพธ์ตรงกับ notebook เป๊ะๆ (Accuracy 75.13% บนข้อมูลชุดเดียวกัน)
# ──────────────────────────────────────────────────────────────────────────
_NOTEBOOK_CAUSE_THEME = {
    "เบรกกะทันหัน": "Defensive Driving", "ขับรถย้อนศร": "Defensive Driving", "ซ้อนท้าย": "Defensive Driving",
    "เลี้ยวกระทันหัน": "Defensive Driving", "ฝ่าสัญญาณไฟจราจร": "Defensive Driving", "ตัดหน้า": "Defensive Driving",
    "รถตัดหน้า": "Defensive Driving", "คนเดินข้ามถนนตัดหน้ารถ": "Defensive Driving", "ชนท้าย": "Defensive Driving",
    "เปลี่ยนช่องทางกระทันหัน": "Defensive Driving", "เฉี่ยวชนคู่กรณี": "Defensive Driving", "แซงไม่พ้น": "Defensive Driving",
    "ใช้โทรศัพท์ขณะขับขี่": "Focus & Attention", "คุยโทรศัพท์": "Focus & Attention", "ผู้ขับขี่ผิดพลาดเอง(ตัดสินใจพลาด)": "Focus & Attention",
    "ง่วงนอน": "Focus & Attention", "ขับรถหลับใน": "Focus & Attention", "ประมาท": "Focus & Attention",
    "มองไม่เห็น": "Focus & Attention", "ไม่ดูกระจกมองข้าง": "Focus & Attention", "ไม่คุ้นชินเส้นทาง": "Focus & Attention",
    "เสียหลักล้ม": "Road & Vehicle Safety", "สัตว์ตัดหน้า": "Road & Vehicle Safety", "ถนนชำรุด": "Road & Vehicle Safety",
    "ถนนลื่น": "Road & Vehicle Safety", "ฝนตกถนนลื่น": "Road & Vehicle Safety", "หลุมบ่อ": "Road & Vehicle Safety",
    "เบรกไม่อยู่": "Road & Vehicle Safety", "ยางแตก": "Road & Vehicle Safety", "โซ่หลุด": "Road & Vehicle Safety",
    "ขับรถเร็ว": "Speed Awareness", "ขับรถเร็วเกินกำหนด": "Speed Awareness", "เข้าโค้งเร็ว": "Speed Awareness", "ออกตัวเร็ว": "Speed Awareness",
}

def map_theme_notebook(val) -> str:
    """แปลงสาเหตุ -> Campaign Theme โดยใช้ _NOTEBOOK_CAUSE_THEME (ดัก float/nan ไม่ให้พัง)"""
    if pd.isna(val) or val is None:
        return "Defensive Driving"
    val_str = str(val).strip()
    for k, v in _NOTEBOOK_CAUSE_THEME.items():
        if k in val_str:
            return v
    return "Defensive Driving"

def map_theme_safe(val) -> str:
    """แปลงสาเหตุเป็น Campaign Theme (ดัก float/nan ไม่ให้พัง)"""
    if pd.isna(val) or val is None:
        return "Defensive Driving"
    val_str = str(val).strip()
    for k, v in CAUSE_THEME.items():
        if str(k) in val_str:
            return v
    return "Defensive Driving"

def classify_4m1e_safe(row, cause_col: str) -> str:
    """จัดกลุ่ม 4M1E อัตโนมัติ (ดัก float/nan ไม่ให้พัง)"""
    cause_val = row.get(cause_col)
    surface_val = row.get('สภาพผิวจราจร')
    
    cause = "" if pd.isna(cause_val) else str(cause_val).lower()
    surface = "" if pd.isna(surface_val) else str(surface_val).lower()
    
    if any(k in cause for k in ['ถนน', 'ลื่น', 'ชำรุด', 'สัตว์', 'หลุม', 'ฝน']) or 'ลื่น' in surface:
        return 'Environment'
    if any(k in cause for k in ['ยางแตก', 'เบรกไม่อยู่', 'เบรกขัดข้อง', 'รถเสีย', 'โซ่หลุด']):
        return 'Machine'
    if any(k in cause for k in ['ย้อนศร', 'ผิดกฎ', 'แซง']):
        return 'Method'
    return 'Man'

def get_mode_safe(series: pd.Series, default_val: str = "Unknown") -> str:
    """หาค่า Mode อย่างปลอดภัย ป้องกัน IndexError"""
    m = series.dropna().mode()
    return str(m.iloc[0]) if not m.empty else default_val


# ==============================================================================
# MAIN PROCESSING & FORECASTING PIPELINE
# ==============================================================================
def process_data_and_forecast(file68_path, file69_path, sheet68=5, sheet69=0):
    """
    ฟังก์ชันหลักสำหรับให้ app.py เรียกใช้
    รับไฟล์ปี 68 และ 69 -> ทำการ Clean -> Train Model -> Forecast Q3-Q4 / 2569

    Returns
    -------
    final_forecast_df : pd.DataFrame
        ผลพยากรณ์ Q3-Q4 ระดับ Sub-Cause พร้อม Q1-Q2 YoY Trend Multiplier
    df : pd.DataFrame
        ข้อมูลรวม (ปี 68 + 69) หลังทำความสะอาดและ feature engineering แล้ว
    accuracy : float
        Model Accuracy (%) จากการประเมินด้วย train/test split (80/20)
    feature_importance_df : pd.DataFrame
        Feature importance (%) ของแต่ละตัวแปร เรียงจากมากไปน้อย
    """
    # 1. Load Data
    df68 = pd.read_excel(file68_path, sheet_name=sheet68)
    df69 = pd.read_excel(file69_path, sheet_name=sheet69)

    # 2. Clean Excel structure
    df68 = clean_excel_dataframe(df68)
    df69 = clean_excel_dataframe(df69)

    # Align columns
    df69 = df69.rename(columns=RENAME_DICT_69)
    df69 = df69.drop(columns=["หลักสูตรความปลอดภัย6ชม.(ไม่เก็บเงิน)"], errors="ignore")

    df68['year'] = 2568
    df69['year'] = 2569

    # 3. Extract Quarter
    df68 = extract_quarter(df68)
    df69 = extract_quarter(df69)

    # 4. Merge Data
    df = pd.concat([df68, df69], ignore_index=True)

    # Clean String Fields
    for col in df.select_dtypes(include=['object', 'string']).columns:
        df[col] = df[col].astype(str).str.strip().replace({"": np.nan, "nan": np.nan, "<NA>": np.nan})

    # 5. Dynamic Cause Column & Feature Engineering
    cause_col = find_cause_column(df)
    df['campaign_theme'] = df[cause_col].apply(map_theme_notebook)
    df['4M1E_Cleaned'] = df.apply(lambda row: classify_4m1e_safe(row, cause_col), axis=1)

    feature_cols = ['พื้นที่', 'ทัศนวิสัย', 'สภาพผิวจราจร', 'ลักษณะเส้นทาง', 'สภาพการจราจร', '4M1E_Cleaned']
    for c in feature_cols:
        if c not in df.columns:
            df[c] = "Unknown"
        else:
            df[c] = df[c].fillna("Unknown")

    # 6. Prepare Model
    model_features = ['quarter', 'พื้นที่', 'ทัศนวิสัย', 'สภาพผิวจราจร', 'ลักษณะเส้นทาง', 'สภาพการจราจร', '4M1E_Cleaned']
    X = df[model_features].astype(str).copy()
    y = df['campaign_theme'].astype(str).copy()

    encoder = OrdinalEncoder(handle_unknown='use_encoded_value', unknown_value=-1)
    X_encoded = pd.DataFrame(encoder.fit_transform(X), columns=model_features)

    # 6b. Model Evaluation (train/test split) — ported from notebook cell 24
    #     ใช้ประเมิน Accuracy และ Feature Importance โดยไม่กระทบโมเดลที่ใช้พยากรณ์จริง (fit บน full data ด้านล่าง)
    try:
        X_train, X_test, y_train, y_test = train_test_split(
            X_encoded, y, test_size=0.2, random_state=42, stratify=y
        )
    except ValueError:
        # stratify ล้มเหลวถ้าบาง class มีข้อมูลน้อยเกินไป -> fallback แบบไม่ stratify
        X_train, X_test, y_train, y_test = train_test_split(
            X_encoded, y, test_size=0.2, random_state=42
        )

    eval_model = RandomForestClassifier(n_estimators=200, max_depth=10, random_state=42, n_jobs=-1)
    eval_model.fit(X_train, y_train)
    y_pred = eval_model.predict(X_test)
    accuracy = round(accuracy_score(y_test, y_pred) * 100, 2)

    feature_importance_df = pd.DataFrame({
        "Feature": model_features,
        "Importance (%)": (eval_model.feature_importances_ * 100).round(2),
    }).sort_values(by="Importance (%)", ascending=False).reset_index(drop=True)

    # Train Random Forest (final model, fit on ALL data — ใช้พยากรณ์ Q3-Q4 จริง)
    rf_model = RandomForestClassifier(n_estimators=200, max_depth=10, random_state=42, n_jobs=-1)
    rf_model.fit(X_encoded, y)

    # 7. Forecasting (Q3 & Q4 / 2569)
    areas = df['พื้นที่'].dropna().unique()
    quarters = ['Q3', 'Q4']
    predict_rows = []

    for q in quarters:
        for a in areas:
            area_df = df[df['พื้นที่'] == a]
            if len(area_df) == 0: 
                continue
            
            predict_rows.append({
                'quarter': q,
                'พื้นที่': a,
                'ทัศนวิสัย': get_mode_safe(area_df['ทัศนวิสัย']),
                'สภาพผิวจราจร': get_mode_safe(area_df['สภาพผิวจราจร']),
                'ลักษณะเส้นทาง': get_mode_safe(area_df['ลักษณะเส้นทาง']),
                'สภาพการจราจร': get_mode_safe(area_df['สภาพการจราจร']),
                '4M1E_Cleaned': get_mode_safe(area_df['4M1E_Cleaned'])
            })

    future_df = pd.DataFrame(predict_rows).astype(str)
    future_encoded = pd.DataFrame(encoder.transform(future_df[model_features]), columns=model_features)

    probs = rf_model.predict_proba(future_encoded)
    classes = rf_model.classes_

    forecast_results = []

    for idx, row in future_df.iterrows():
        area = row['พื้นที่']
        q = row['quarter']
        
        for c_idx, theme in enumerate(classes):
            theme_prob = probs[idx, c_idx]
            area_df = df[df['พื้นที่'] == area]
            
            # คำนวณ YoY Trend Factor (Q1-Q2 68 vs 69)
            q12_68 = area_df[(area_df['year'] == 2568) & (area_df['quarter'].isin(['Q1', 'Q2'])) & (area_df['campaign_theme'] == theme)]
            q12_69 = area_df[(area_df['year'] == 2569) & (area_df['quarter'].isin(['Q1', 'Q2'])) & (area_df['campaign_theme'] == theme)]
            
            total_68 = len(area_df[area_df['year'] == 2568])
            total_69 = len(area_df[area_df['year'] == 2569])
            
            trend_factor = 1.0
            if total_68 > 0 and total_69 > 0:
                ratio_68 = len(q12_68) / total_68
                ratio_69 = len(q12_69) / total_69
                if ratio_68 > 0:
                    trend_factor = ratio_69 / ratio_68
                    trend_factor = np.clip(trend_factor, 0.8, 1.5)
                elif ratio_69 > 0:
                    trend_factor = 1.2

            area_causes = area_df[area_df['campaign_theme'] == theme]
            if len(area_causes) > 0:
                top_causes = area_causes[cause_col].value_counts(normalize=True).head(3)
                
                for cause_name, ratio in top_causes.items():
                    adjusted_cause_prob = theme_prob * ratio * trend_factor * 100
                    
                    forecast_results.append({
                        'Quarter': q,
                        'พื้นที่': area,
                        'Predicted_Theme': theme,
                        'Base_Theme_Prob (%)': round(theme_prob * 100, 2),
                        'Q1-Q2_Trend_Multiplier': round(trend_factor, 2),
                        'Sub_Cause (สาเหตุย่อย)': cause_name,
                        'Estimated_Cause_Probability (%)': round(adjusted_cause_prob, 2)
                    })

    final_forecast_df = pd.DataFrame(forecast_results)
    return final_forecast_df, df, accuracy, feature_importance_df