# analysis.py — pipeline functions (ไม่มี Streamlit dependency ทั้งหมด)
# ใช้ได้ทั้งใน app.py และ notebook
# ─────────────────────────────────────────────────────────────────
# from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

from config import (
    AREA_INSURANCE_MAP,
    CAUSE_THEME,
    COL_AGE,
    COL_AREA,
    COL_CAUSE,
    COL_DRIVER_EXP,
    COL_EMP_ID,
    COL_LEAVE,
    COL_RIDER_EXP,
    COL_SEVERITY,
    COL_SPEED,
    COL_TIME,
    DROP_COLS,
    FUZZY_THRESHOLD,
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
    files = sorted(
        f.name for f in p.glob("*.xls*") if not f.name.startswith("~$")
    )
    return files


def get_sheet_names(folder: str, filename: str) -> list[str]:
    """
    คืน list ชื่อ sheet ในไฟล์ Excel
    ใช้ ExcelFile เพื่อไม่ต้องโหลด data ทั้งหมด
    """
    path = Path(folder) / filename
    with pd.ExcelFile(path, engine="openpyxl") as xf:
        return xf.sheet_names


def load_raw(folder: str, filename: str, sheet) -> pd.DataFrame:
    """
    โหลด sheet จาก Excel → DataFrame ดิบ
    sheet: ชื่อ (str) หรือ index (int)
    """
    path = Path(folder) / filename
    df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    return df


# ═══════════════════════════════════════════════════════════════════
# 2. COLUMN RESOLVER (fuzzy-safe — กัน \n, space เกิน, พิมพ์เพี้ยน)
# ═══════════════════════════════════════════════════════════════════

_REQUIRED_COLS = [COL_AREA, COL_CAUSE, COL_SEVERITY]
_OPTIONAL_COLS = [COL_AGE, COL_SPEED, COL_LEAVE, COL_TIME, COL_RIDER_EXP, COL_DRIVER_EXP]
_FUZZY_COL_THRESHOLD = 80


def _col_key(text: str) -> str:
    """ลบ whitespace ทั้งหมด → เปรียบเทียบแบบ whitespace-agnostic"""
    return re.sub(r"\s+", "", str(text))


def resolve_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    1. strip whitespace/\\n จากชื่อ column ทั้งหมด
    2. fuzzy match column ที่ยังหาไม่เจอ
    3. raise ValueError ชัดๆ ถ้า required column หายไป
    คืน (df ที่ rename แล้ว, col_report สำหรับ debug)
    """
    df = df.copy()

    # Step 1: strip whitespace/\n
    df.columns = pd.Index([
        re.sub(r"\s+", " ", str(c)).strip() for c in df.columns
    ])

    all_targets = _REQUIRED_COLS + _OPTIONAL_COLS + DROP_COLS
    rename_map: dict[str, str] = {}
    col_report: dict[str, str] = {}
    missing_required: list[str] = []
    actual_key_map = {_col_key(c): c for c in df.columns}

    for target in all_targets:
        target_key = _col_key(target)

        # exact match (whitespace-agnostic)
        if target_key in actual_key_map:
            actual = actual_key_map[target_key]
            if actual != target:
                rename_map[actual] = target
                col_report[target] = f"normalized '{actual}'"
            else:
                col_report[target] = "exact"
            continue

        # fuzzy match
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
# 3. CLEANING
# ═══════════════════════════════════════════════════════════════════

def clean(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    2.0 resolve_columns (fuzzy-safe) — หยุดทันทีถ้า required col หายไป
    2.1 Drop PII / admin columns
    2.2 Strip whitespace + normalize nan strings
    2.3 Convert numeric cols (age, speed, leave)
    คืน (df, col_report)
    """
    df = df.copy()

    # 2.0 resolve (normalize + fuzzy + error)
    df, col_report = resolve_columns(df)

    # 2.1 Drop PII + admin
    drop_normalized = {_col_key(c) for c in DROP_COLS}
    actual_drop = [c for c in df.columns if _col_key(c) in drop_normalized]
    df = df.drop(columns=actual_drop, errors="ignore")

    # 2.2 Strip + normalize string
    str_cols = df.select_dtypes("object").columns
    for col in str_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.strip()
            .replace({"": np.nan, "nan": np.nan, "<NA>": np.nan, "None": np.nan})
        )

    # 2.3 Convert numeric
    for col in [COL_AGE, COL_SPEED, COL_LEAVE]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df, col_report


# ═══════════════════════════════════════════════════════════════════
# 3. FEATURE ENGINEERING
# ═══════════════════════════════════════════════════════════════════

def _exp_to_months(text) -> float:
    """แปลงข้อความประสบการณ์ → จำนวนเดือน"""
    if pd.isna(text):
        return np.nan
    text = str(text)
    y = int(m.group(1)) if (m := re.search(r"(\d+)ปี", text))      else 0
    mo= int(m.group(1)) if (m := re.search(r"(\d+)เดือน", text))   else 0
    d = int(m.group(1)) if (m := re.search(r"(\d+)วัน", text))     else 0
    w = int(m.group(1)) if (m := re.search(r"(\d+)สัปดาห์", text)) else 0
    return y * 12 + mo + d / 30 + w / 4


def _extract_hour(text) -> float:
    if pd.isna(text):
        return np.nan
    match = re.search(r"(\d{1,2})[:\.](\d{1,2})", str(text).strip())
    return int(match.group(1)) if match else np.nan


def _get_period(hour) -> str | float:
    if pd.isna(hour): return np.nan
    if hour < 6:      return "Late Night"
    if hour < 12:     return "Morning"
    if hour < 18:     return "Afternoon"
    return "Night"


def feature_engineer(df: pd.DataFrame) -> pd.DataFrame:
    """เพิ่ม derived columns: exp_month, exp_group, age_group, hour, period"""
    df = df.copy()

    # Experience → months
    if COL_RIDER_EXP in df.columns:
        df["rider_exp_month"] = df[COL_RIDER_EXP].apply(_exp_to_months)
        df["rider_exp_group"] = pd.cut(
            df["rider_exp_month"],
            bins=[-1, 3, 12, 24, 999],
            labels=["<3m", "3-12m", "1-2y", ">2y"],
        )

    if COL_DRIVER_EXP in df.columns:
        df["driver_exp_month"] = df[COL_DRIVER_EXP].apply(_exp_to_months)

    # Age group
    if COL_AGE in df.columns:
        df["age_group"] = pd.cut(
            df[COL_AGE],
            bins=[0, 20, 25, 30, 100],
            labels=["<20", "20-25", "26-30", ">30"],
        )

    # Time → hour + period
    if COL_TIME in df.columns:
        df["hour"]   = df[COL_TIME].apply(_extract_hour)
        df["period"] = df["hour"].apply(_get_period)

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
    """ตัดข้อความในวงเล็บออก แล้วแตก '/' เป็น token ย่อย
    กันเคส cause ที่เขียนคำสลับที่/ไม่ครบคำ เช่น
    'เลี้ยว/กลับรถ/เปลี่ยนช่องทางกะทันหัน' vs 'เลี้ยว/เปลี่ยนช่องทางกะทันหัน'"""
    text_clean = re.sub(r"\(.*?\)", "", str(text))
    tokens = [t.strip() for t in text_clean.split("/") if t.strip()]
    return tokens if tokens else [text_clean.strip()]


def _token_theme(text: str) -> str | None:
    """แตก cause เป็น token ย่อยก่อน แล้วเทียบแต่ละ token กับ key ของ CAUSE_THEME
    ด้วย substring สองทาง (token อยู่ใน key หรือ key อยู่ใน token)"""
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
    """เหมือน _fuzzy_theme แต่เทียบทีละ token ย่อย ไม่เทียบทั้งประโยค
    จับเคสคำในแต่ละ token สะกดเพี้ยน/ต่างจาก key เล็กน้อย"""
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
    """Exact match (ทั้งประโยค) → token exact match → fuzzy (ทั้งประโยค) → fuzzy token → None"""
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
    """
    เพิ่ม column 'campaign_theme' ใน df
    คืน (df, unmapped_df)  ← unmapped_df = rows ที่ map ไม่ได้
    """
    df = df.copy()
    if COL_CAUSE not in df.columns:
        df["campaign_theme"] = np.nan
        return df, pd.DataFrame()

    df["campaign_theme"] = df[COL_CAUSE].apply(map_theme)

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
    """หา column 'หยุดงานเกิน3วัน' แบบ safe fallback"""
    target = "หยุดงานเกิน3วัน"
    if target in severity_table.columns:
        return target
    candidates = [c for c in severity_table.columns if "เกิน" in c]
    return candidates[0] if candidates else severity_table.columns[-1]


def build_area_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    รัน Risk Score pipeline ทั้งหมด → คืน area_summary DataFrame
    Columns:
      severity_pct, defensive_pct, focus_pct, road_pct, speed_pct,
      *_score, risk_score, priority_level,
      dominant_theme, recommended_campaign, supporting_campaigns,
      confidence, reason, insight, insurance_recommendation
    """
    if COL_AREA not in df.columns or "campaign_theme" not in df.columns:
        return pd.DataFrame()

    # 5.1 Theme % per area
    ct = pd.crosstab(df[COL_AREA], df["campaign_theme"])
    for t in THEMES:
        if t not in ct.columns:
            ct[t] = 0
    theme_pct = (ct.div(ct.sum(axis=1), axis=0) * 100).round(1)

    # 5.2 Severity % per area
    sev_table = (
        pd.crosstab(df[COL_AREA], df[COL_SEVERITY], normalize="index") * 100
    ).round(1)
    sev_col = _find_sev_col(sev_table)
    severity_pct = sev_table[sev_col]

    # 5.3 Impact map (theme → % หยุดงานเกิน3วัน)
    theme_impact_table = (
        pd.crosstab(df["campaign_theme"], df[COL_SEVERITY], normalize="index") * 100
    )
    if sev_col not in theme_impact_table.columns:
        theme_impact_table[sev_col] = 0
    impact_map = theme_impact_table[sev_col].to_dict()

    # 5.4 Build summary
    PCT_MAP = {
        "Defensive Driving":    "defensive_pct",
        "Focus & Attention":    "focus_pct",
        "Road & Vehicle Safety": "road_pct",
        "Speed Awareness":       "speed_pct",
    }
    SCORE_MAP = {v.replace("_pct", "_score"): k for k, v in PCT_MAP.items()}

    area = pd.DataFrame({
        "severity_pct":  severity_pct,
        "defensive_pct": theme_pct.get("Defensive Driving",    0),
        "focus_pct":     theme_pct.get("Focus & Attention",    0),
        "road_pct":      theme_pct.get("Road & Vehicle Safety", 0),
        "speed_pct":     theme_pct.get("Speed Awareness",       0),
    })

    score_cols = []
    for pct_col, theme in {v: k for k, v in PCT_MAP.items()}.items():
        sc = pct_col.replace("_pct", "_score")
        area[sc] = (area[pct_col] * impact_map.get(theme, 0)).round(2)
        score_cols.append(sc)

    # Risk score (weighted)
    area["risk_score"] = (
        area["severity_pct"] * RISK_WEIGHTS["severity_pct"]
        + area["road_pct"]   * RISK_WEIGHTS["road_pct"]
        + area["speed_pct"]  * RISK_WEIGHTS["speed_pct"]
    ).round(2)

    area["priority_level"] = pd.cut(
        area["risk_score"],
        bins=PRIORITY_BINS,
        labels=PRIORITY_LABELS,
    )

    # 5.5 Campaign ranking
    pct_to_theme = {v: k for k, v in PCT_MAP.items()}
    score_to_theme = {v.replace("_pct", "_score"): k for k, v in PCT_MAP.items()}

    area["dominant_theme"] = (
        area[list(PCT_MAP.values())].idxmax(axis=1).map(pct_to_theme)
    )

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

    # 5.6 Confidence (top1 / top1+top2)
    sorted_scores = np.sort(area[score_cols].values, axis=1)
    top1, top2 = sorted_scores[:, -1], sorted_scores[:, -2]
    denom = top1 + top2
    area["confidence"] = np.where(
        denom > 0, (top1 / denom * 100).round(1), np.nan
    )

    # 5.7 Reason + Insight
    def generate_reason(row):
        if row["dominant_theme"] == row["recommended_campaign"]:
        # theme ที่เจอบ่อยสุด กับ theme ที่แนะนำ เป็นตัวเดียวกัน → ใช้ "เช่นกัน" ไม่ใช่ "แต่"
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

    # 5.8 Insurance recommendation
    area["insurance_recommendation"] = area.apply(
        lambda r: AREA_INSURANCE_MAP.get(
            (r["recommended_campaign"], str(r["priority_level"])), "PA พื้นฐาน"
        ),
        axis=1,
    )

    return area.sort_values("risk_score", ascending=False)
# ═══════════════════════════════════════════════════════════════════
# 5.9 CAUSE DETAIL (ใช้ประกอบ recommended_campaign ให้เห็น cause จริง)
# ═══════════════════════════════════════════════════════════════════

def build_cause_detail(df: pd.DataFrame, top_n: int = 3) -> dict:
    """
    หา top cause จริง (ไม่ใช่แค่ theme) แยกตาม (area, theme)
    ใช้ประกอบ recommended_campaign ให้คนทำ campaign เห็น detail จริง
    ไม่ใช่แค่ชื่อ theme กว้างๆ อย่าง 'Defensive Driving'

    คืน dict {(area, theme): "cause1 (n ครั้ง) • cause2 (n ครั้ง) • ..."}
    """
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
    """
    เพิ่ม column 'recommended_campaign_detail' เข้า area_summary
    โดยดึง top cause จริงของ (area, recommended_campaign) มาต่อท้าย
    เรียกต่อจาก build_area_summary() ได้เลย
    """
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
# 6. RIDER-LEVEL INSURANCE
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
# ═══════════════════════════════════════════════════════════════════
# 6.5 PERSONALIZED OUTPUT (แนบชื่อพนักงานกลับเข้า recommendation)
# ═══════════════════════════════════════════════════════════════════

def build_personalized_output(raw: pd.DataFrame, df: pd.DataFrame) -> pd.DataFrame:
    """
    เอาไฟล์ดิบที่มีชื่อ-รหัสพนักงานอยู่แล้ว (raw)
    มาแปะ column recommendation ที่คำนวณจาก df (cleaned) ต่อท้าย
      - Campaign ที่ควรได้รับ (มาจาก cause จริงที่เกิดกับคนนั้น)
      - ประกันที่แนะนำ (มาจาก theme + ความรุนแรงของคนนั้น)
    ไม่ต้อง join กับไฟล์อื่นเลย เพราะ raw กับ df แถวเรียงตรงกันอยู่แล้ว
    """
    result = raw.copy()
    rename_map = {
        COL_CAUSE:                 "สาเหตุที่เกิดขึ้นจริง",
        "campaign_theme":          "Campaign ที่แนะนำ",
        "insurance_recommendation": "ประกันที่แนะนำ",
    }
    for src_col, out_col in rename_map.items():
        if src_col in df.columns:
            result[out_col] = df[src_col].values

    return result

# ═══════════════════════════════════════════════════════════════════
# 7. JOIN WITH PERSONAL DATA (PII — save to output only)
# ═══════════════════════════════════════════════════════════════════

def join_personal_data(
    df_accident: pd.DataFrame,
    bank_folder: str,
    bank_filename: str,
    bank_sheet,
    join_key: str = COL_EMP_ID,
) -> pd.DataFrame:
    """
    Join df_accident กับข้อมูลส่วนตัว (bank.xlsx)
    คืน df ที่ join แล้ว — **ไม่** นำขึ้น dashboard
    """
    bank_path = Path(bank_folder) / bank_filename
    df_bank = pd.read_excel(bank_path, sheet_name=bank_sheet, engine="openpyxl")

    # normalize join key
    for d in [df_accident, df_bank]:
        if join_key in d.columns:
            d[join_key] = d[join_key].astype(str).str.strip()

    merged = df_accident.merge(df_bank, on=join_key, how="left", suffixes=("", "_bank"))
    return merged


# ═══════════════════════════════════════════════════════════════════
# 8. EXPORT
# ═══════════════════════════════════════════════════════════════════

def _ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def export_outputs(
    area_summary: pd.DataFrame,
    df_rider: pd.DataFrame,
    unmapped: pd.DataFrame,
    df_joined: Optional[pd.DataFrame] = None,
    output_dir: str = OUTPUT_DIR,
    tag: str = "",                 # e.g.  "2568_Sheet2"
) -> dict[str, str]:
    """
    บันทึกไฟล์ทั้งหมดใน output_dir
    คืน dict {label: filepath} สำหรับแสดงบน Streamlit
    """
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    suffix = f"_{tag}_{_ts()}" if tag else f"_{_ts()}"
    saved = {}

    # Area summary
    p = Path(output_dir) / OUT_AREA_FILE.replace(".xlsx", f"{suffix}.xlsx")
    area_summary.to_excel(p, index=True)
    saved["area_summary"] = str(p)

    # Rider cleaned (no PII)
    rider_cols = [
        COL_AREA, COL_CAUSE, COL_SEVERITY,
        "campaign_theme", "insurance_recommendation",
        "period", "rider_exp_group", "age_group",
    ]
    out_rider = df_rider[[c for c in rider_cols if c in df_rider.columns]]
    p2 = Path(output_dir) / OUT_RIDER_FILE.replace(".xlsx", f"{suffix}.xlsx")
    out_rider.to_excel(p2, index=False)
    saved["rider_cleaned"] = str(p2)

    # Unmapped causes
    if not unmapped.empty:
        p3 = Path(output_dir) / OUT_UNMAPPED_FILE.replace(".csv", f"{suffix}.csv")
        unmapped.to_csv(p3, index=False, encoding="utf-8-sig")
        saved["unmapped"] = str(p3)

    # Joined with personal data (PII)
    if df_joined is not None and not df_joined.empty:
        p4 = Path(output_dir) / OUT_JOINED_FILE.replace(".xlsx", f"{suffix}.xlsx")
        df_joined.to_excel(p4, index=False)
        saved["joined_pii"] = str(p4)

    return saved


# ═══════════════════════════════════════════════════════════════════
# 9. FULL PIPELINE (one-call convenience)
# ═══════════════════════════════════════════════════════════════════

def run_pipeline(
    data_folder: str,
    data_filename: str,
    data_sheet,
    bank_folder: Optional[str] = None,
    bank_filename: Optional[str] = None,
    bank_sheet=0,
    output_dir: str = OUTPUT_DIR,
) -> dict:
    """
    รัน pipeline ทั้งหมด → คืน dict ที่ app.py ใช้แสดงผล

    Returns
    -------
    {
        "df":           pd.DataFrame,   # rider-level (no PII)
        "raw_with_id":  pd.DataFrame,   # for join_personal_data (PII) only
        "area_summary": pd.DataFrame,
        "unmapped":     pd.DataFrame,
        "saved_files":  dict[str, str],
        "tag":          str,
    }
    """
    tag = f"{Path(data_filename).stem}_sheet{data_sheet}"

    # Load + clean + feature eng
    raw = load_raw(data_folder, data_filename, data_sheet)
    df, col_report = clean(raw)
    df  = feature_engineer(df)

    raw_with_id, _ = resolve_columns(raw.copy())

    # Theme mapping
    df, unmapped = apply_theme_mapping(df)

    # Rider insurance
    df = apply_rider_insurance(df)

    # Area summary
    area_summary = build_area_summary(df)
    area_summary = attach_cause_detail(area_summary, df, top_n=3)

    # Optional PII join (save to output only — ไม่ return ขึ้น dashboard)
    df_joined = None
    if bank_folder and bank_filename:
        try:
            df_joined = join_personal_data(
                raw_with_id, bank_folder, bank_filename, bank_sheet
            )
        except Exception as e:
            print(f"⚠️  Join PII failed: {e}")

    # Export
    saved = export_outputs(
        area_summary, df, unmapped, df_joined,
        output_dir=output_dir, tag=tag,
    )

    return {
        "df":           df,
        "raw_with_id":  raw_with_id,
        "area_summary": area_summary,
        "unmapped":     unmapped,
        "saved_files":  saved,
        "col_report":   col_report,
        "tag":          tag,
    }