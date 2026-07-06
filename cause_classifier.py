import re
import pandas as pd

try:
    from rapidfuzz import process as fuzz_process, fuzz
    _RAPIDFUZZ_AVAILABLE = True
except ImportError:
    _RAPIDFUZZ_AVAILABLE = False

_FUZZY_THEME_THRESHOLD = 85

# ═══════════════════════════════════════════════════════════════════
# 1. THEME DEFINITION — key = ชื่อ theme, value = set ของคำย่อยที่บ่งบอก theme นั้น
#    เพิ่ม/แก้ตรงนี้ได้เรื่อยๆ ตามข้อมูลจริงที่เจอ ไม่ต้องแก้ logic
# ═══════════════════════════════════════════════════════════════════
THEME_KEYWORDS: dict[str, set[str]] = {
    "ขับรถกะทันหัน/เปลี่ยนช่องทาง": {
        "เลี้ยว", "กลับรถ", "เปลี่ยนช่องทาง", "แซง", "ตัดหน้า",
    },
    "ชนคู่กรณี": {
        "คู่กรณี", "ชนท้าย", "เฉี่ยวชน", "เฉี่ยว", "เกี่ยว", "เบียดชน",
    },
    "ชนคนเดินเท้า/จักรยาน": {
        "คนเดินเท้า", "จักรยาน", "คนข้ามถนน", "คนปั่น",
    },
    "สภาพถนน/สิ่งแวดล้อม": {
        "ถนนชำรุด", "ขรุขระ", "ถนนลื่น", "พื้นลื่น",
        "ทราย", "น้ำมัน", "หิน", "ดิน",
    },
    "สภาพร่างกายผู้ขับ": {
        "ความพร้อมของร่างกาย", "วูบ", "หลับใน", "อ่อนเพลีย",
    },
    "สภาพรถ": {
        "ความพร้อมของรถ", "ขาตั้ง", "รถลื่นล้ม",
    },
    "ชนสัตว์/สิ่งกีดขวาง": {
        "สัตว์", "ทรัพย์สินอื่น",
    },
    "ตกร่อง/คลอง": {
        "ตกร่องน้ำ", "คลอง", "คูน้ำ",
    },
    "ความประมาท/ไม่ระวัง": {
        "ไม่คุ้นชิน", "ไม่เว้นระยะ", "แสงสะท้อน",
        "ตกใจหักหลบ", "สะดุดล้ม", "ถอยชน",
    },
    "ถูกชน (ไม่ใช่ฝ่ายก่อเหตุ)": {
        "ถูกชน", "เปิดประตู",
    },
}

UNMAPPED_LABEL = "อื่นๆ (ต้องตรวจสอบ)"


def _normalize_and_split(cause: str) -> list[str]:
    """ตัดข้อความในวงเล็บออก แล้วแตกด้วย / เป็น token ย่อย"""
    cause_clean = re.sub(r"\(.*?\)", "", str(cause))
    tokens = [t.strip() for t in cause_clean.split("/") if t.strip()]
    return tokens if tokens else [cause_clean.strip()]


def classify_cause(cause: str) -> tuple[str, str]:
    """
    classify cause 1 ตัว → (theme, debug_report)
    ลำดับ: exact token match -> fuzzy token match -> UNMAPPED
    """
    if not cause or str(cause).strip() == "" or str(cause).lower() == "nan":
        return UNMAPPED_LABEL, "empty input"

    tokens = _normalize_and_split(cause)

    # 1. exact token match (เร็วสุด แม่นสุด)
    for token in tokens:
        for theme, keywords in THEME_KEYWORDS.items():
            if any(kw in token for kw in keywords):
                return theme, f"exact token '{token}'"

    # 2. fuzzy token match (จับ variation การพิมพ์/สลับคำเล็กน้อย)
    if _RAPIDFUZZ_AVAILABLE:
        best_theme, best_score, best_token, best_kw = None, 0, "", ""
        for token in tokens:
            for theme, keywords in THEME_KEYWORDS.items():
                for kw in keywords:
                    score = fuzz.partial_ratio(token, kw)
                    if score > best_score:
                        best_score, best_theme = score, theme
                        best_token, best_kw = token, kw
        if best_score >= _FUZZY_THEME_THRESHOLD:
            return best_theme, f"fuzzy {best_score:.0f}% '{best_token}'≈'{best_kw}'"

    return UNMAPPED_LABEL, f"no match, tokens={tokens}"


def classify_cause_series(cause_series) -> tuple["pd.Series", dict]:
    """
    รับ pandas Series ของ cause -> คืน (Series ของ theme, report dict สำหรับ debug)
    report: {cause_text: (theme, debug_reason)}  — ใช้ print/log ดูว่า match เพราะอะไร
    """
    import pandas as pd

    unique_causes = cause_series.dropna().unique().tolist()
    report: dict[str, tuple[str, str]] = {}
    mapping: dict[str, str] = {}

    for cause in unique_causes:
        theme, reason = classify_cause(cause)
        mapping[cause] = theme
        report[cause] = (theme, reason)

    result = cause_series.map(mapping).fillna(UNMAPPED_LABEL)
    return result, report