# config.py — centralized constants for the Rider Accident pipeline
# ─────────────────────────────────────────────────────────────────

from pathlib import Path

# ── Column names (ตรงกับ Excel จริง) ─────────────────────────────
COL_AREA         = "พื้นที่"
COL_CAUSE        = "สาเหตุที่แท้จริงจากการเกิดอุบัติเหตุ (เช่น คุยโทรศัพท์ ขับรถมือเดียว รถตัดหน้า ซ้อนท้าย )"
COL_SEVERITY     = "ระดับความรุนแรงของการหยุดงาน"
COL_AGE          = "อายุพนักงาน(ใส่ตัวเลข)"
COL_SPEED        = "ความเร็วรถขณะเกิดเหตุ \n(ใส่เฉพาะตัวเลข)"
COL_LEAVE        = "วันหยุดงาน (ใส่เฉพาะตัวเลข)"
COL_TIME         = "เวลาเกิดเหตุ"
COL_RIDER_EXP    = "ประสบการณ์ขับรถเป็นRider (มาแล้วกี่เดือนกี่ปี)"
COL_DRIVER_EXP   = "ประสบการณ์ขับรถ (ขับรถมาแล้วกี่เดือนกี่ปี)"
COL_EMP_ID       = "รหัสพนักงาน (รหัสต้องครบ7หลัก)"  # PII — ใช้ join เท่านั้น

# ── Columns to drop before dashboard (PII + admin) ───────────────
DROP_COLS = [
    "ลำดับ",
    "รหัสพนักงาน (รหัสต้องครบ7หลัก)",
    "ชื่อ",
    "สกุล",
    "ทะเบียนรถ\n(ติดกันทุกตัว)",
    "รหัสสาขา",
    "สาขาร้าน",
    "ฝ่าย",
    "เขต",
    "FC",
    "อบรมหลักสูตร คปภ.",
    "ประเภทนศ.ทวีภาคี",
]

# ── 4 campaign themes ─────────────────────────────────────────────
THEMES = [
    "Defensive Driving",
    "Focus & Attention",
    "Road & Vehicle Safety",
    "Speed Awareness",
]

# ── Cause → Theme mapping (exact match, then fuzzy fallback) ──────
CAUSE_THEME = {
    # Defensive Driving
    "เบรกกะทันหัน": "Defensive Driving",
    "ขับรถย้อนศร": "Defensive Driving",
    "นั่งซ้อนท้าย": "Defensive Driving",
    "ซ้อนท้าย": "Defensive Driving",
    "เลี้ยวกระทันหัน": "Defensive Driving",
    "เลี้ยวกะทันหัน": "Defensive Driving",
    "ฝ่าสัญญาณไฟจราจร": "Defensive Driving",
    "เว้นระยะห่างน้อยเกินไป": "Defensive Driving",
    "ตัดหน้า": "Defensive Driving",
    "รถตัดหน้า": "Defensive Driving",
    "พนง.ถูกชน (พนง.เป็นฝ่ายถูก)": "Defensive Driving",
    "คนเดินข้ามถนนตัดหน้ารถ": "Defensive Driving",
    "คนเดินข้ามถนน": "Defensive Driving",
    "พนง.ออกรถกะทันหัน(ไม่มองรถ/ถนน)": "Defensive Driving",
    "สตาร์ทออกตัวกะทันหัน": "Defensive Driving",
    "เปลี่ยนช่องทางกระทันหัน": "Defensive Driving",
    "ขนส่งสินค้าจำนวนมาก": "Defensive Driving",
    "แขวนสินค้ากับแฮนด์รถ": "Defensive Driving",
    "ขับรถมือเดียว": "Defensive Driving",
    "ชนท้าย": "Defensive Driving",
    "ชนรถจอดข้างทาง": "Defensive Driving",
    "ความเร็วรถของคู่กรณี": "Defensive Driving",
    "คู่กรณีฝ่าไฟแดง": "Defensive Driving",
    "คู่กรณีตัดหน้า": "Defensive Driving",
    "คนอื่นตัดหน้า": "Defensive Driving",
    "รถพ่วงตัดหน้า": "Defensive Driving",
    "รถพ่วงเบียด": "Defensive Driving",
    "รถบัสหักหลบ": "Defensive Driving",
    "รถสวนชน": "Defensive Driving",
    "พนักงานชนคู่กรณี": "Defensive Driving",
    "ออกรถกระทันหัน": "Defensive Driving",
    # เพิ่มจากรอบตรวจ unmapped (32 รายการ)
    "กลับรถกะทันหัน": "Defensive Driving",
    "คู่กรณีชนท้าย": "Defensive Driving",
    "คู่กรณีเฉี่ยวชน": "Defensive Driving",
    "ชนคนปั่นจักรยาน": "Defensive Driving",
    "ชนท้ายคู่กรณี": "Defensive Driving",
    "ถอยชน": "Defensive Driving",
    "ถูกชน(กรณีอยู่ในเลนปกติ)": "Defensive Driving",
    "เกี่ยว/เฉี่ยว/เบียดชน": "Defensive Driving",
    "เฉี่ยวชนคนเดินเท้า/คนข้ามถนนตัดหน้ารถ": "Defensive Driving",
    "เฉี่ยวชนคู่กรณี": "Defensive Driving",
    "เลี้ยว/กลับรถ/เปลี่ยนช่องทางกะทันหัน": "Defensive Driving",
    "เลี้ยว/เปลี่ยนช่องทางกะทันหัน": "Defensive Driving",
    "เลี้ยวตัดหน้า": "Defensive Driving",
    "แซงไม่พ้น": "Defensive Driving",
    "ไม่เว้นระยะห่างให้ปลอดภัย": "Defensive Driving",

    # Focus & Attention
    "ใช้โทรศัพท์ขณะขับขี่": "Focus & Attention",
    "คุยโทรศัพท์": "Focus & Attention",
    "ผู้ขับขี่ผิดพลาดเอง(ตัดสินใจพลาด)": "Focus & Attention",
    "ผู้ขับขี่ผิดพลาดเอง": "Focus & Attention",
    "ตัดสินใจพลาด": "Focus & Attention",
    "ความพร้อมของพนักงาน(วูบ/หลับใน/อ่อนเพลีย)": "Focus & Attention",
    "ความพร้อมของพนักงาน": "Focus & Attention",
    "ง่วงนอน": "Focus & Attention",
    "ขับรถหลับใน": "Focus & Attention",
    "เมา/สารเสพติด": "Focus & Attention",
    "สะพายกระเป๋ามากกว่า 1 ใบ": "Focus & Attention",
    "การทรงตัวของพนักงานขณะขับขี่": "Focus & Attention",
    "เสียสมาธิ": "Focus & Attention",
    "ประมาท": "Focus & Attention",
    "ไม่ระวัง": "Focus & Attention",
    "เผลอ": "Focus & Attention",
    "มองไม่เห็น": "Focus & Attention",
    "ไม่ดูกระจกมองข้าง": "Focus & Attention",
    "สะดุ้งตกใจ": "Focus & Attention",
    "เดินตัดหน้า": "Focus & Attention",
    # เพิ่มจากรอบตรวจ unmapped (32 รายการ)
    "ความพร้อมของร่างกาย(วูบ/หลับใน/อ่อนเพลีย)": "Focus & Attention",
    "ตกใจหักหลบรถคันอื่น/วัสดุ/อื่นๆ": "Focus & Attention",
    "เปิดประตู(รถ)ชน/ขับชนประตู(รถ)": "Focus & Attention",
    "แสงสะท้อนเข้าตา": "Focus & Attention",
    "ไม่คุ้นชินรถ": "Focus & Attention",
    "ไม่คุ้นชินเส้นทาง": "Focus & Attention",

    # Road & Vehicle Safety
    "เสียหลักล้ม": "Road & Vehicle Safety",
    "ชนไม้กั้น/กำแพง/วัตถุอื่นๆ": "Road & Vehicle Safety",
    "ชนอุปกรณ์ก่อสร้าง": "Road & Vehicle Safety",
    "ชนเสาไฟ": "Road & Vehicle Safety",
    "ชนต้นไม้": "Road & Vehicle Safety",
    "ชนแผงกั้น": "Road & Vehicle Safety",
    "สัตว์ตัดหน้า": "Road & Vehicle Safety",
    "สัตว์ทำร้าย": "Road & Vehicle Safety",
    "สุนัขวิ่งตัดหน้า": "Road & Vehicle Safety",
    "ถนนมืด": "Road & Vehicle Safety",
    "ถนนชำรุด": "Road & Vehicle Safety",
    "ถนนลื่น": "Road & Vehicle Safety",
    "ถนนชื้น": "Road & Vehicle Safety",
    "ฝนตกถนนลื่น": "Road & Vehicle Safety",
    "น้ำท่วมถนน": "Road & Vehicle Safety",
    "สะดุดอุปกรณ์ชะลอเร็วบนถนน": "Road & Vehicle Safety",
    "พื้นลื่น(เดิน)": "Road & Vehicle Safety",
    "หลุมบ่อ": "Road & Vehicle Safety",
    "ถนนแคบ": "Road & Vehicle Safety",
    "ทางโค้ง": "Road & Vehicle Safety",
    "วัตถุสิ่งกีดขวางบนถนน": "Road & Vehicle Safety",
    "สิ่งกีดขวาง": "Road & Vehicle Safety",
    "ทรายบนถนน": "Road & Vehicle Safety",
    "ฝุ่น/ทราย": "Road & Vehicle Safety",
    "กรวดบนถนน": "Road & Vehicle Safety",
    "หินบนถนน": "Road & Vehicle Safety",
    "มีน้ำมันบนถนน": "Road & Vehicle Safety",
    "น้ำมันรั่วบนถนน": "Road & Vehicle Safety",
    "ล้อรถไม่มียาง": "Road & Vehicle Safety",
    "คันเร่งค้าง": "Road & Vehicle Safety",
    "ผ้าเบรกสึก": "Road & Vehicle Safety",
    "ความพร้อมของรถ (เบรค ล้อ ไฟ)": "Road & Vehicle Safety",
    "เบรกไม่อยู่": "Road & Vehicle Safety",
    "ยางแตก": "Road & Vehicle Safety",
    "โช้คหน้าพัง": "Road & Vehicle Safety",
    "โซ่หลุด": "Road & Vehicle Safety",
    "ล้อรถติดหลุม": "Road & Vehicle Safety",
    "สะดุดฝาท่อ": "Road & Vehicle Safety",
    "ฝาท่อระบายน้ำ": "Road & Vehicle Safety",
    "ท่อระบายน้ำ": "Road & Vehicle Safety",
    "ฝาท่อน้ำ": "Road & Vehicle Safety",
    "ของหล่นจากรถ": "Road & Vehicle Safety",
    "สายไฟพาดถนน": "Road & Vehicle Safety",
    "สายไฟรัด": "Road & Vehicle Safety",
    "ไฟฟ้าดูด": "Road & Vehicle Safety",
    "รถลื่น": "Road & Vehicle Safety",
    "ลื่นล้ม": "Road & Vehicle Safety",
    "รถไม่มีไฟท้าย": "Road & Vehicle Safety",
    "สะดุดก้อนหิน": "Road & Vehicle Safety",
    # เพิ่มจากรอบตรวจ unmapped (32 รายการ)
    "ความพร้อมของรถ (ระบุรายละเอียด)": "Road & Vehicle Safety",
    "จอดรถลงขาตั้งไม่มั่นคง": "Road & Vehicle Safety",
    "ชนทรัพย์สินอื่นๆที่ไม่ใช่รถ": "Road & Vehicle Safety",
    "ชนสัตว์": "Road & Vehicle Safety",
    "ตกร่องน้ำ/คลอง/คูน้ำ": "Road & Vehicle Safety",
    "ถนนชำรุด/ขรุขระ": "Road & Vehicle Safety",
    "ถนนลื่นทราย/ดิน/หิน/น้ำมัน/อื่นๆ": "Road & Vehicle Safety",
    "พื้นลื่น/ถนนลื่น": "Road & Vehicle Safety",
    "รถลื่นล้มไถลไปชน": "Road & Vehicle Safety",
    "สะดุดล้ม(ระบุรายละเอียด)": "Road & Vehicle Safety",
    "สัตว์วิ่งไล่ตาม": "Road & Vehicle Safety",

    # Speed Awareness
    "ขับรถเร็ว": "Speed Awareness",
    "ขับเร็ว": "Speed Awareness",
    "ขับรถเร็วเกินกำหนด": "Speed Awareness",
    "ขับรถเร็วเกินไป": "Speed Awareness",
    "เร็วเกินไป": "Speed Awareness",
    "ความเร็วสูง": "Speed Awareness",
    "ความเร็วรถเกินไป": "Speed Awareness",
    "ความเร็วรถไม่ลดลง": "Speed Awareness",
    "ความเร็วไม่ลดลง": "Speed Awareness",
    "เร่งออกตัวเร็ว": "Speed Awareness",
    "ออกตัวเร็ว": "Speed Awareness",
    "เร่งเครื่องเร็ว": "Speed Awareness",
    "เข้าโค้งเร็ว": "Speed Awareness",
    "แซงรถออกตัวเร็ว": "Speed Awareness",
    "หมุน": "Speed Awareness",
}

# ── Synonyms normalize before mapping ────────────────────────────
NORMALIZE_DICT = {
    "เลี้ยวกะทันหัน": "เลี้ยวกระทันหัน",
    "ขับเร็ว": "ขับรถเร็ว",
    "รถตัดหน้า": "ตัดหน้า",
    "เดินตัดหน้า": "คนเดินข้ามถนนตัดหน้ารถ",
    "ฝนตกถนนลื่น": "ถนนลื่น",
    "ความพร้อมของพนักงาน": "ความพร้อมของพนักงาน(วูบ/หลับใน/อ่อนเพลีย)",
}

# Fuzzy match threshold (rapidfuzz)
FUZZY_THRESHOLD = 97

# ── Risk score weights ────────────────────────────────────────────
RISK_WEIGHTS = {
    "severity_pct": 0.5,
    "road_pct":     0.3,
    "speed_pct":    0.2,
}

# Priority bins
PRIORITY_BINS   = [0, 20, 30, 100]
PRIORITY_LABELS = ["Low", "Medium", "High"]

# ── Insurance map (recommended_campaign × priority) ───────────────
AREA_INSURANCE_MAP = {
    ("Road & Vehicle Safety", "High"):   "PA เต็มรูปแบบ + ค่ารักษาพยาบาลสูง",
    ("Road & Vehicle Safety", "Medium"): "PA มาตรฐาน + ค่ารักษาพยาบาล",
    ("Road & Vehicle Safety", "Low"):    "PA พื้นฐาน",
    ("Speed Awareness",       "High"):   "PA เต็มรูปแบบ + ทุพพลภาพ",
    ("Speed Awareness",       "Medium"): "PA มาตรฐาน + ทุพพลภาพ",
    ("Speed Awareness",       "Low"):    "PA พื้นฐาน",
    ("Focus & Attention",     "High"):   "PA มาตรฐาน + อุบัติเหตุกลุ่ม",
    ("Focus & Attention",     "Medium"): "PA พื้นฐาน + อุบัติเหตุกลุ่ม",
    ("Focus & Attention",     "Low"):    "PA พื้นฐาน",
    ("Defensive Driving",     "High"):   "PA มาตรฐาน + ค่ารักษาพยาบาล",
    ("Defensive Driving",     "Medium"): "PA พื้นฐาน + ค่ารักษาพยาบาล",
    ("Defensive Driving",     "Low"):    "PA พื้นฐาน",
}

RIDER_INSURANCE_MAP = {
    ("Road & Vehicle Safety", "หยุดงานเกิน3วัน"):   "PA เต็มรูปแบบ + ค่ารักษาพยาบาลสูง",
    ("Road & Vehicle Safety", "หยุดงานไม่เกิน3วัน"): "PA มาตรฐาน + ค่ารักษาพยาบาล",
    ("Road & Vehicle Safety", "ไม่หยุดงาน"):          "PA มาตรฐาน",
    ("Speed Awareness",       "หยุดงานเกิน3วัน"):   "PA เต็มรูปแบบ + ทุพพลภาพ",
    ("Speed Awareness",       "หยุดงานไม่เกิน3วัน"): "PA มาตรฐาน + ทุพพลภาพ",
    ("Speed Awareness",       "ไม่หยุดงาน"):          "PA พื้นฐาน",
    ("Focus & Attention",     "หยุดงานเกิน3วัน"):   "PA มาตรฐาน + อุบัติเหตุกลุ่ม",
    ("Focus & Attention",     "หยุดงานไม่เกิน3วัน"): "PA พื้นฐาน + อุบัติเหตุกลุ่ม",
    ("Focus & Attention",     "ไม่หยุดงาน"):          "PA พื้นฐาน",
    ("Defensive Driving",     "หยุดงานเกิน3วัน"):   "PA มาตรฐาน + ค่ารักษาพยาบาล",
    ("Defensive Driving",     "หยุดงานไม่เกิน3วัน"): "PA พื้นฐาน",
    ("Defensive Driving",     "ไม่หยุดงาน"):          "PA พื้นฐาน",
}

# THEME_COLUMNS = {

#     "Defensive Driving": "defensive_pct",

#     "Focus & Attention": "focus_pct",

#     "Road & Vehicle Safety": "road_pct",

#     "Speed Awareness": "speed_pct",
# }


# SCORE_COLUMNS = {

#     "Defensive Driving": "defensive_score",

#     "Focus & Attention": "focus_score",

#     "Road & Vehicle Safety": "road_score",

#     "Speed Awareness": "speed_score",
# }

# ── Output paths ──────────────────────────────────────────────────
OUTPUT_DIR            = "output"
OUT_AREA_FILE         = "area_campaign_recommendation.xlsx"
OUT_RIDER_FILE        = "rider_accident_cleaned.xlsx"
OUT_JOINED_FILE       = "rider_with_personal_info.xlsx"
OUT_UNMAPPED_FILE     = "unmapped_causes.csv"

# ── Data paths ───────────────────────────────────────────────────
DATA_DIR = Path("data")
TEMPLATE_FILE = DATA_DIR / "template.xlsx"
BANK_FILE = DATA_DIR / "bank.xlsx"