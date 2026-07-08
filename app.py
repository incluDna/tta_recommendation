# app.py — Rider Accident Dashboard
# streamlit run app.py
# ─────────────────────────────────────────────────────────────────
import io
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from analysis import build_personalized_output, get_sheet_names, list_excel_files, run_pipeline, join_personal_data, export_outputs, natural_sort_key, get_month_order, build_area_summary, attach_cause_detail
from config import COL_AREA, COL_CAUSE, COL_SEVERITY, COL_MONTH, OUTPUT_DIR, THEMES

st.set_page_config(
    page_title="Rider Accident Dashboard",
    page_icon="🛵",
    layout="wide",
)

DATA_DIR = "data"

# ── Colors ────────────────────────────────────────────────────────
PRI_COLOR  = {"High": "#e74c3c", "Medium": "#f39c12", "Low": "#2ecc71"}
THEME_COLOR = {
    "Defensive Driving":     "#3498db",
    "Focus & Attention":     "#9b59b6",
    "Road & Vehicle Safety": "#e67e22",
    "Speed Awareness":       "#e74c3c",
}
THEME_ICON = {
    "Defensive Driving":     "🚗",
    "Focus & Attention":     "📱",
    "Road & Vehicle Safety": "🛣️",
    "Speed Awareness":       "⚡",
}
THEME_DESC = {
    "Defensive Driving":
        "อุบัติเหตุที่เกิดจากพฤติกรรมการขับขี่ไม่ระมัดระวัง เช่น ตัดหน้า ซ้อนท้าย "
        "ฝ่าไฟแดง หรือเปลี่ยนช่องทางกะทันหัน",
    "Focus & Attention":
        "อุบัติเหตุที่เกิดจากการขาดสมาธิขณะขับขี่ เช่น ใช้โทรศัพท์ ง่วงนอน "
        "หรือตัดสินใจพลาด",
    "Road & Vehicle Safety":
        "อุบัติเหตุที่เกิดจากสภาพถนนและยานพาหนะ เช่น ถนนลื่น หลุมบ่อ "
        "ยางแตก หรืออุปกรณ์ชำรุด",
    "Speed Awareness":
        "อุบัติเหตุที่เกิดจากการขับขี่ด้วยความเร็วสูงเกินไป เช่น เข้าโค้งเร็ว "
        "ออกตัวเร็ว หรือไม่ลดความเร็ว",
}

SEV_ORDER = ["หยุดงานเกิน3วัน", "หยุดงานไม่เกิน3วัน", "ไม่หยุดงาน"]

# ─────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("🛵 Rider Accident")
    st.markdown("---")

    st.subheader("📂 ข้อมูล")
    uploaded = st.file_uploader(
        "อัปโหลดไฟล์ Excel ใหม่ (.xlsx)",
        type=["xlsx"],
        key="uploader",
    )
    if uploaded is not None:
        save_path = Path(DATA_DIR) / uploaded.name
        # เตือนถ้าชื่อไฟล์ซ้ำของเดิม (จะถูก overwrite)
        Path(DATA_DIR).mkdir(parents=True, exist_ok=True)
        if save_path.exists():
            st.warning(f"⚠️ ไฟล์ '{uploaded.name}' มีอยู่แล้ว จะถูกเขียนทับ")
        with open(save_path, "wb") as f:
            f.write(uploaded.getbuffer())
        st.success(f"✅ บันทึก '{uploaded.name}' แล้ว")
        st.session_state["just_uploaded"] = uploaded.name
        st.rerun()   # rerun ให้ selectbox ด้านล่างเห็นไฟล์ใหม่ทันที

    st.markdown("---")
    data_files = list_excel_files(DATA_DIR)
    if not data_files:
        st.error(f"ไม่พบไฟล์ .xlsx ใน `{DATA_DIR}/`")
        st.stop()

    sel_file  = st.selectbox("ไฟล์", data_files)
    sheets    = get_sheet_names(DATA_DIR, sel_file)
    sel_sheet = st.selectbox("Sheet", sheets, index=min(1, len(sheets) - 1))


    run_btn = st.button("▶️ รัน Analysis", use_container_width=True, type="primary")

    st.markdown("---")
    st.subheader("🔐 Merge ข้อมูลส่วนตัว")
    use_pii   = st.checkbox("Merge ข้อมูลส่วนตัว", value=False)
    sel_bank  = sel_bank_sheet = None
    if use_pii:
        bank_files = list_excel_files(DATA_DIR)
        if bank_files:
            sel_bank = st.selectbox("ไฟล์ PII", bank_files, key="bank")
            sel_bank_sheet = st.selectbox(
                "Sheet PII", get_sheet_names(DATA_DIR, sel_bank), key="bank_sheet"
            )
            st.info("📁 บันทึกใน output/ เท่านั้น — ไม่แสดงบน dashboard")
            run_joined_btn = st.button("▶️ รัน merged excel", use_container_width=True, type="primary")

            if run_joined_btn:
                if "result" not in st.session_state:
                    st.error("⚠️ กรุณากด '▶️ รัน Analysis' ก่อน แล้วค่อยกดปุ่มนี้")
                elif "raw_with_id" not in st.session_state["result"]:
                    st.error("⚠️ กรุณากด '▶️ รัน Analysis' ใหม่อีกครั้ง (analysis.py อัปเดตแล้ว)")
                else:
                    res = st.session_state["result"]
                    personalized = build_personalized_output(res["raw_with_id"], res["df"])   # ← ตัด area_summary ออก

                    buf = io.BytesIO()
                    personalized.to_excel(buf, index=False)
                    st.success("✅ สร้างไฟล์สำเร็จ")
                    st.download_button(
                        "⬇️ ดาวน์โหลด Personalized Recommendation",
                        data=buf.getvalue(),
                        file_name=f"personalized_recommendation_{sel_file}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

    # st.markdown("---")
    # run_btn = st.button("▶️ รัน joined excel", use_container_width=True, type="primary")


# ─────────────────────────────────────────────────────────────────
# PIPELINE
# ─────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def _run(file, sheet, bank, bank_sheet):
    return run_pipeline(
        data_folder=DATA_DIR, data_filename=file, data_sheet=sheet,
        bank_folder=DATA_DIR if bank else None,
        bank_filename=bank, bank_sheet=bank_sheet or 0,
        output_dir=OUTPUT_DIR,
    )

if "result" not in st.session_state or run_btn:
    with st.spinner("⏳ กำลังประมวลผล..."):
        try:
            st.session_state["result"] = _run(
                sel_file, sel_sheet,
                sel_bank if use_pii else None,
                sel_bank_sheet if use_pii else None,
            )
        except ValueError as e:
            st.error(f"❌ ไม่พบ column ที่จำเป็น:\n\n```\n{e}\n```")
            st.stop()
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")
            st.stop()

res          = st.session_state["result"]
df_full      = res["df"]
area         = res["area_summary"]
unmapped     = res["unmapped"]

# ─────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────
st.title("🛵 Rider Accident — TTA Campaign Recommendation Dashboard")
st.caption(
    f"ไฟล์: **{sel_file}** | Sheet: **{sel_sheet}**")
# ─────────────────────────────────────────────────────────────────
# FILTER: เลือกเดือน / พื้นที่ — กรองทั้งหน้าตามที่เลือก (อยู่ติดแถบ tab)
# ─────────────────────────────────────────────────────────────────
f1, f2 = st.columns(2)
 
with f1:
    if COL_MONTH in df_full.columns:
        month_opts = get_month_order(df_full[COL_MONTH])
        sel_months = st.multiselect(
            "📅 เลือกเดือน", options=month_opts, default=month_opts,
            placeholder="เลือกเดือน (ค่าเริ่มต้น: ทั้งหมด)",
        )
    else:
        sel_months = None
 
with f2:
    if COL_AREA in df_full.columns:
        area_opts = sorted(df_full[COL_AREA].dropna().unique().tolist(), key=natural_sort_key)
        sel_areas = st.multiselect(
            "📍 เลือกพื้นที่", options=area_opts, default=area_opts,
            placeholder="เลือกพื้นที่ (ค่าเริ่มต้น: ทั้งหมด)",
        )
    else:
        sel_areas = None
 
# กรอง df ตามที่เลือก (ถ้าไม่เลือกอะไรเลย = โชว์ทั้งหมด กันหน้าว่างเปล่า)
df = df_full.copy()
if sel_months is not None and len(sel_months) > 0:
    df = df[df[COL_MONTH].isin(sel_months)]
if sel_areas is not None and len(sel_areas) > 0:
    df = df[df[COL_AREA].isin(sel_areas)]
 
if df.empty:
    st.warning("⚠️ ไม่มีข้อมูลตรงกับตัวกรองที่เลือก ลองเลือกเดือน/พื้นที่เพิ่ม")
    st.stop()
 
# recompute area summary ใหม่จากข้อมูลที่กรองแล้ว (เดิมคำนวณจากข้อมูลทั้งหมด)
area = build_area_summary(df)
area = attach_cause_detail(area, df, top_n=3)
 
st.caption(f"Data: **{len(df):,}** / {len(df_full):,} | พื้นที่: **{len(area)}** แห่ง")
 
if not unmapped.empty:
    with st.expander(f"⚠️ พบ {len(unmapped)} สาเหตุที่ยัง map ไม่ได้"):
        st.dataframe(unmapped, use_container_width=True, hide_index=True)
else:
    st.markdown("✅ ทุก cause mapped ครบแล้ว")

# st.markdown("---")

# ─────────────────────────────────────────────────────────────────
# 4 TABS
# ─────────────────────────────────────────────────────────────────

st.markdown("""
<style>
/* container ของแถบ tab ทั้งหมด */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    # background-color: #f5f5f5;
    padding: 8px;
    border-radius: 12px;
}

/* แต่ละ tab */
.stTabs [data-baseweb="tab"] {
    height: 55px;
    # white-space: pre-wrap;
    # background-color: #ffffff;
    border-radius: 10px;
    padding: 10px 24px;
    font-size: 18px;
    font-weight: 600;
    # color: #444;
    border: 1px solid #9c9a9a;
    transition: all 0.2s ease-in-out;
}

/* hover ให้รู้สึกว่ากดได้ */
.stTabs [data-baseweb="tab"]:hover {
    # background-color: #e8ebf5;
    # color: #1a1a1a;
    transform: translateY(-2px);
    cursor: pointer;
}

/* tab ที่ถูกเลือกอยู่ ให้เด่นสุด */
.stTabs [aria-selected="true"] {
    background-color: #4C6EF5 !important;
    color: white !important;
    box-shadow: 0 4px 10px rgba(76, 110, 245, 0.4);
}

/* ตัวอักษรข้างในให้ใหญ่ขึ้นด้วย */
.stTabs [data-baseweb="tab"] p {
    font-size: 18px;
}
</style>
""", unsafe_allow_html=True)

tab_dash, tab_insight, tab_campaign, tab_insurance = st.tabs([
    "📊 dashboard", "🔍 insight", "🎯 campaign", "🛡️ insurance",
])
# tab_dash, tab_insight, tab_campaign, tab_insurance = st.tabs([
#     "📊 dashboard", "🔍 insight", "🎯 campaign", "🛡️ insurance",
# ])


# ══════════════════════════════════════════════════════════════════
# TAB 1 — DASHBOARD
# ══════════════════════════════════════════════════════════════════
with tab_dash:

    # KPI row
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("จำนวนรายงาน", f"{len(df):,}")
    k2.metric("High Risk Areas", int((area["priority_level"] == "High").sum()))
    k3.metric("Risk Score เฉลี่ย", f"{area['risk_score'].mean():.1f}")
    k4.metric("Mapped %", f"{df['campaign_theme'].notna().mean()*100:.1f}%")
    top_theme = df["campaign_theme"].value_counts().idxmax() if df["campaign_theme"].notna().any() else "-"
    k5.metric("Theme หลัก", top_theme)

    st.markdown("---")

    # ── Section 1: Rider overview + Theme definition + Risk formula ──
    st.markdown("#### ข้อมูล Rider (ระดับบุคคล — ไม่มี PII)")
    r1c1, r1c2, r1c3 = st.columns(3)

    with r1c1:
        if "age_group" in df.columns:
            ag = df["age_group"].value_counts().reset_index()
            ag.columns = ["กลุ่ม", "จำนวน"]
            st.plotly_chart(
                px.bar(ag, x="กลุ่ม", y="จำนวน", title="กลุ่มอายุ",
                       color_discrete_sequence=["#3498db"]),
                use_container_width=True, config={"displayModeBar": False}
            )

    with r1c2:
        if "rider_exp_group" in df.columns:
            eg = df["rider_exp_group"].value_counts().reset_index()
            eg.columns = ["ประสบการณ์", "จำนวน"]
            st.plotly_chart(
                px.bar(eg, x="ประสบการณ์", y="จำนวน", title="ประสบการณ์ Rider",
                       color_discrete_sequence=["#9b59b6"]),
                use_container_width=True, config={"displayModeBar": False}
            )

    with r1c3:
        if "period" in df.columns:
            pg = df["period"].value_counts().reset_index()
            pg.columns = ["ช่วงเวลา", "จำนวน"]
            st.plotly_chart(
                px.bar(pg, x="ช่วงเวลา", y="จำนวน", title="ช่วงเวลาเกิดเหตุ",
                       color_discrete_sequence=["#e67e22"]),
                use_container_width=True, config={"displayModeBar": False}
            )

    # ── เดือนที่เกิดเหตุ (เรียงตามปฏิทิน ไม่ใช่ตัวอักษร) ──────────────
    if COL_MONTH in df.columns:
        month_order = get_month_order(df[COL_MONTH])
        mg = df[COL_MONTH].value_counts().reindex(month_order).reset_index()
        mg.columns = ["เดือน", "จำนวน"]
        fig_month = px.bar(
            mg, x="เดือน", y="จำนวน", title="อุบัติเหตุตามเดือน",
            color_discrete_sequence=["#16a085"],
            category_orders={"เดือน": month_order},   # ← เรียง Jan → Dec ตามปฏิทิน
        )
        fig_month.update_layout(height=340, margin=dict(t=40, b=10))
        st.plotly_chart(fig_month, use_container_width=True, config={"displayModeBar": False})
    
    st.markdown("---")

    # ── Section 2: Theme card + Risk Score definition ──────────────
    st.markdown("#### นิยาม Theme & Risk Score")
    def_col, risk_col = st.columns([3, 2], gap="large")

    with def_col:
        st.markdown("**Campaign Themes**")
        for theme in THEMES:
            color = THEME_COLOR[theme]
            icon  = THEME_ICON[theme]
            desc  = THEME_DESC[theme]
            st.markdown(
                f"""<div style="border-left:4px solid {color};padding:8px 12px;
                margin-bottom:8px;border-radius:4px;background:rgba(0,0,0,0.05)">
                <b>{icon} {theme}</b><br>
                <span style="font-size:0.85rem;color:#aaa">{desc}</span>
                </div>""",
                unsafe_allow_html=True,
            )

    with risk_col:
        st.markdown("**Risk Score คำนวณจาก**")
        st.markdown("""
```
Risk Score =
  severity_pct  × 0.5
+ road_pct      × 0.3
+ speed_pct     × 0.2
```
        """)
        st.markdown("**Priority Level**")
        for label, color, rng in [
            ("High",   "#e74c3c", "> 30"),
            ("Medium", "#f39c12", "20 – 30"),
            ("Low",    "#2ecc71", "< 20"),
        ]:
            st.markdown(
                f'<span style="background:{color};color:#fff;padding:2px 10px;'
                f'border-radius:10px;font-size:0.8rem">{label}</span> '
                f'<span style="font-size:0.85rem"> Risk Score {rng}</span>',
                unsafe_allow_html=True,
            )
            st.write("")

    st.markdown("---")

    # ── Section 3: Charts ─────────────────────────────────────────
    ch1, ch2 = st.columns(2)

    with ch1:
        # Stacked bar: theme % per area
        pct_cols = {
            "defensive_pct": "Defensive Driving",
            "focus_pct":     "Focus & Attention",
            "road_pct":      "Road & Vehicle Safety",
            "speed_pct":     "Speed Awareness",
        }
        avail = [c for c in pct_cols if c in area.columns]
        if avail:
            melt = (
                area[avail]
                .rename(columns=pct_cols)
                .reset_index()
                .melt(id_vars=COL_AREA, var_name="Theme", value_name="%")
            )
            fig = px.bar(
                melt, x=COL_AREA, y="%", color="Theme",
                color_discrete_map=THEME_COLOR,
                barmode="stack",
                title="สัดส่วน Theme ต่อพื้นที่ (%)",
            )
            fig.update_layout(height=360, margin=dict(t=40, b=10))
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with ch2:
        # Bar: risk score per area colored by priority
        fig2 = px.bar(
            area.reset_index(),
            x=COL_AREA, y="risk_score",
            color="priority_level",
            color_discrete_map=PRI_COLOR,
            text="risk_score",
            title="Risk Score ต่อพื้นที่",
        )
        fig2.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        fig2.update_layout(height=360, margin=dict(t=40, b=10), showlegend=True)
        st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

    # Top 20 causes
    st.markdown("#### 🔝 Top 20 สาเหตุ")
    if COL_CAUSE in df.columns:
        top = df[COL_CAUSE].value_counts().head(20).reset_index()
        top.columns = ["สาเหตุ", "จำนวน"]
        fig3 = px.bar(
            top, x="จำนวน", y="สาเหตุ", orientation="h",
            color="จำนวน", color_continuous_scale="Blues",
        )
        fig3.update_layout(
            yaxis={"categoryorder": "total ascending"},
            height=520, margin=dict(t=20, b=10),
            coloraxis_showscale=False,
        )
        st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})


# ══════════════════════════════════════════════════════════════════
# TAB 2 — INSIGHT
# ══════════════════════════════════════════════════════════════════
with tab_insight:

    st.header("📊 Overall Theme Distribution")
    st.caption(
        "ภาพรวมของสัดส่วนอุบัติเหตุในแต่ละ Theme และความสัมพันธ์กับระดับความรุนแรง"
    )
    # pie_col, heat_col = st.columns(2)

    with st.container():
        st.markdown("**Global Theme Share**")
        theme_dist = df["campaign_theme"].value_counts().reset_index()
        theme_dist.columns = ["theme", "count"]
        fig_pie = px.pie(
            theme_dist, names="theme", values="count",
            color="theme", color_discrete_map=THEME_COLOR,
            hole=0.35,
        )
        fig_pie.update_traces(textposition="inside", textinfo="percent+label")
        fig_pie.update_layout(showlegend=False, height=340, margin=dict(t=10, b=10))
        st.plotly_chart(fig_pie, use_container_width=True, config={"displayModeBar": False})

    # with heat_col:
    #     st.markdown("**Theme × Severity**")
    #     if COL_SEVERITY in df.columns and "campaign_theme" in df.columns:
    #         heat = pd.crosstab(df["campaign_theme"], df[COL_SEVERITY])
    #         # reorder severity columns
    #         sev_order = [c for c in SEV_ORDER if c in heat.columns]
    #         heat = heat[sev_order]
    #         fig_heat = px.imshow(
    #             heat, text_auto=True,
    #             color_continuous_scale="Reds",
    #             aspect="auto",
    #         )
    #         fig_heat.update_layout(height=340, margin=dict(t=10, b=10))
    #         st.plotly_chart(fig_heat, use_container_width=True, config={"displayModeBar": False})

    # ==========================================
    # Rider Experience Insight
    # ==========================================

    exp_pct = (
        pd.crosstab(
            df["rider_exp_group"],
            df[COL_SEVERITY],
            normalize="index",
        ) * 100
    ).round(1)

    highest_group = exp_pct["หยุดงานเกิน3วัน"].idxmax()
    highest_value = exp_pct["หยุดงานเกิน3วัน"].max()

    
    st.markdown("---")
    # st.markdown("👤 Rider Experience Analysis")
    st.info(f"""
    ### 👤 Rider Experience Insight

    พบว่า Rider ที่มีประสบการณ์ **{highest_group}**
    มีสัดส่วนการหยุดงานเกิน 3 วันสูงที่สุด
    คิดเป็น **{highest_value:.1f}%**

    Recommendation

    ควรศึกษาปัจจัยที่ทำให้ Rider กลุ่มนี้เกิดอุบัติเหตุรุนแรง
    และออกแบบกิจกรรมหรือการอบรมให้เหมาะกับกลุ่มดังกล่าว
    """)

    if "rider_exp_group" in df.columns and COL_SEVERITY in df.columns:
        ct = (
            pd.crosstab(df["rider_exp_group"], df[COL_SEVERITY],
        normalize="index",)* 100
        ).round(1)
        sev_order = [c for c in SEV_ORDER if c in ct.columns]
        ct = ct[sev_order]

        # color cells by value
        styled = (
            ct.style
            .background_gradient(cmap="Reds", subset=["หยุดงานเกิน3วัน"])
            .background_gradient(cmap="Oranges", subset=["หยุดงานไม่เกิน3วัน"])
            .background_gradient(cmap="Greens", subset=["ไม่หยุดงาน"])
            .format("{:,.0f}%")
        )
        st.dataframe(styled, use_container_width=True)

    st.markdown("---")
    # st.markdown("🚦 Campaign Theme Analysis")
    # st.caption(
    # "ตารางแสดงสัดส่วนความรุนแรงของอุบัติเหตุในแต่ละ Campaign Theme "
    # "เพื่อระบุ Theme ที่ควรได้รับการรณรงค์ก่อน"
    # )
    if COL_SEVERITY in df.columns and "campaign_theme" in df.columns:
        ct2 = (
            pd.crosstab(df["campaign_theme"], df[COL_SEVERITY], normalize="index") * 100
        ).round(1)
        top_theme = ct2["หยุดงานเกิน3วัน"].idxmax()
        top_value = ct2["หยุดงานเกิน3วัน"].max()
        st.info(
            f"""
        ### 🚦 Campaign Theme Insight

        - Theme ที่มีอัตราการหยุดงานเกิน 3 วันสูงที่สุดคือ

        ## {top_theme}

        คิดเป็น **{top_value:.1f}%**

        💡 Recommendation

        ควรจัดลำดับการรณรงค์ด้าน **{top_theme}**
        เป็นอันดับแรก
        """
        )
        sev_order = [c for c in SEV_ORDER if c in ct2.columns]
        ct2 = ct2[sev_order]
        styled2 = (
            ct2.style
            .background_gradient(cmap="Reds", subset=["หยุดงานเกิน3วัน"])
            .background_gradient(cmap="Oranges", subset=["หยุดงานไม่เกิน3วัน"])
            .background_gradient(cmap="Greens", subset=["ไม่หยุดงาน"])
            .format("{:.1f}%")
        )
        st.dataframe(styled2, use_container_width=True)


# ══════════════════════════════════════════════════════════════════
# TAB 3 — CAMPAIGN
# ══════════════════════════════════════════════════════════════════
with tab_campaign:

    if area.empty:
        st.info("ยังไม่มีข้อมูล area summary")
    else:
        for area_name, row in area.iterrows():
            pri   = str(row["priority_level"])
            color = PRI_COLOR.get(pri, "#888")
            rec   = row.get("recommended_campaign", "-")
            sup   = row.get("supporting_campaigns", "-")
            risk  = row["risk_score"]
            conf  = row.get("confidence", float("nan"))
            ins   = row.get("insurance_recommendation", "-")
            theme = row.get("dominant_theme", "-")

            # card
            st.markdown(
                f"""<div style="border:1px solid #333;border-radius:8px;
                padding:16px 20px;margin-bottom:12px;
                border-left:5px solid {color}">
                <div style="display:flex;justify-content:space-between;align-items:center">
                  <b style="font-size:1.05rem">📍 {area_name}</b>
                  <span style="background:{color};color:#fff;padding:2px 12px;
                  border-radius:12px;font-size:0.8rem">{pri}</span>
                </div>
                </div>""",
                unsafe_allow_html=True,
            )

            c1, c2, c3 = st.columns(3)
            c1.metric("Risk Score", f"{risk:.1f}")
            c2.metric("Confidence", f"{conf:.1f}%" if not np.isnan(conf) else "-")
            c3.metric("Recommended (อิงจากความถี่และความรุนแรง)", f"{THEME_ICON.get(rec,'')} {rec}")

            rc1, rc2 = st.columns(2)
            rc1.markdown(f"""
<div style="font-size:16px; color:gray;">➕ Supporting</div>
<div style="font-size:24px; font-weight:500; margin-top:2px;">{sup}</div>
""", unsafe_allow_html=True)
            rc2.markdown(f"""
<div style="font-size:16px; color:gray;">▶ Dominant Theme (อิงจากความถี่)</div>
<div style="font-size:24px; font-weight:500; margin-top:2px;">{theme}</div>
""", unsafe_allow_html=True)
        
            if "reason" in row:
                st.caption(row["reason"])

            # theme breakdown bar (mini)
            pct_data = {
                v: row.get(k, 0)
                for k, v in {
                    "defensive_pct": "Defensive Driving",
                    "focus_pct":     "Focus & Attention",
                    "road_pct":      "Road & Vehicle Safety",
                    "speed_pct":     "Speed Awareness",
                }.items()
            }
            mini_df = pd.DataFrame({
                "Theme": list(pct_data.keys()),
                "%": list(pct_data.values()),
            })
            fig_mini = px.bar(
                mini_df, x="Theme", y="%",
                color="Theme", color_discrete_map=THEME_COLOR,
                height=180,
            )
            fig_mini.update_layout(
                showlegend=False,
                margin=dict(t=4, b=4, l=4, r=4),
                xaxis_title=None, yaxis_title=None,
            )
            st.plotly_chart(fig_mini, use_container_width=True, config={"displayModeBar": False})
            st.markdown("---")

        # Full table
        st.markdown("#### ตารางสรุปทุกพื้นที่")
        show_cols = [
            "risk_score", "priority_level",
            "recommended_campaign", "recommended_campaign_detail", "supporting_campaigns",
            "confidence", "insurance_recommendation",
        ]
        tbl = area[[c for c in show_cols if c in area.columns]].reset_index()
        tbl["priority_level"] = tbl["priority_level"].astype(str)
        tbl = tbl.sort_values(COL_AREA, key=lambda s: s.map(natural_sort_key))  # default: เรียงตามพื้นที่แบบ natural sort
        st.caption("💡 คลิกหัวตาราง (เช่น 'พื้นที่' หรือ 'risk_score') เพื่อเรียงลำดับได้")
        st.dataframe(
            tbl.style.background_gradient(subset=["risk_score"], cmap="RdYlGn_r"),
            use_container_width=True,
            hide_index=True,
        )

        # Download
        buf = io.BytesIO()
        area.to_excel(buf, index=True)
        st.download_button(
            "⬇️ ดาวน์โหลด Area Summary",
            data=buf.getvalue(),
            file_name="area_campaign_recommendation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("---")
        st.markdown("#### สาเหตุที่พบได้ในแต่ละ Theme")
        with st.expander("📖 ดูรายละเอียด: สาเหตุในแต่ละ theme"):
            st.caption("ใช้ตอนคิด message campaign — เห็นคำจริงที่ระบบจัดกลุ่มไว้ในแต่ละ theme")

            from config import CAUSE_THEME

            mapping_df = pd.DataFrame(
                [{"สาเหตุ (cause)": cause, "Theme": theme} for cause, theme in CAUSE_THEME.items()]
            ).sort_values(["Theme", "สาเหตุ (cause)"]).reset_index(drop=True)

            # เลือกดูเฉพาะ theme ที่สนใจ ไม่ต้อง scroll ทั้งหมด
            selected_theme = st.selectbox(
                "เลือก Theme ที่จะดู", ["ทั้งหมด"] + sorted(mapping_df["Theme"].unique().tolist())
            )
            if selected_theme != "ทั้งหมด":
                mapping_df = mapping_df[mapping_df["Theme"] == selected_theme]

            st.dataframe(mapping_df, use_container_width=True, hide_index=True)
            st.caption(f"รวม {len(mapping_df)} รายการ")
            # ปุ่มดาวน์โหลดฉบับเต็ม (ไม่กรอง)
        full_mapping_df = pd.DataFrame(
            [{"สาเหตุ (cause)": cause, "Theme": theme} for cause, theme in CAUSE_THEME.items()]
        ).sort_values(["Theme", "สาเหตุ (cause)"])

        buf_mapping = io.BytesIO()
        full_mapping_df.to_excel(buf_mapping, index=False)
        st.download_button(
            "⬇️ ดาวน์โหลดตาราง Cause → Theme ฉบับเต็ม",
            data=buf_mapping.getvalue(),
            file_name="cause_theme_mapping.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ══════════════════════════════════════════════════════════════════
# TAB 4 — INSURANCE
# ══════════════════════════════════════════════════════════════════
with tab_insurance:
    col_title, col_info = st.columns([10, 1])
    with col_title:
        st.markdown("##### สัดส่วนประเภทประกันที่แนะนำ (Rider Level)")
    with col_info:
        with st.popover("ℹ️"):
            st.markdown("""
            **🟢 PA พื้นฐาน (Basic Personal Accident)**
            เหมาะสำหรับพื้นที่ที่มีความเสี่ยงต่ำ
            - คุ้มครองอุบัติเหตุพื้นฐาน
            - ค่ารักษาพยาบาลจากอุบัติเหตุ
            - เงินชดเชยกรณีเสียชีวิตหรือสูญเสียอวัยวะจากอุบัติเหตุ
            - เหมาะสำหรับพื้นที่ที่มีโอกาสเกิดอุบัติเหตุรุนแรงต่ำ

            ---

            **🟡 PA มาตรฐาน (Standard Personal Accident)**
            เหมาะสำหรับพื้นที่ที่มีความเสี่ยงปานกลาง
            - ครอบคลุมสิทธิของ PA พื้นฐาน
            - เพิ่มวงเงินค่ารักษาพยาบาล
            - เพิ่มความคุ้มครองกรณีบาดเจ็บที่ต้องพักรักษาตัว
            - เหมาะสำหรับพื้นที่ที่มีความเสี่ยงปานกลางและอุบัติเหตุมีผลกระทบต่อการทำงาน

            ---

            **🔴 PA เต็มรูปแบบ (Comprehensive Personal Accident)**
            เหมาะสำหรับพื้นที่ที่มีความเสี่ยงสูง
            - ความคุ้มครองอุบัติเหตุแบบครอบคลุม
            - วงเงินค่ารักษาพยาบาลสูง
            - คุ้มครองกรณีเสียชีวิต สูญเสียอวัยวะ และทุพพลภาพจากอุบัติเหตุ
            - เหมาะสำหรับพื้นที่ที่มีโอกาสเกิดอุบัติเหตุรุนแรงหรือบาดเจ็บหนัก

            ---

            **🏥 ความคุ้มครองเพิ่มเติม**

            💙 **ค่ารักษาพยาบาลสูง**
            - เพิ่มวงเงินค่ารักษาพยาบาลจากอุบัติเหตุ
            - เหมาะสำหรับพื้นที่ที่มีโอกาสเกิดการบาดเจ็บรุนแรง

            🧑‍🦽 **ทุพพลภาพ**
            - คุ้มครองกรณีสูญเสียความสามารถในการทำงานจากอุบัติเหตุ
            - ช่วยลดผลกระทบด้านรายได้ในระยะยาว

            👥 **อุบัติเหตุกลุ่ม**
            - ความคุ้มครองอุบัติเหตุสำหรับพนักงานในรูปแบบกลุ่ม
            - เหมาะสำหรับองค์กรที่ต้องการดูแลพนักงานหลายคนพร้อมกัน
                    """)
    if "insurance_recommendation" in df.columns:
        ins_dist = df["insurance_recommendation"].value_counts().reset_index()
        ins_dist.columns = ["ประเภทประกัน", "จำนวน"]

        d1, d2 = st.columns([2, 3])
        with d1:
            fig_donut = px.pie(
                ins_dist, names="ประเภทประกัน", values="จำนวน",
                hole=0.5,
                color_discrete_sequence=px.colors.qualitative.Set2,
            )
            fig_donut.update_traces(textposition="inside", textinfo="percent", hovertemplate="<b>%{label}</b><br>จำนวน: %{value:,}<br>สัดส่วน: %{percent}<extra></extra>")
            fig_donut.update_layout(height=420, margin=dict(t=10, b=120), legend=dict(
            orientation="h",        # legend แนวนอน
            yanchor="top",
            y=-0.15,                # ดันลงไปใต้ chart
            xanchor="center",
            x=0.5,
            font=dict(size=11),
        ),
        showlegend=True)
            st.plotly_chart(fig_donut, use_container_width=True, config={"displayModeBar": False})

        with d2:
            st.dataframe(ins_dist, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("#### Insurance Recommendation ต่อพื้นที่")

    if not area.empty:
        ins_cols = [
            "risk_score", "priority_level",
            "recommended_campaign", "insurance_recommendation",
        ]
        ins_tbl = area[[c for c in ins_cols if c in area.columns]].reset_index()
        ins_tbl["priority_level"] = ins_tbl["priority_level"].astype(str)
        ins_tbl = ins_tbl.sort_values(COL_AREA, key=lambda s: s.map(natural_sort_key))

        def _color_pri(val):
            return f"background-color:{PRI_COLOR.get(val,'#555')};color:white;border-radius:4px"

        st.caption("💡 คลิกหัวตารางเพื่อเรียงลำดับได้")
        st.dataframe(
            ins_tbl.style
            .map(_color_pri, subset=["priority_level"])
            .background_gradient(subset=["risk_score"], cmap="RdYlGn_r"),
            use_container_width=True,
            hide_index=True,
        )

    if use_pii and "joined_pii" in res.get("saved_files", {}):
        st.info(f"📁 ไฟล์ PII join: `{res['saved_files']['joined_pii']}`")
        st.warning("🔒 ไฟล์ PII ไม่สามารถดาวน์โหลดผ่าน dashboard ได้")