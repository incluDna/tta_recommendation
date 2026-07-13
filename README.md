# TTA Recommendation🛵🗺️
> ระบบในการประเมินความเสี่ยงของพื้นที่ต่าง ๆ ในรูปแบบของ risk score เพื่อทำ Campaign & Insurance recommendation (แสดงผ่าน streamlit dashboard)

https://tta-recommendation.streamlit.app/

### [คู่มือการใช้งาน](https://github.com/incluDna/tta_recommendation/blob/7995db961d02b4aaded3e841acb42202d6a694a3/assets/tta_handbook.pdf)

### Upload file
- upload file ใน `/data` และ เลือก sheet ที่มี data อยู่ (กรณี run localhost)
  <br>or
- upload ผ่านหน้า [web](https://tta-recommendation.streamlit.app/)

### Library & Run command (กรณี run localhost)
```
pip install -r requirements.txt

streamlit run app.py
```
---

### web จะแสดง
- **Dashboard**: วิเคราะห์ข้อมูลทั่วไปของ rider, นิยาม Theme & Risk Score, สัดส่วน Theme ต่อพื้นที่(%), Risk score ต่อพื้นที่, Top 20 สาเหตุ

| Terms | Definitions |
|-------------------|-------------|
| Campaign Themes | - 🚗 Defensive Driving: อุบัติเหตุที่เกิดจากพฤติกรรมการขับขี่ไม่ระมัดระวัง เช่น ตัดหน้า ซ้อนท้าย ฝ่าไฟแดง หรือเปลี่ยนช่องทางกะทันหัน <br> - 📱 Focus & Attention: อุบัติเหตุที่เกิดจากการขาดสมาธิขณะขับขี่ เช่น ใช้โทรศัพท์ ง่วงนอน หรือตัดสินใจพลาด <br> - 🛣️ Road & Vehicle Safety: อุบัติเหตุที่เกิดจากสภาพถนนและยานพาหนะ เช่น ถนนลื่น หลุมบ่อ ยางแตก หรืออุปกรณ์ชำรุด <br> - ⚡ Speed Awareness: อุบัติเหตุที่เกิดจากการขับขี่ด้วยความเร็วสูงเกินไป เช่น เข้าโค้งเร็ว ออกตัวเร็ว หรือไม่ลดความเร็ว |
| Risk Score | ```Risk Score = (severity_pct × 0.5) + (road_pct × 0.3) + (speed_pct × 0.2)``` |
| Priority Level | 🔴High Risk Score > 30 <br> 🟡Medium Risk Score 20 – 30 <br> 🟢Low Risk Score < 20 |

- **Insight**: Overall Theme (pie-chart), Rider Experience Insight, Campaign Theme Insight
- **Campaign**: Recommended (อิงจากความถี่และความรุนแรง), Dominant Theme (อิงจากความถี่), Supporting theme, Risk Score, Confidence แต่ละพื้นที่
- **Insurance**: Insurance Recommendation ต่อพื้นที่

---

### Output
นอกจากข้อมูล risk score และ insurance recommendation ยังมี ***personalize recommendation*** (รายบุคคล) อยู่ใน `/output` ด้วย
