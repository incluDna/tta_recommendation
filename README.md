## TTA Recommendation🛵🗺️
> ระบบในการประเมินความเสี่ยงของพื้นที่ต่าง ๆ ในรูปแบบของ risk score เพื่อทำ Campaign & Insurance recommendation (แสดงผ่าน streamlit dashboard)

https://tta-recommendation.streamlit.app/

#### Upload file
- upload file ใน `/data` และ เลือก sheet ที่มี data อยู่
  or
- upload ผ่านหน้า [web](https://tta-recommendation.streamlit.app/)

#### Run command
```
streamlit run app.py
```
web จะแสดง
- Dashboard
- Insight
- Campaign
- Insurance

#### Output
นอกจากข้อมูล risk score และ insurance recommendation ยังมี ***personalize recommendation*** (รายบุคคล) อยู่ใน `/output` ด้วย