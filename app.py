import streamlit as st
import pandas as pd
import requests
import zipfile
import io
import re

st.set_page_config(page_title="📸 Excel → ZIP Image Downloader", layout="centered")

st.title("📸 แปลง URL รูปภาพจาก Excel เป็น ZIP ไฟล์")
st.caption("📎 แนบไฟล์ Excel ที่มีคอลัมน์ชื่อ 'Item' และ 'URL'")

# 🔵 แสดงรูปตัวอย่าง
st.image("example.png", caption="ตัวอย่างไฟล์ Excel ที่ต้องการ", use_container_width=True)

# 🔵 ปุ่มดาวน์โหลดไฟล์ Template
with open("Template.xlsx", "rb") as template_file:
    st.download_button(
        label="📥 ดาวน์โหลดไฟล์ตัวอย่าง (Template)",
        data=template_file,
        file_name="Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# รับไฟล์ Excel
uploaded_file = st.file_uploader("Drag and drop file here", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    if 'Item' not in df.columns or 'URL' not in df.columns:
        st.error("❌ โปรดตรวจสอบว่าไฟล์มีคอลัมน์ 'Item' และ 'URL'")
    else:
        if st.button("🚀 เริ่มดาวน์โหลดและบีบอัดรูปภาพ"):
            zip_buffer = io.BytesIO()

            total = len(df)
            progress_bar = st.progress(0, text="📥 เริ่มต้นการดาวน์โหลด...")

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                success_count = 0
                for i, row in df.iterrows():
                    item = str(row['Item']).strip()
                    url = row['URL']

                    if pd.isna(url):
                        continue

                    safe_item_name = re.sub(r'[\\/*?:"<>|]', "_", item)

                    try:
                        response = requests.get(url, timeout=10)
                        if response.status_code == 200:
                            zipf.writestr(f"{safe_item_name}.jpg", response.content)
                            success_count += 1
                        else:
                            st.warning(f"⚠️ ไม่สามารถโหลดรูป {item} (status: {response.status_code})")
                    except Exception as e:
                        st.warning(f"⚠️ เกิดข้อผิดพลาดกับ {item}: {e}")

                    progress_bar.progress((i + 1) / total, text=f"📷 กำลังดาวน์โหลดรูปภาพ ({i + 1}/{total})")

            zip_buffer.seek(0)

            st.success(f"✅ ดาวน์โหลดรูปเสร็จสิ้น {success_count} รายการ!")
            st.download_button(
                label="📦 ดาวน์โหลด ZIP",
                data=zip_buffer,
                file_name="downloaded_images.zip",
                mime="application/zip"
            )
