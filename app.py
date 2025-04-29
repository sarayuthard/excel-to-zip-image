import streamlit as st
import pandas as pd
import requests
import zipfile
import io
import re
import concurrent.futures

st.set_page_config(page_title="📸 Excel → ZIP Image Downloader", layout="centered")

st.title("📸 แปลง URL รูปภาพจาก Excel เป็น ZIP ไฟล์")
st.caption("📌 แนบไฟล์ Excel ที่มีคอลัมน์ชื่อ 'Item' และ 'URL'")

# 🔵 แสดงรูปตัวอย่าง
st.image("example.png", caption="ตัวอย่างไฟล์ Excel ที่ต้องการ", use_container_width=True)

# 🔵 ปุ่มดาวน์โหลดไฟล์ Template
with open("Template.xlsx", "rb") as template_file:
    st.download_button(
        label="👥 ดาวน์โหลดไฟล์ตัวอย่าง (Template)",
        data=template_file,
        file_name="Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# รับไฟล์ Excel
uploaded_file = st.file_uploader("Drag and drop file here", type=["xlsx", "xls"])

# สร้างตัวแปร failed_downloads เก็บไว้ข้างนอก เพื่อไม่ให้หายไป
if 'failed_downloads' not in st.session_state:
    st.session_state.failed_downloads = []

def download_image(index_item_url):
    index, item, url = index_item_url
    if pd.isna(url):
        return None
    safe_item_name = re.sub(r'[\\/*?:"<>|]', "_", str(item).strip())

    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return (index, safe_item_name, url, response.content)
    except Exception:
        return (index, safe_item_name, url, None)
    return (index, safe_item_name, url, None)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    if 'Item' not in df.columns or 'URL' not in df.columns:
        st.error("❌ โปรดตรวจสอบว่าไฟล์มีคอลัมน์ 'Item' และ 'URL'")
    else:
        success_count = 0
        zip_buffer = io.BytesIO()

        if st.button("🚀 เริ่มดาวน์โหลดและบีบอัดรูปภาพ"):
            st.session_state.failed_downloads = []

            total = len(df)
            progress_bar = st.progress(0, text="📅 เริ่มต้นการดาวน์โหลด...")

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
                    futures = [executor.submit(download_image, (idx + 2, row['Item'], row['URL'])) for idx, row in df.iterrows()]

                    for i, future in enumerate(concurrent.futures.as_completed(futures)):
                        result = future.result()
                        if result:
                            index, filename, url, content = result
                            if content:
                                zipf.writestr(f"{filename}.jpg", content)
                                success_count += 1
                            else:
                                st.session_state.failed_downloads.append({'Row': index, 'Item': filename, 'URL': url})

                        progress_bar.progress((i + 1) / total, text=f"📷 กำลังดาวน์โหลดรูปภาพ ({i + 1}/{total})")

            zip_buffer.seek(0)

            st.success(f"✅ ดาวน์โหลดรูปเสร็จสิ้น {success_count} รายการ!")
            st.download_button(
                label="📦 ดาวน์โหลด ZIP",
                data=zip_buffer,
                file_name="downloaded_images.zip",
                mime="application/zip"
            )

        if st.session_state.failed_downloads:
            st.warning(f"🚧 Item ที่ดาวน์โหลดไม่สำเร็จ {len(st.session_state.failed_downloads)} รายการ:")
            failed_df_display = pd.DataFrame(st.session_state.failed_downloads)[['Row', 'Item', 'URL']]
            st.dataframe(failed_df_display, use_container_width=True, hide_index=True)

            failed_df_csv = pd.DataFrame(st.session_state.failed_downloads)[['Item', 'URL']]
            failed_csv = failed_df_csv.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="🔧 รายชื่อไฟล์ที่โหลดไม่สำเร็จ (CSV)",
                data=failed_csv,
                file_name="failed_downloads.csv",
                mime="text/csv"
            )
