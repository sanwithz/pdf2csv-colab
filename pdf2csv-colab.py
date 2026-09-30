# 1. ติดตั้ง Library ที่จำเป็นทั้งหมด
!pip install -q pdfplumber pandas openpyxl gdown google-genai

import gdown
import pdfplumber
import pandas as pd
import numpy as np
import json
import time
from google import genai
from google.genai import types

# ==========================================
# 2. ใส่ GEMINI API KEY ของคุณตรงนี้
# ==========================================
GEMINI_API_KEY = "XXXXXXXXX"
client = genai.Client(api_key=GEMINI_API_KEY)

# ==========================================
# 3. ดาวน์โหลดไฟล์จาก Google Drive
# ==========================================
file_id = "1dX1PxLxrQCDLHthisB-xxxxx"
pdf_path = "downloaded_file.pdf"
gdown.download(f"https://drive.google.com/uc?id={file_id}", pdf_path, quiet=False)

# ==========================================
# 4. สกัดข้อมูลตารางจาก PDF ด้วย pdfplumber
# ==========================================
all_rows = []
with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        table_settings = {
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
            "snap_tolerance": 3,
            "join_tolerance": 3,
        }
        tables = page.extract_tables(table_settings)
        for table in tables:
            for row in table:
                cleaned_row = [cell.strip() if cell else None for cell in row]
                # กรองแถวว่างทิ้งตั้งแต่แรก
                if any(cleaned_row):
                    all_rows.append(cleaned_row)

df_raw = pd.DataFrame(all_rows)

# กำหนด Header เบื้องต้น
header = df_raw.iloc[0]
df = df_raw[1:].copy()
df.columns = header
df.replace(r'^\s*$', np.nan, regex=True, inplace=True)
df.dropna(how='all', inplace=True)

print(f"สกัดตารางดิบได้ทั้งหมด: {len(df)} แถว")

# ==========================================
# 5. ใช้ Gemini (รุ่น Flash ประหยัดสุด) คลีนและจัดระเบียบตาราง
# ==========================================
system_instruction = """
คุณคือผู้เชี่ยวชาญด้าน Data Cleaning และ Normalization ตารางข้อมูลภาษาไทย
หน้าที่ของคุณ:
1. แตกข้อมูลในตารางที่ถูกรวมหลายบรรทัดด้วย \n ให้ออกมาเป็นแถวเดี่ยวที่ถูกต้องสมบูรณ์ (1 แถว = 1 วิชา)
2. แก้ไขสระและวรรณยุกต์ภาษาไทยที่เพี้ยน/ลอย ให้ถูกต้องตามพจนานุกรม เช่น "กลุมสาระ" -> "กลุ่มสาระ", "หอง" -> "ห้อง", "คณิตศาสตร" -> "คณิตศาสตร์"
3. ละเว้นแถวที่เป็นหัวตารางซ้ำซ้อน
4. ส่งออกผลลัพธ์เป็น JSON Array ของ Object ที่มีคีย์ตรงตามนี้เท่านั้น:
   - ที่
   - ประเภทวิชา
   - รหัสวิชา
   - ชื่อวิชา
   - ชม
   - หน่วยกิต
   - ชั้น
   - ห้องที่สอน
   - จำนวนห้อง
   - จำนวนชม
   - หมายเหตุ
"""

cleaned_records = []
batch_size = 15  # ส่งทีละ 15 แถวดิบ ป้องกันข้อมูลตกหล่นและประหยัด Token

print("กำลังส่งข้อมูลให้ Gemini Flash ช่วยจัดระเบียบและแก้ภาษาไทย...")

for i in range(0, len(df), batch_size):
    batch_df = df.iloc[i:i + batch_size]
    batch_csv = batch_df.to_csv(index=False)
    
    prompt = f"จัดระเบียบและคลีนข้อมูล CSV ชุดนี้ให้เป็น JSON Array ตามคำสั่ง:\n\n{batch_csv}"
    
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",  # รุ่นประหยัดที่สุดและเร็วมาก
            contents=prompt,
            config=types.GenerateContentConfig(
                system_instruction=system_instruction,
                response_mime_type="application/json",
                temperature=0.0
            )
        )
        batch_result = json.loads(response.text)
        if isinstance(batch_result, list):
            cleaned_records.extend(batch_result)
        print(f"ประมวลผลแล้ว {min(i + batch_size, len(df))}/{len(df)} แถวดิบ...")
    except Exception as e:
        print(f"เกิดข้อผิดพลาดที่ Batch {i}: {e}")
    
    time.sleep(0.5)  # เว้นช่วงสั้นๆ ป้องกัน Rate limit

# ==========================================
# 6. บันทึกผลลัพธ์ลง Excel และ CSV
# ==========================================
df_final = pd.DataFrame(cleaned_records)

print("\n--- ตัวอย่างข้อมูลที่คลีนและแตกแถวเรียบร้อยแล้ว ---")
display(df_final.head(15))

output_excel = "ตารางสอน_สมบูรณ์.xlsx"
df_final.to_excel(output_excel, index=False)
df_final.to_csv("ตารางสอน_สมบูรณ์.csv", index=False)

print(f"\nเสร็จสิ้น! ได้ตารางทั้งหมด {len(df_final)} แถว บันทึกไฟล์ไว้ที่ '{output_excel}' เรียบร้อยแล้ว")
