# Thailand Motorcycle Tire Shops Scraper (Google Maps -> GitHub)

ระบบดึงข้อมูลร้านยางมอเตอร์ไซค์ทั่วประเทศไทย (77 จังหวัด + 50 เขตในกรุงเทพฯ) จาก Google Maps อัตโนมัติผ่าน **GitHub Actions** โดยไม่ต้องติดตั้งโปรแกรมลงในเครื่องคอมพิวเตอร์

## โครงสร้างไฟล์ใน Repository
- `.github/workflows/scrape.yml` — สคริปต์สั่งรัน `gosom/google-maps-scraper` และคัดกรองข้อมูลอัตโนมัติบนคลาวด์ของ GitHub
- `queries.txt` — รายการคำค้นหาครอบคลุมทั้ง 77 จังหวัด, 50 เขตในกรุงเทพฯ และหัวเมืองหลัก (รวม 240+ คำค้นหา)
- `filter_shops.py` — สคริปต์ Python สำหรับลบข้อมูลซ้ำ คัดแยกร้านยางมอเตอร์ไซค์ออกจากร้านยางรถยนต์ 3 ชั้น (หมวดหมู่ + ชื่อร้าน + ข้อความรีวิว) แท็กแบรนด์ยางที่ถูกพูดถึงในรีวิว และรองรับ Gemini Vision AI สำหรับเช็กรูปปกหน้าร้านก้ำกึ่ง
- `requirements.txt` — ไลบรารี Python ที่จำเป็น (`pandas`)

## วิธีใช้งานบน GitHub (ทำผ่านหน้าเว็บ 100%)
1. สร้าง Repository ใหม่บน GitHub (แนะนำตั้งเป็น **Private**)
2. อัปโหลดไฟล์ทั้งหมดในชุดนี้เข้า Repository โดยให้ไฟล์ `scrape.yml` อยู่ในโฟลเดอร์ `.github/workflows/scrape.yml`
3. *(ทางเลือก)* หากต้องการใช้ Gemini Vision AI ช่วยเช็กรูปหน้าร้านที่ชื่อก้ำกึ่ง ให้ไปที่ **Settings > Secrets and variables > Actions > New repository secret** ตั้งชื่อ `GEMINI_API_KEY` แล้วใส่คีย์ของคุณ
4. ไปที่แท็บ **Actions** ด้านบน > เลือก **Scrape Thailand Motorcycle Tire Shops** > กดปุ่ม **Run workflow**
5. เมื่อรันเสร็จ ระบบจะบันทึกไฟล์ `raw_results.csv` และ `clean_motorcycle_tire_shops.csv` (เข้ารหัส UTF-8-SIG เปิดใน Excel ภาษาไทยไม่เพี้ยน) กลับเข้ามาใน Repository อัตโนมัติ
