# ABM v4.5 — PWA + Google Sheet Sync

## 1) ตั้ง Google Sheet + Apps Script

1. สร้าง Google Sheet ใหม่ → เมนู **ส่วนขยาย \> Apps Script**
2. ลบโค้ดเดิม วางเนื้อหา `Code.gs` แล้วเปลี่ยน `TOKEN` เป็นรหัสผ่านของคุณ → บันทึก
3. **Deploy \> New deployment \> Web app** — Execute as: **Me**, Who has access: **Anyone** → Deploy → อนุญาตสิทธิ์ → คัดลอก URL (ลงท้าย `/exec`)
4. แก้โค้ดภายหลัง ต้อง Deploy \> Manage deployments \> Edit \> New version
