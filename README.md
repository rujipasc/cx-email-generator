# CardX Email Generator Tool 🚀

เครื่องมืออัตโนมัติสำหรับตรวจสอบและสร้างอีเมลพนักงานใหม่ของ CardX โดยตรวจสอบความซ้ำซ้อนผ่าน 3 เลเยอร์: Entra ID (Active), Recycle Bin (30 days), และ SharePoint History (Master Log).

## 📁 Folder Structure
เมื่อรันโปรแกรมครั้งแรก ระบบจะสร้างโฟลเดอร์เหล่านี้ให้โดยอัตโนมัติ:
- `inputs/`: วางไฟล์ Excel (.xlsx) ที่ต้องการประมวลผลที่นี่
- `processed_results/`: ไฟล์ผลลัพธ์ที่ใส่ Email ให้แล้วจะถูกเก็บที่นี่
- `archives/`: ไฟล์ต้นฉบับที่ประมวลผลเสร็จแล้วจะถูกย้ายมาเก็บที่นี่เพื่อป้องกันการรันซ้ำ

## 🛠 How to Use
1. ดาวน์โหลดไฟล์ `cardx-email-gen.exe` จากหน้า [Releases](https://github.com/cx-email-generator/releases).
2. นำไฟล์ไปวางในโฟลเดอร์ที่ต้องการใช้งาน.
3. นำไฟล์ Excel รายชื่อพนักงานวางไว้ในโฟลเดอร์ `inputs/`.
   - *หมายเหตุ: หัวตารางควรมีคอลัมน์ `Name`, `Surname`, และ `Employee ID`*
4. ดับเบิลคลิกที่ `cardx-email-gen.exe` เพื่อเริ่มทำงาน.
5. ตรวจสอบผลลัพธ์ในโฟลเดอร์ `processed_results/`.

## 🌐 Network & Security
- **Proxy:** ระบบตรวจสอบ Proxy จาก Windows Settings ให้อัตโนมัติ (รองรับเน็ตองค์กร).
- **SSL:** ข้ามการตรวจสอบ SSL Certificate (Bypass mode) เพื่อรองรับ SSL Inspection ภายในองค์กร.