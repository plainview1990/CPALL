โจทย์ข้อ 2 ดำเนินการดังนี้
1. เตรียมไฟล์ Excel ที่เป็น source data ใน local path ที่พร้อมใช้งาน โดยต้องไม่มีไฟล์อื่นที่ไม่เกี่ยวข้องใน path นี้ ในตัวอย่างของผู้ใช้ เก็บไฟล์ไว้ที่ C:\Data Science\CPALL\01.xlsx - 03.xlsx เป็นต้น
2. เปิดไฟล์ Excel ใหม่ขึ้นมา 1 ไฟล์
3. import File Module.bas บน VBA Editor
4. Run macro ตัวที่ 1 ชื่อ ConslidateWorkbooks_1 สำหรับการรวมไฟล์ excel 01-03.xlsx ให้เป็นไฟล์เดียว
5. Run macro ตัวที่ 2 ชื่อ DeleteWorksheet_2 สำหรับลบ worksheet ที่ไม่ต้องการออกไป ในที่นี้คือ worksheet ที่ชื่อ "Sheet1"
6. Run macro ตัวที่ 3 ชื่อ SortWorksheets_3 สำหรับการเรียงหน้า worksheet โดยหลังจากสั่งรัน macro แล้ว จะมีให้เลือกกด Yes สำหรับการเรียงเลขหน้า worksheet จากน้อยไปมาก (Ascending)
7. มาที่ VBA Editor : ไปที่ Tab Tools > References > ให้ติ๊กเครื่องหมายถูก : Microsoft PowerPoint 16.0 Object Library > OK สำหรับการทำงาน Macro ร่วมกับ MS.Powerpoint
8. Run Macro ตัวที่ 4 CopyMultiObject_4 สำหรับการแสดงผล จาก Excel ไปที่ Powerpoint บนเงื่อนไข ให้เลือกข้อมูลทั้งหมดจากทุก worksheet ของ Excel มาแสดงบน slide หน้าเดียวของ PowerPoint
