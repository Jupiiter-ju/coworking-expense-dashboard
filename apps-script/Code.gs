/**
 * Google Apps Script Web App — รวมข้อมูลจากทุกแท็บเดือน (มิ.ย.–ธ.ค. 69) ในสเปรดชีต
 * แล้วส่งกลับเป็น CSV เดียวให้ index.html
 *
 * ทำไมต้องอ่านตรงจากสเปรดชีต (SpreadsheetApp) แทนการดึง published CSV URL:
 * สเปรดชีตนี้แยกแต่ละเดือนเป็นคนละแท็บ (sheet) แต่ "Publish to web -> CSV"
 * export ได้แค่แท็บเดียว (แท็บที่ publish ไว้) เท่านั้น ทำให้เดือน ก.ค./ส.ค./ก.ย. หายไป
 * การเปิดสเปรดชีตตรงด้วย ID แล้ววนอ่านทุกแท็บ ทำให้ได้ข้อมูลครบทุกเดือน
 *
 * วิธี deploy / อัปเดต:
 * 1. ไปที่ https://script.google.com/ -> เปิดโปรเจกต์เดิมที่เคย deploy ไว้
 * 2. ลบโค้ดเดิมทั้งหมด แล้ววางไฟล์นี้แทน
 * 3. กด Deploy -> Manage deployments -> (ไอคอนดินสอ) Edit -> Version: "New version" -> Deploy
 *    (ใช้ deployment เดิม จะได้ URL /exec เดิม ไม่ต้องเปลี่ยนใน index.html)
 * 4. รอบนี้สคริปต์ต้องขอสิทธิ์เพิ่ม (อ่านสเปรดชีต) — จะมีหน้าจอ Authorize ขึ้นมาอีกรอบ
 *    เลือกบัญชี jupiiterlegacy@gmail.com -> Advanced -> Go to ... (unsafe) -> Allow
 */

const SPREADSHEET_ID = "1SPe9qvTeXNYExJ6gtGs2G74SE-PnCgeh4WrwfBizyD8";

// ชื่อแท็บเดือนตามลำดับที่อยากให้แสดงผล — เพิ่มเดือนใหม่ต่อท้ายได้เรื่อยๆ ตามที่สร้างแท็บจริง
const MONTH_SHEET_NAMES = [
  "มิ.ย. 69",
  "ก.ค. 69",
  "ส.ค. 69",
  "ก.ย. 69",
  "ต.ค. 69",
  "พ.ย. 69",
  "ธ.ค. 69"
];

function csvEscapeCell(val) {
  if (val === null || val === undefined) return "";
  let s = val instanceof Date ? val.toString() : String(val);
  if (/[",\n\r]/.test(s)) {
    s = '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const lines = [];

    MONTH_SHEET_NAMES.forEach(function (name) {
      const sheet = ss.getSheetByName(name);
      if (!sheet) return; // แท็บยังไม่ถูกสร้าง ข้ามไป

      const values = sheet.getDataRange().getValues();
      values.forEach(function (row) {
        lines.push(row.map(csvEscapeCell).join(","));
      });
    });

    const csv = lines.join("\n");
    return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
  } catch (err) {
    return ContentService
      .createTextOutput("ERROR: " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}
