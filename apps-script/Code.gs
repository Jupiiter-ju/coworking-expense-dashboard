/**
 * Google Apps Script Web App — รวมข้อมูลจากทุกแท็บเดือน (มิ.ย.–ธ.ค. 69) ในสเปรดชีตค่าใช้จ่าย
 * บวกกับข้อมูลการใช้งาน (booking) จากสเปรดชีตอีกตัว แล้วส่งกลับเป็น CSV เดียวให้ index.html
 *
 * ทำไมต้องอ่านตรงจากสเปรดชีต (SpreadsheetApp) แทนการดึง published CSV URL:
 * สเปรดชีตนี้แยกแต่ละเดือนเป็นคนละแท็บ (sheet) แต่ "Publish to web -> CSV"
 * export ได้แค่แท็บเดียว (แท็บที่ publish ไว้) เท่านั้น ทำให้เดือน ก.ค./ส.ค./ก.ย. หายไป
 * การเปิดสเปรดชีตตรงด้วย ID แล้ววนอ่านทุกแท็บ ทำให้ได้ข้อมูลครบทุกเดือน
 *
 * ข้อมูลการใช้งาน (usage/booking) อยู่คนละสเปรดชีต แต่เป็นไฟล์ที่เปิดดูได้แบบสาธารณะ
 * (Anyone with the link) จึงดึงผ่าน UrlFetchApp ตรงๆ (ไม่ต้องขอสิทธิ์เพิ่ม เพราะ
 * external_request scope ได้รับอนุญาตอยู่แล้ว) แล้วต่อท้ายด้วย marker "###USAGE###"
 * ให้ index.html แยกส่วนนี้ออกจาก CSV หลักได้
 *
 * วิธี deploy / อัปเดต:
 * 1. ไปที่ https://script.google.com/ -> เปิดโปรเจกต์เดิมที่เคย deploy ไว้ (ผูกกับสเปรดชีต
 *    ค่าใช้จ่าย 1SPe9qv... — เปิดจาก Extensions > Apps Script ในสเปรดชีตนั้นจะตรงที่สุด)
 * 2. ลบโค้ดเดิมทั้งหมด แล้ววางไฟล์นี้แทน
 * 3. กด Deploy -> Manage deployments -> (ไอคอนดินสอ) Edit -> Version: "New version" -> Deploy
 *    (ใช้ deployment เดิม จะได้ URL /exec เดิม ไม่ต้องเปลี่ยนใน index.html)
 * 4. ถ้าโค้ดใช้บริการ Google ใหม่ที่ยังไม่เคยขอสิทธิ์ (เช่น SpreadsheetApp หรือ UrlFetchApp
 *    ครั้งแรก) ให้เลือกฟังก์ชัน doGet ในตัวแก้ไข กด Run (เรียกใช้) หนึ่งครั้ง จะมี popup
 *    "จำเป็นต้องให้สิทธิ์" ขึ้นมา -> Review permissions -> เลือกบัญชี jupiiterlegacy@gmail.com
 *    -> Advanced -> Go to ... (unsafe) -> Allow
 *    (แค่ deploy ใหม่อย่างเดียวไม่พอ ต้อง Run ให้ authorize จริงก่อน แล้วค่อย deploy)
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

// แท็บค่าใช้ห้องประชุม — แยกเป็นสเปรดชีตเดียวกันแต่คนละแท็บ มีบล็อกแยกตามเดือน
// (มิถุนายน 2569, กรกฎาคม 2569, ...) พร้อมคอลัมน์ No./Transaction Date/Name/
// Account used Booking/Credit Card/Price/Invoice Number
const MEETING_ROOM_SHEET_NAME = "MeetingRoom";

// สเปรดชีตข้อมูลการใช้งาน (booking) — เปิดดูได้แบบสาธารณะ ดึงตรงผ่าน UrlFetchApp
const USAGE_CSV_URL = "https://docs.google.com/spreadsheets/d/1CvaE43lwSzJtnAbtTV0rr15KEHH6y9v6G5bQpMNTt2U/export?format=csv&gid=0";
const USAGE_MARKER = "###USAGE###";

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

      // แนบชื่อแท็บจริงไว้เป็น marker แทนการให้ฝั่ง client เดาเดือนจากข้อความหัวตาราง
      // (ข้อความ "รอบบิล ..." ในชีตเปลี่ยนรูปแบบไปแล้วหลายครั้ง แต่ชื่อแท็บคงที่)
      lines.push("###MONTH:" + name + "###");

      const values = sheet.getDataRange().getValues();
      values.forEach(function (row) {
        lines.push(row.map(csvEscapeCell).join(","));
      });
    });

    const meetingRoomSheet = ss.getSheetByName(MEETING_ROOM_SHEET_NAME);
    if (meetingRoomSheet) {
      const mrValues = meetingRoomSheet.getDataRange().getValues();
      mrValues.forEach(function (row) {
        lines.push(row.map(csvEscapeCell).join(","));
      });
    }

    lines.push(USAGE_MARKER);
    try {
      const usageRes = UrlFetchApp.fetch(USAGE_CSV_URL, { muteHttpExceptions: true, followRedirects: true });
      if (usageRes.getResponseCode() === 200) {
        lines.push(usageRes.getContentText());
      }
    } catch (usageErr) {
      // ดึงข้อมูล usage ไม่ได้ ไม่ให้กระทบข้อมูลหลัก
    }

    const csv = lines.join("\n");
    return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
  } catch (err) {
    return ContentService
      .createTextOutput("ERROR: " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}
