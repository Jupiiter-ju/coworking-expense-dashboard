/**
 * Google Apps Script Web App — proxy สำหรับดึง CSV จาก Google Sheets ที่ publish ไว้
 * แล้วส่งกลับให้ index.html โดยไม่ติดปัญหา CORS (เพราะ UrlFetchApp ฝั่ง server
 * ไม่ถูกจำกัดด้วย CORS แบบที่ browser โดนจำกัด)
 *
 * วิธี deploy:
 * 1. ไปที่ https://script.google.com/ -> New project
 * 2. ลบโค้ดเริ่มต้นทั้งหมด แล้ววางไฟล์นี้แทน
 * 3. กด Deploy -> New deployment
 *    - Select type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. กด Deploy แล้ว copy "Web app URL" (ลงท้ายด้วย /exec)
 * 5. เอา URL นั้นมาใส่ในตัวแปร APPS_SCRIPT_CSV_URL ใน index.html
 *
 * หมายเหตุ: ทุกครั้งที่แก้โค้ดไฟล์นี้ ต้องทำ "New deployment" ใหม่
 * (หรือ Manage deployments -> Edit -> เปลี่ยนเวอร์ชันเป็น New version) ไม่งั้น URL เดิมจะยังใช้โค้ดเก่าอยู่
 */

const CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRynAPo__dwu0xNzSrYswKO-8Scm3yJCB-HXadUEj32N2cCZ9WSramZFxHcP8AwQvYM8mjHiCK8vHeZ/pub?output=csv";

function doGet(e) {
  try {
    const response = UrlFetchApp.fetch(CSV_URL, {
      muteHttpExceptions: true,
      followRedirects: true
    });

    if (response.getResponseCode() !== 200) {
      return ContentService
        .createTextOutput("ERROR: upstream returned " + response.getResponseCode())
        .setMimeType(ContentService.MimeType.TEXT);
    }

    return ContentService
      .createTextOutput(response.getContentText())
      .setMimeType(ContentService.MimeType.CSV);
  } catch (err) {
    return ContentService
      .createTextOutput("ERROR: " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}
