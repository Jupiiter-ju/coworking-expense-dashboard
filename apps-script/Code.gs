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
 * แคช (PropertiesService) + Time-driven trigger:
 * เดิม doGet() อ่านสเปรดชีตสด + ยิง UrlFetchApp ทุกครั้งที่มีคนเปิดหน้าเว็บ ซึ่งช้าและ
 * เสี่ยงโดน quota ถ้ามีคนเข้าพร้อมกันเยอะ ตอนนี้เปลี่ยนเป็น: มีฟังก์ชัน refreshCache()
 * ที่คำนวณ CSV จริงแล้วเก็บลง PropertiesService (แบ่งเป็นชิ้นๆ เพราะแต่ละ property
 * เก็บได้ไม่เกิน 9KB) ทำงานอัตโนมัติวันละครั้งผ่าน time-driven trigger ส่วน doGet()
 * ปกติจะอ่านจากแคชนี้ตรงๆ (เร็วมาก) ยกเว้นมี query param ?fresh=1 (ปุ่ม "รีเฟรช" ใน
 * index.html) ถึงจะคำนวณสดแล้วอัปเดตแคชไปด้วย หรือถ้ายังไม่เคยมีแคชเลยก็คำนวณสดครั้งแรก
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
 * 5. ตั้ง trigger อัตโนมัติ (ทำครั้งเดียว ไม่ต้องทำซ้ำทุกครั้งที่ deploy ใหม่): เลือกฟังก์ชัน
 *    setupDailyTrigger ในตัวแก้ไข -> กด Run หนึ่งครั้ง (ครั้งแรกจะขอสิทธิ์เพิ่มเรื่อง trigger
 *    ให้กด Allow เหมือนขั้นตอนที่ 4) จะได้ trigger ที่เรียก refreshCache() ทุกวันตอนตี 3
 *    เช็คได้ที่เมนูนาฬิกา (Triggers) ทางซ้ายของตัวแก้ไข
 * 6. รัน refreshCache() หนึ่งครั้งด้วยตัวเอง (กด Run) เพื่อสร้างแคชชุดแรกทันที ไม่ต้องรอถึงตี 3
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

// ตั้งค่าแคช — เก็บใน Script Properties เป็นชิ้นๆ เพราะแต่ละ property ใส่ได้ไม่เกิน 9KB
const CACHE_CHUNK_PREFIX = "CSV_CACHE_CHUNK_";
const CACHE_CHUNK_COUNT_KEY = "CSV_CACHE_CHUNK_COUNT";
const CACHE_TIME_KEY = "CSV_CACHE_TIME";
const CACHE_CHUNK_SIZE = 8000;

function csvEscapeCell(val) {
  if (val === null || val === undefined) return "";
  let s = val instanceof Date ? val.toString() : String(val);
  if (/[",\n\r]/.test(s)) {
    s = '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

/**
 * คำนวณ CSV รวมทั้งหมดแบบสด (อ่านสเปรดชีต + ยิง UrlFetchApp) — ใช้ทั้งตอน refreshCache()
 * (รันตาม trigger) และตอน doGet() เจอ ?fresh=1 หรือยังไม่มีแคชเลย
 */
function buildCsv() {
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

  return lines.join("\n");
}

/** เก็บ CSV ลง Script Properties แบบแบ่งชิ้น พร้อมล้างชิ้นเก่าที่เหลือทิ้งก่อน */
function saveCsvToCache(csv) {
  const props = PropertiesService.getScriptProperties();

  const oldCount = parseInt(props.getProperty(CACHE_CHUNK_COUNT_KEY) || "0", 10);
  for (let i = 0; i < oldCount; i++) {
    props.deleteProperty(CACHE_CHUNK_PREFIX + i);
  }

  const chunkKeys = {};
  let chunkCount = 0;
  for (let i = 0; i < csv.length; i += CACHE_CHUNK_SIZE) {
    chunkKeys[CACHE_CHUNK_PREFIX + chunkCount] = csv.slice(i, i + CACHE_CHUNK_SIZE);
    chunkCount++;
  }
  chunkKeys[CACHE_CHUNK_COUNT_KEY] = String(chunkCount);
  chunkKeys[CACHE_TIME_KEY] = new Date().toISOString();
  props.setProperties(chunkKeys, false);
}

/** อ่าน CSV จากแคช คืนค่า null ถ้ายังไม่เคยมีแคช หรือแคชเสียหาย (ชิ้นหาย) */
function readCsvFromCache() {
  const props = PropertiesService.getScriptProperties();
  const count = parseInt(props.getProperty(CACHE_CHUNK_COUNT_KEY) || "0", 10);
  if (count === 0) return null;

  const parts = [];
  for (let i = 0; i < count; i++) {
    const part = props.getProperty(CACHE_CHUNK_PREFIX + i);
    if (part === null) return null;
    parts.push(part);
  }
  return parts.join("");
}

/**
 * ฟังก์ชันสำหรับ time-driven trigger (ตั้งด้วย setupDailyTrigger) — คำนวณ CSV สด
 * แล้วเขียนทับแคช เรียกวันละครั้งอัตโนมัติ หรือจะกด Run เองตอนไหนก็ได้เพื่อบังคับอัปเดต
 */
function refreshCache() {
  const csv = buildCsv();
  saveCsvToCache(csv);
}

/**
 * ตั้ง time-driven trigger ให้เรียก refreshCache() ทุกวันตอนตี 3 (ช่วงคนใช้งานน้อยที่สุด)
 * รันฟังก์ชันนี้เองหนึ่งครั้งในตัวแก้ไข (เลือกชื่อฟังก์ชันนี้แล้วกด Run) — ไม่ต้องรันซ้ำอีก
 * เว้นแต่ trigger หาย หรืออยากเปลี่ยนเวลา (ฟังก์ชันนี้ลบ trigger เก่าของ refreshCache ก่อน
 * สร้างใหม่เสมอ กันสร้างซ้ำซ้อนถ้าเผลอรันหลายครั้ง)
 */
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "refreshCache") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("refreshCache")
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}

function doGet(e) {
  try {
    const forceFresh = !!(e && e.parameter && e.parameter.fresh === "1");

    if (!forceFresh) {
      const cached = readCsvFromCache();
      if (cached !== null) {
        return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.CSV);
      }
      // ยังไม่เคยมีแคชเลย (เช่น deploy ใหม่ครั้งแรก) -> คำนวณสดแล้วเก็บแคชไว้เลย
    }

    const csv = buildCsv();
    saveCsvToCache(csv);
    return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
  } catch (err) {
    return ContentService
      .createTextOutput("ERROR: " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}
