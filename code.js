/**
 * FP Finished Good Stock — Executive Dashboard
 * -------------------------------------------------------------
 * วิธีติดตั้ง:
 * 1. เปิด Google Sheet ที่มีชีทข้อมูล (เช่น "FP 80E", "FP GNE", ... )
 * 2. เมนู Extensions > Apps Script
 * 3. วางไฟล์นี้ทับ Code.gs และสร้างไฟล์ HTML ชื่อ "Index" แล้ววางเนื้อหาของ Index.html
 * 4. กด Deploy > New deployment > Web app
 *      - Execute as: Me
 *      - Who has access: Anyone within your organization (หรือปรับตามต้องการ)
 * 5. เปิดลิงก์ Web app ที่ได้ — Dashboard จะดึงข้อมูลจากชีทนี้แบบสด (Real-time) ทุกครั้งที่โหลด/รีเฟรช
 *
 * โครงสร้างคอลัมน์ที่คาดหวังในแต่ละชีทสินค้า (แถวหัวข้ออยู่แถว 2-3, ข้อมูลเริ่มแถว 4):
 * A: Product Name | B: Batch No. | C: Mfg. Date | D: Qty In (Drum) | E: Kgs./Drum | F: Total In (Kgs.)
 * G: Record Date  | H: Qty Out (Drum, ต่อรายการ) | I: Kgs./Drum out | J: Total Out (Kgs., ต่อรายการ)
 * K: Delivered Date | L: PO No./Doc. No. | M: Remark
 * N: Qty Out รวมทั้ง Batch (Drum) | O: Remain (Drum) | P: Total Out รวม (Kgs.) | Q: Remain (Kgs.)
 */

// ตั้งค่าชื่อชีทที่ไม่ต้องการนำมาแสดง (ถ้ามี) — ปรับเพิ่มได้ตามต้องการ
var EXCLUDE_SHEETS = [];

var TZ = Session.getScriptTimeZone() || 'Asia/Bangkok';

var MONTHS = {
  'jan': 0, 'feb': 1, 'mar': 2, 'apr': 3, 'may': 4, 'jun': 5,
  'jul': 6, 'aug': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dec': 11
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('FP Finished Good Stock — Executive Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ฟังก์ชันหลักที่หน้าเว็บเรียกผ่าน google.script.run เพื่อดึงข้อมูลสดจาก Google Sheets
 */
function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var records = [];

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var name = sheet.getName();
    if (EXCLUDE_SHEETS.indexOf(name) !== -1) continue;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 4 || lastCol < 17) continue; // ชีทนี้ไม่มีข้อมูลพอ หรือไม่ตรงโครงสร้าง

    // ตรวจลายเซ็นหัวตาราง เพื่อข้ามชีทที่ไม่ใช่ชีทข้อมูลสต๊อก
    var a2 = sheet.getRange(2, 1).getValue();
    if (String(a2).trim().toLowerCase() !== 'product name') continue;

    var range = sheet.getRange(4, 1, lastRow - 3, 17).getValues();

    for (var i = 0; i < range.length; i++) {
      var row = range[i];
      var product = row[0];   // A
      var batch = row[1];     // B
      var qtyInDrum = row[3]; // D

      if (product === '' || product === null) continue;
      if (batch === '' || batch === null) continue;
      if (qtyInDrum === '' || qtyInDrum === null) continue;

      var mfgDateRaw = row[2];   // C
      var recordDateRaw = row[6]; // G
      var inKgsHeader = row[5];  // F

      var outDrumSum = row[13]; // N
      var remainDrum = row[14]; // O
      var outKgsSum = row[15];  // P
      var remainKgs = row[16];  // Q

      if (outDrumSum === '' || outDrumSum === null) outDrumSum = 0;
      if (outKgsSum === '' || outKgsSum === null) outKgsSum = 0;
      if (remainDrum === '' || remainDrum === null) remainDrum = qtyInDrum;
      if (remainKgs === '' || remainKgs === null) remainKgs = inKgsHeader || 0;

      var inDrumTotal = Number(outDrumSum) + Number(remainDrum);
      var inKgsTotal = Number(outKgsSum) + Number(remainKgs);

      var mfgDate = parseFlexibleDate_(mfgDateRaw);
      var recordDate = parseFlexibleDate_(recordDateRaw);

      records.push({
        sheet: name.trim(),
        product: String(product).trim(),
        batch: String(batch).trim(),
        mfgDateIso: mfgDate ? formatIso_(mfgDate) : null,
        mfgDateDisplay: mfgDate ? formatDisplay_(mfgDate) : (mfgDateRaw ? String(mfgDateRaw) : ''),
        recordDateIso: recordDate ? formatIso_(recordDate) : null,
        recordDateDisplay: recordDate ? formatDisplay_(recordDate) : (recordDateRaw ? String(recordDateRaw) : ''),
        inDrum: round2_(inDrumTotal),
        outDrum: round2_(Number(outDrumSum)),
        remainDrum: round2_(Number(remainDrum)),
        inKgs: round2_(inKgsTotal),
        outKgs: round2_(Number(outKgsSum)),
        remainKgs: round2_(Number(remainKgs))
      });
    }
  }

  return {
    records: records,
    generatedAtIso: formatIso_(new Date()),
    generatedAtDisplay: Utilities.formatDate(new Date(), TZ, 'dd-MMM-yyyy HH:mm')
  };
}

function round2_(n) {
  if (isNaN(n)) return 0;
  return Math.round(n * 100) / 100;
}

/**
 * รองรับทั้งกรณี Google Sheets ตีความเซลล์เป็น Date object แล้ว
 * และกรณีเป็นข้อความรูปแบบ "16-Nov-23" / "2-Apr-2024" เป็นต้น
 */
function parseFlexibleDate_(value) {
  if (!value && value !== 0) return null;

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }

  var str = String(value).trim();
  if (str === '') return null;

  var m = str.match(/^(\d{1,2})[-\/\s]([A-Za-z]{3,})[-\/\s](\d{2,4})$/);
  if (m) {
    var day = parseInt(m[1], 10);
    var monKey = m[2].substring(0, 3).toLowerCase();
    var monIdx = MONTHS[monKey];
    var year = parseInt(m[3], 10);
    if (year < 100) year += 2000;
    if (monIdx !== undefined) {
      var d = new Date(year, monIdx, day);
      if (!isNaN(d.getTime())) return d;
    }
  }

  // รูปแบบ dd/mm/yyyy หรือ d-m-yyyy
  var m2 = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
  if (m2) {
    var dd = parseInt(m2[1], 10);
    var mm = parseInt(m2[2], 10) - 1;
    var yy = parseInt(m2[3], 10);
    if (yy < 100) yy += 2000;
    var d2 = new Date(yy, mm, dd);
    if (!isNaN(d2.getTime())) return d2;
  }

  var fallback = new Date(str);
  if (!isNaN(fallback.getTime())) return fallback;

  return null;
}

function formatIso_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function formatDisplay_(d) {
  return Utilities.formatDate(d, TZ, 'dd-MMM-yy');
}
