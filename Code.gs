/**
 * BOWLING 7 COMPANIES - LIVE SCOREBOARD SYNC (Google Apps Script)
 * เชื่อมต่อกับ Google Sheet:
 * https://docs.google.com/spreadsheets/d/1mpOiyu9x7dv3web2uuYUVJeK3XMjwdb5g6VoUqXWVss/edit?gid=1219114810
 *
 * วิธีติดตั้ง:
 * 1. เปิด Google Sheet > เมนู Extensions (ส่วนขยาย) > Apps Script
 * 2. ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดนี้ลงไป > กด Save (บันทึก)
 * 3. (แนะนำ) รันฟังก์ชัน setupTriggers หรือกดเมนู "🎳 Bowling Sync > ตั้งค่า Auto-Sync Trigger" ใน Sheet
 * 4. กด Deploy > New deployment > เลือก Type: Web app
 *    - Execute as: Me (ตัวคุณเอง)
 *    - Who has access: Anyone (ทุกคน)
 * 5. คัดลอก Web App URL (ลงท้ายด้วย /exec) ไปใส่ในหน้า Dashboard (ปุ่ม ⚙️ ตั้งค่าซิงค์)
 *    หรือวางในตัวแปร APPS_SCRIPT_URL ในไฟล์ index.html
 */

const TARGET_SHEET_GID = 1219114810;
const FALLBACK_SHEET_NAME = 'กรอกคะแนน';
const PROP_VERSION_KEY = 'BOWLING_SYNC_VERSION';
const PROP_UPDATED_AT_KEY = 'BOWLING_SYNC_UPDATED_AT';

/**
 * สร้างเมนูช่วยจัดการใน Google Sheets
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎳 Bowling Sync')
    .addItem('⚡ ตั้งค่า Auto-Sync Trigger (รันครั้งเดียว)', 'setupTriggers')
    .addItem('🔄 บังคับรีเฟรชเวอร์ชันซิงค์ตอนนี้', 'bumpSyncVersion')
    .addItem('📊 ดูสถานะซิงค์ล่าสุด', 'showSyncStatus')
    .addToUi();
}

/**
 * ค้นหาชีตเป้าหมายจาก GID (รองรับแม้เปลี่ยนชื่อแท็บ)
 */
function getTargetSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === TARGET_SHEET_GID) {
      return sheets[i];
    }
  }
  return ss.getSheetByName(FALLBACK_SHEET_NAME) || sheets[0];
}

/**
 * Simple Trigger: ทำงานอัตโนมัติทุกครั้งที่มีคนแก้ไขเซลล์ใน Sheet
 */
function onEdit(e) {
  handleSheetMutation_(e);
}

/**
 * Installable Trigger: รองรับทั้งการแก้ไข, Paste, และการอัปเดตโครงสร้าง
 */
function onSheetChangeInstallable(e) {
  handleSheetMutation_(e);
}

function handleSheetMutation_(e) {
  try {
    if (e && e.range) {
      const editedSheet = e.range.getSheet();
      if (editedSheet.getSheetId() !== TARGET_SHEET_GID && editedSheet.getName() !== FALLBACK_SHEET_NAME) {
        return;
      }
    }
    SpreadsheetApp.flush(); // บังคับคำนวณสูตรคอลัมน์ K, High Game, อันดับ ทันที
    bumpSyncVersion();
  } catch (err) {
    console.error('handleSheetMutation_ error:', err);
  }
}

/**
 * อัปเดตเลข Version และ Timestamp ใน ScriptProperties
 */
function bumpSyncVersion() {
  const props = PropertiesService.getScriptProperties();
  const current = parseInt(props.getProperty(PROP_VERSION_KEY) || '0', 10) || 0;
  const nextVersion = current + 1;
  const nowIso = new Date().toISOString();
  props.setProperties({
    [PROP_VERSION_KEY]: String(nextVersion),
    [PROP_UPDATED_AT_KEY]: nowIso
  });
  return { version: nextVersion, updatedAt: nowIso };
}

/**
 * ติดตั้ง Installable Triggers เพื่อความเสถียรตลอดเวลา
 */
function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'onSheetChangeInstallable') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onSheetChangeInstallable')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  const state = bumpSyncVersion();
  SpreadsheetApp.getUi().alert(
    '✅ ตั้งค่า Auto-Sync Trigger เรียบร้อยแล้ว!\n' +
    'Version: ' + state.version + '\n' +
    'ต่อไปสามารถกด Deploy > New deployment > Web app ได้เลยครับ'
  );
}

function showSyncStatus() {
  const props = PropertiesService.getScriptProperties();
  const sheet = getTargetSheet_();
  const version = props.getProperty(PROP_VERSION_KEY) || '1';
  const updatedAt = props.getProperty(PROP_UPDATED_AT_KEY) || 'ยังไม่มีการบันทึก';
  SpreadsheetApp.getUi().alert(
    '🎳 Bowling Sync Status\n' +
    'Sheet: ' + sheet.getName() + ' (GID: ' + sheet.getSheetId() + ')\n' +
    'Rows: ' + sheet.getLastRow() + '\n' +
    'Version: ' + version + '\n' +
    'Last Updated: ' + updatedAt
  );
}

/**
 * Web App Endpoint: ให้หน้า index.html ดึงคะแนนสด (รองรับทั้ง JSON และ JSONP)
 * - ?mode=version  -> คืนค่าเฉพาะ version & updatedAt (เช็คไวมาก)
 * - default        -> คืนค่า rows ทั้งหมด + version + updatedAt
 * - ?callback=fn   -> คืนค่าแบบ JSONP (ใช้ได้แม้เปิด index.html แบบ file://)
 */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const mode = params.mode || 'full';
  const callback = params.callback || '';

  const props = PropertiesService.getScriptProperties();
  const version = parseInt(props.getProperty(PROP_VERSION_KEY) || '1', 10) || 1;
  const updatedAt = props.getProperty(PROP_UPDATED_AT_KEY) || new Date().toISOString();

  let payload;
  if (mode === 'version') {
    payload = {
      ok: true,
      mode: 'version',
      version: version,
      updatedAt: updatedAt
    };
  } else {
    const sheet = getTargetSheet_();
    const values = sheet.getDataRange().getValues();
    payload = {
      ok: true,
      mode: 'full',
      sheetName: sheet.getName(),
      gid: sheet.getSheetId(),
      version: version,
      updatedAt: updatedAt,
      rows: values
    };
  }

  const json = JSON.stringify(payload);

  // รองรับ JSONP เผื่อกรณีเปิดไฟล์ HTML โดยตรงจากเครื่อง (file://)
  if (callback && /^[a-zA-Z_$][0-9a-zA-Z_$.]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
