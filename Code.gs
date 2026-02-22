// ==========================================
// 1. ส่วนตั้งค่าระบบ (CONFIGURATION)
// ==========================================

const SPREADSHEET_ID = '1Kr1GOn5F8rBNJGA7_Sqp4be7JirrRvVci0AGhtkA5hQ'; // ID ของ Google Sheet
const FOLDER_ID = '1pPtZlI8XYBle02byB5lthAhtLX8012Pa'; // ID ของ Google Drive Folder
const FOLDER_RETURN_ID = '16Rn35Lv0gC3HRt2ohUWmi_DdQShRcUfn'; // โฟลเดอร์เก็บเอกสารการคืน iPad

const SHEET_NAMES = {
  STUDENTS: [
    'รายชื่อนักเรียน ม.3', 
    'รายชื่อนักเรียน ม.4', 
    'รายชื่อนักเรียน ม.5', 
    'รายชื่อนักเรียน ม.6'
  ],
  TEACHERS: 'รายชื่อครู',
  ASSETS: 'รายงานทะเบียนทรัพย์สิน',
  ALL_NAMES: 'รายชื่อทั้งหมด',
  DATA_DB: 'ข้อมูล',
  LOGS: 'Log',
  ADMIN: 'แอดมิน',
  ADVISOR: 'ครูที่ปรึกษา'
};

// ==========================================
// 2. ฟังก์ชันพื้นฐาน & Helper
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('ระบบตรวจสอบสถานะ iPad โรงเรียนอรัญประเทศ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { 
  return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}

/** เคลียร์ Cache ข้อมูลหลักเมื่อมีการอัปเดต */
function invalidateSystemDataCache() {
  try {
    const cache = CacheService.getScriptCache();
    // ลบทั้ง key หลักและ chunk keys
    cache.remove('systemData_' + SPREADSHEET_ID);
    cache.remove('systemData_' + SPREADSHEET_ID + '_meta');
    cache.remove('systemData_' + SPREADSHEET_ID + '_1');
    cache.remove('systemData_' + SPREADSHEET_ID + '_2');
    cache.remove('systemData_' + SPREADSHEET_ID + '_3');
    cache.remove('systemData_' + SPREADSHEET_ID + '_4');
  } catch (_) {}
}

/** ฟังก์ชันปรับชื่อให้สะอาดที่สุดเพื่อการจับคู่ที่แม่นยำ */
function normalizeName(name) {
  if (!name) return "";
  let n = name.toString().normalize('NFC');
  const titleRegex = /^(?:ว่าที่\s*ร(?:้อย)?\.?\s*[ตทพ]\.?|จ(?:่า)?\.?ส(?:ิบ)?\.?[อทต]\.?|ส(?:ิบ)?\.?[อทต]\.?|พล(?:ทหาร)?\.?|ส\.อ\.|จ\.ส\.อ\.|ร\.ต\.|ดร\.?|ผศ\.?|รศ\.?|ศ\.?|เด็กชาย|เด็กหญิง|นางสาว|ด\.?\s*ช\.?|ด\.?\s*ญ\.?|น\.?\s*ส\.?|นาย|นาง|ครู|อ\.?|mr\.?|mrs\.?|ms\.?|miss)[\s\.]*/gi;
  n = n.replace(titleRegex, ''); 
  n = n.replace(/[^ก-๙a-zA-Z]/g, ''); 
  return n.toLowerCase();
}

/** Levenshtein Distance สำหรับ Fuzzy Match */
function getEditDistance(a, b) {
  if (a.length === 0) return b.length; 
  if (b.length === 0) return a.length; 
  var matrix = [];
  for (var i = 0; i <= b.length; i++) { matrix[i] = [i]; }
  for (var j = 0; j <= a.length; j++) { matrix[0][j] = j; }
  for (var i = 1; i <= b.length; i++) {
    for (var j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) == a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1)
        );
      }
    }
  }
  return matrix[b.length][a.length];
}

// ==========================================
// 3. ฟังก์ชันดึงข้อมูลหลัก (MAIN DATA ENGINE)
// ==========================================

function getAllSystemData() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'systemData_' + SPREADSHEET_ID;

  // ── FIX #6: รองรับ Cache หลาย Chunk (กัน overflow 100KB) ──
  try {
    const metaRaw = cache.get(cacheKey + '_meta');
    if (metaRaw) {
      const meta = JSON.parse(metaRaw);
      let combined = [];
      let allFound = true;
      for (let c = 1; c <= meta.chunks; c++) {
        const part = cache.get(cacheKey + '_' + c);
        if (!part) { allFound = false; break; }
        combined = combined.concat(JSON.parse(part));
      }
      if (allFound && combined.length > 0) return combined;
    } else {
      // ลอง key เดิม (ไม่มี chunk)
      const cached = cache.get(cacheKey);
      if (cached) {
        try { return JSON.parse(cached); } catch (_) {}
      }
    }
  } catch (_) {}

  // ── FIX #7: LockService ป้องกัน Thundering Herd ──
  // ถ้ามีหลาย request พร้อมกัน ให้คิวรอ request แรกที่คำนวณเสร็จ
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // รอสูงสุด 15 วินาที
  } catch (e) {
    throw new Error("ระบบยุ่งอยู่ กรุณารีเฟรชใหม่อีกครั้ง");
  }

  // ตรวจ Cache อีกครั้งหลังได้ Lock (อาจมี request ก่อนหน้าคำนวณเสร็จแล้ว)
  try {
    const metaRaw2 = cache.get(cacheKey + '_meta');
    if (metaRaw2) {
      const meta2 = JSON.parse(metaRaw2);
      let combined2 = [];
      let allFound2 = true;
      for (let c = 1; c <= meta2.chunks; c++) {
        const part = cache.get(cacheKey + '_' + c);
        if (!part) { allFound2 = false; break; }
        combined2 = combined2.concat(JSON.parse(part));
      }
      if (allFound2 && combined2.length > 0) {
        lock.releaseLock();
        return combined2;
      }
    } else {
      const cached2 = cache.get(cacheKey);
      if (cached2) {
        try {
          const parsed2 = JSON.parse(cached2);
          lock.releaseLock();
          return parsed2;
        } catch (_) {}
      }
    }
  } catch (_) {}

  let allPeople = [];
  try {
    // ── เปิด SpreadsheetApp ครั้งเดียว แล้วส่งต่อทุกฟังก์ชัน ──
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
    // --- A. ดึงข้อมูล Asset (ทะเบียนทรัพย์สิน) ---
    const assetSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
    let assetMap = {};
    let assetKeys = [];

    if (assetSheet) {
      const lastRow = assetSheet.getLastRow();
      if (lastRow > 1) {
        const assetData = assetSheet.getRange(1, 1, lastRow, 10).getValues();
        for (let i = 1; i < assetData.length; i++) {
          let rawName = assetData[i][4]; 
          let serial = assetData[i][2] || assetData[i][3]; 
          if (rawName) {
            let cName = normalizeName(rawName);
            if (cName.length > 0) {
              assetMap[cName] = { 
                serial: serial ? serial.toString() : '-', 
                status: 'ยืมอยู่' 
              };
              assetKeys.push(cName);
            }
          }
        }
      }
    }

    // --- B. ดึงข้อมูล Log (Database) ---
    const dbSheet = ss.getSheetByName(SHEET_NAMES.DATA_DB);
    let dbMap = {}; 
    
    if (dbSheet) {
      const dbLastRow = dbSheet.getLastRow();
      if (dbLastRow > 1) {
        const dbData = dbSheet.getRange(1, 1, dbLastRow, 14).getValues();
        for (let i = 1; i < dbData.length; i++) {
          let id = dbData[i][1];
          if (!id) continue;
          id = id.toString();

          let rowSerial = dbData[i][5] ? dbData[i][5].toString() : '';
          let rowStatus = dbData[i][8] ? dbData[i][8].toString() : '';
          let rowNote   = dbData[i][7] ? dbData[i][7].toString() : '';
          let rowAction = dbData[i][6] ? dbData[i][6].toString() : '';
          let hasFiles  = (dbData[i][9] || dbData[i][10] || dbData[i][11] || dbData[i][12]);
          
          if (!dbMap[id]) dbMap[id] = {
            borrowStatus: 'ยังไม่ยืม', docStatus: 'ยังไม่ส่ง',
            serial: '-', files: {}, returnDocUrl: '',
            docFinal: false
          };
          
          if (rowSerial && rowSerial !== '-' && rowSerial !== '') dbMap[id].serial = rowSerial;
          
          if (hasFiles) {
            dbMap[id].files = { 
              agreement: dbData[i][9], card_std: dbData[i][10], 
              card_parent: dbData[i][11], house: dbData[i][12], phone: dbData[i][13] 
            };
            if (!dbMap[id].docFinal && !rowStatus.includes('ADMIN') && !rowStatus.includes('ADVISOR')) {
              dbMap[id].docStatus = 'รอตรวจสอบ';
            }
          }

          if (rowStatus === 'เอกสารผ่าน' || rowStatus.includes('เอกสารผ่าน') || rowAction === 'ADVISOR_APPROVE') {
            dbMap[id].docStatus = 'เอกสารผ่าน';
            dbMap[id].docFinal = true;
          } else if (rowStatus === 'เอกสารไม่ผ่าน' || rowStatus.includes('ไม่ผ่าน')) {
            dbMap[id].docStatus = 'เอกสารไม่ผ่าน';
            dbMap[id].docFinal = true;
          } else if (!dbMap[id].docFinal) {
            if (rowStatus.includes('รอตรวจสอบเอกสาร') || rowStatus.includes('รอตรวจสอบ')) {
              dbMap[id].docStatus = 'รอตรวจสอบ';
            } else if (rowStatus === 'ยังไม่ส่ง') {
              dbMap[id].docStatus = 'ยังไม่ส่ง';
            }
          }

          if ((rowAction === 'ADVISOR_RETURN' || rowAction === 'USER_RETURN') && rowNote.indexOf('[หลักฐานการคืน]:') >= 0) {
            const match = rowNote.match(/\[หลักฐานการคืน\]:\s*(https?:\/\/[^\s\n]+)/);
            if (match) dbMap[id].returnDocUrl = match[1].trim();
          }

          if (rowStatus.indexOf('อยู่ระหว่างการส่งคืน') >= 0) {
            dbMap[id].borrowStatus = 'อยู่ระหว่างการส่งคืน';
          } else if (rowStatus === 'คืนแล้ว') {
            dbMap[id].borrowStatus = 'คืนแล้ว';
          } else if (rowStatus.includes('ยืมอยู่') || rowStatus === 'ยืมได้') {
            dbMap[id].borrowStatus = 'ยืมอยู่';
          } else if (rowStatus.includes('ซ่อม')) {
            dbMap[id].borrowStatus = 'ส่งซ่อม';
          } else if (rowStatus.includes('สละ')) {
            dbMap[id].borrowStatus = 'สละสิทธิ์';
          } else if (rowStatus === 'ยังไม่ยืม') {
            dbMap[id].borrowStatus = 'ยังไม่ยืม';
          }
        }
      }
    }

    // --- C. รวมข้อมูล (Merge + Fuzzy Match) ---
    const processPerson = (type, no, id, name, room, source) => {
      if (!name) return;
      id = id.toString();
      let cleanedName = normalizeName(name);
      
      let finalBorrow = 'ยังไม่ยืม';
      let finalDoc = 'ยังไม่ส่ง';
      let finalSerial = '-';
      let finalFiles = {};
      let isInAssetSheet = false;

      // 1. Exact Match — เร็วสุด
      if (assetMap[cleanedName]) { 
        finalBorrow = assetMap[cleanedName].status; 
        finalSerial = assetMap[cleanedName].serial;
        isInAssetSheet = true;
      } else {
        // 2. Fuzzy Match — ── FIX #2: กรองความยาวก่อน ลด loop ~70% ──
        const lenA = cleanedName.length;
        const allowedErrors = lenA > 5 ? 2 : 1;

        for (let i = 0; i < assetKeys.length; i++) {
          let assetKey = assetKeys[i];

          // ✅ ถ้าความยาวต่างกันเกิน threshold ข้ามเลย ไม่ต้องคำนวณ
          if (Math.abs(assetKey.length - lenA) > allowedErrors) continue;

          let dist = getEditDistance(cleanedName, assetKey);
          if (dist <= allowedErrors) {
            finalBorrow = assetMap[assetKey].status;
            finalSerial = assetMap[assetKey].serial;
            isInAssetSheet = true;
            break;
          }
        }
      }

      if (dbMap[id]) {
        if (dbMap[id].borrowStatus !== 'ยังไม่ยืม') {
          finalBorrow = dbMap[id].borrowStatus;
        } else if (finalBorrow === 'ยังไม่ยืม' && isInAssetSheet) {
          finalBorrow = 'ยืมอยู่'; 
        }
        finalDoc    = dbMap[id].docStatus;
        if (dbMap[id].serial !== '-') finalSerial = dbMap[id].serial;
        finalFiles  = dbMap[id].files;
      }

      let finalReturnDocUrl = (dbMap[id] && dbMap[id].returnDocUrl) ? dbMap[id].returnDocUrl : '';

      allPeople.push({ 
        type: type, no: no, id: id, name: name, room: room, source_sheet: source, 
        serial: finalSerial, borrowStatus: finalBorrow, docStatus: finalDoc, 
        files: finalFiles, returnDocUrl: finalReturnDocUrl, inAsset: isInAssetSheet 
      });
    };

    // วนลูปอ่านรายชื่อนักเรียน
    SHEET_NAMES.STUDENTS.forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName);
      if (sheet) { 
        let lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          let data = sheet.getRange(1, 1, lastRow, 4).getValues(); 
          for (let i = 1; i < data.length; i++) { 
            processPerson('student', data[i][0], data[i][1], data[i][2], data[i][3], sheetName); 
          } 
        }
      }
    });

    // วนลูปอ่านรายชื่อครู
    let teacherSheet = ss.getSheetByName(SHEET_NAMES.TEACHERS);
    if (teacherSheet) { 
      let lastRow = teacherSheet.getLastRow();
      if (lastRow > 1) {
        let tData = teacherSheet.getRange(1, 1, lastRow, 2).getValues(); 
        for (let i = 1; i < tData.length; i++) { 
          processPerson('teacher', tData[i][0], 'T-'+tData[i][0], tData[i][1], 'ห้องพักครู', SHEET_NAMES.TEACHERS); 
        } 
      }
    }

    // ── FIX #1 + #6: Cache 600 วินาที + รองรับ chunk เมื่อ >90KB ──
    try {
      const serialized = JSON.stringify(allPeople);
      const CHUNK_SIZE = 85000; // 85KB ต่อ chunk (เผื่อ overhead)

      if (serialized.length <= CHUNK_SIZE) {
        // เก็บ key เดียว (เดิม 60 → ปรับเป็น 600 วินาที)
        cache.put(cacheKey, serialized, 600);
      } else {
        // แบ่ง chunk
        const totalChunks = Math.ceil(serialized.length / CHUNK_SIZE);
        for (let c = 0; c < totalChunks; c++) {
          cache.put(
            cacheKey + '_' + (c + 1),
            serialized.slice(c * CHUNK_SIZE, (c + 1) * CHUNK_SIZE),
            600
          );
        }
        cache.put(cacheKey + '_meta', JSON.stringify({ chunks: totalChunks }), 600);
      }
    } catch (_) { /* cache เต็มหรือ error ไม่กระทบผลลัพธ์ */ }

    return allPeople;

  } catch (e) {
    invalidateSystemDataCache();
    throw new Error("ไม่สามารถโหลดข้อมูลได้: " + e.toString());
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// ==========================================
// 4. ฟังก์ชันสำหรับระบบจับคู่ชื่อ (AUDIT & FIX)
// ==========================================

function getAssetAuditData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let peopleList = [];
  
  const fetchPeople = (sheetNames, isTeacherSheet) => {
    if(!Array.isArray(sheetNames)) sheetNames = [sheetNames];
    sheetNames.forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName);
      if(sheet && sheet.getLastRow() > 1) {
        let lastR = Math.max(2, sheet.getLastRow() - 1);
        let cols = isTeacherSheet ? 2 : 4;
        let data = sheet.getRange(2, 1, lastR, cols).getValues();
        data.forEach(r => {
          let originalName = isTeacherSheet ? r[1] : r[2];
          if(originalName) { 
            let norm = normalizeName(originalName);
            if (norm) {
              peopleList.push({
                id: isTeacherSheet ? ('T-'+r[0]) : r[1], 
                name: originalName, 
                room: isTeacherSheet ? 'ห้องพักครู' : (r[3]||''), 
                sheet: sheetName, 
                norm: norm
              });
            }
          }
        });
      }
    });
  };

  fetchPeople(SHEET_NAMES.STUDENTS, false);
  fetchPeople(SHEET_NAMES.TEACHERS, true);

  let peopleNormMap = new Map();
  peopleList.forEach(p => peopleNormMap.set(p.norm, p.name));

  let assetSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  let orphans = [];

  if(assetSheet && assetSheet.getLastRow() > 1) {
    let assetData = assetSheet.getRange(2, 1, assetSheet.getLastRow() - 1, 10).getValues(); 
    
    assetData.forEach((r, index) => {
      let assetName = r[4];
      let serial = r[2] || r[3] || 'ไม่มี Serial'; 
      
      if(assetName) {
        let assetNorm = normalizeName(assetName);
        if(!assetNorm) return;
        if(peopleNormMap.has(assetNorm)) return;

        let suggestions = [];
        for(let p of peopleList) {
          let dist = getEditDistance(assetNorm, p.norm);
          let isPartial = assetNorm.includes(p.norm) || p.norm.includes(assetNorm);
          let threshold = Math.ceil(Math.max(assetNorm.length, p.norm.length) * 0.3);
          if(isPartial || dist <= threshold) {
            suggestions.push({
              name: p.name, sheet: p.sheet, room: p.room || '-',
              diff: isPartial ? 0 : dist, isPartial: isPartial
            });
          }
        }
        suggestions.sort((a,b) => (a.isPartial === b.isPartial) ? a.diff - b.diff : (a.isPartial ? -1 : 1));
        orphans.push({
          row: index + 2, assetName: assetName,
          serial: serial.toString(), suggestions: suggestions.slice(0, 5) 
        });
      }
    });
  }
  return orphans;
}

function adminFixAssetName(data) {
  if (!data || data.oldAssetName == null || data.correctName == null) {
    return { success: false, message: "ข้อมูลไม่ครบถ้วน" };
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงานรายงานทะเบียนทรัพย์สิน" };
  
  try {
    const rows = sheet.getDataRange().getDisplayValues();
    let rowToUpdate = -1;
    for(let i=0; i<rows.length; i++) {
      let currentName   = rows[i][4];
      let currentSerial = rows[i][2] || rows[i][3];
      if(currentName == data.oldAssetName && String(currentSerial) == String(data.serial)) {
        rowToUpdate = i + 1;
        break;
      }
    }
    if(rowToUpdate > -1) {
      sheet.getRange(rowToUpdate, 5).setValue(data.correctName);
      invalidateSystemDataCache();
      // ── FIX #5: ไม่เรียก syncAllNamesToSheet ทันที (ให้ Trigger จัดการ) ──
      return { success: true, message: "แก้ไขชื่อในทะเบียนทรัพย์สินเรียบร้อยแล้ว" };
    } else {
      return { success: false, message: "ไม่พบข้อมูลเดิมในทะเบียนทรัพย์สิน (อาจถูกแก้ไขไปแล้ว)" };
    }
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function adminFixAssetNameBulk(updateList) {
  if (!updateList || !Array.isArray(updateList) || updateList.length === 0) {
    return { success: false, message: "ไม่มีรายการที่จะแก้ไข" };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงานรายงานทะเบียนทรัพย์สิน" };

  const lock = LockService.getScriptLock();
  let lockAcquired = false;

  try {
    lockAcquired = lock.tryLock(10000);
    if (!lockAcquired) {
      return { success: false, message: "ระบบกำลังดำเนินการอื่นอยู่ กรุณาลองใหม่อีกครั้ง" };
    }

    const rows = sheet.getDataRange().getDisplayValues();
    // ── FIX #4: สร้าง lookup map แทนการวนลูปซ้อน แล้ว batch write ──
    const pendingUpdates = {}; // { rowNumber(1-based): newName }

    updateList.forEach(item => {
      if (!item || item.oldAssetName == null || item.correctName == null) return;
      const targetName   = String(item.oldAssetName).trim();
      const targetSerial = String(item.serial || '').trim();

      for (let i = 1; i < rows.length; i++) {
        const currentName   = String(rows[i][4] || '').trim();
        const currentSerial = String(rows[i][2] || rows[i][3] || '').trim();
        if (currentName === targetName && currentSerial === targetSerial) {
          pendingUpdates[i + 1] = item.correctName;
          rows[i][4] = item.correctName; // อัปเดต local copy ป้องกัน duplicate match
          break;
        }
      }
    });

    // เขียนลง Sheet — ทีละ row แต่รวมไว้ใน batch loop เดียว
    const updateRows = Object.keys(pendingUpdates);
    updateRows.forEach(rowNum => {
      sheet.getRange(Number(rowNum), 5).setValue(pendingUpdates[rowNum]);
    });

    SpreadsheetApp.flush();
    invalidateSystemDataCache();
    // ── FIX #5: ไม่เรียก syncAllNamesToSheet ทันที (ให้ Trigger จัดการ) ──
    return { success: true, count: updateRows.length };

  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() };
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

// ==========================================
// 5. STANDARD HELPERS (Form, Auth, etc.)
// ==========================================

/**
 * แปลง base64 เป็น Blob สำหรับอัปโหลดไฟล์ผ่าน google.script.run
 */
function base64ToBlob(base64Data, fileName) {
  if (!base64Data || typeof base64Data !== 'string') return null;
  try {
    let base64  = base64Data;
    let mimeType = 'application/octet-stream';
    if (base64Data.indexOf('base64,') >= 0) {
      const parts = base64Data.split(';base64,');
      mimeType = parts[0].replace('data:', '') || mimeType;
      base64   = parts[1] || base64;
    }
    const bytes = Utilities.base64Decode(base64);
    return Utilities.newBlob(bytes, mimeType, fileName);
  } catch (e) {
    return null;
  }
}

/**
 * อัปโหลดเอกสารจากหน้าแรก (การยืม หรือ การคืน) - นักเรียน/ครู
 */
function processUploadDoc(obj) {
  if (!obj || !obj.userId || !obj.userName) {
    return { success: false, message: "ข้อมูลไม่ครบถ้วน" };
  }
  const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetData   = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  const folder      = DriveApp.getFolderById(FOLDER_ID);
  const folderReturn = DriveApp.getFolderById(FOLDER_RETURN_ID);
  
  try {
    const timestamp = new Date();
    let url = "";
    let action       = "USER_UPDATE";
    let statusToSave = "ยืมอยู่";
    let note         = obj.note || "";
    
    if (obj.uploadType === 'borrow') {
      if (!obj.file_b64) return { success: false, message: "กรุณาเลือกไฟล์ PDF" };
      const fullName = "AGREEMENT_" + obj.userName + "_" + timestamp.getTime();
      const blob = base64ToBlob(obj.file_b64, fullName);
      if (!blob) return { success: false, message: "ไม่สามารถอ่านไฟล์ได้" };
      url          = folder.createFile(blob).setName(fullName).getUrl();
      statusToSave = "ยืมอยู่ | รอตรวจสอบเอกสาร";
    } else {
      if (!obj.file_b64) return { success: false, message: "กรุณาเลือกไฟล์ PDF" };
      const fullName = "RETURN_ห้อง" + (obj.userRoom || "") + "_" + obj.userName + "_" + timestamp.getTime();
      const blob = base64ToBlob(obj.file_b64, fullName);
      if (!blob) return { success: false, message: "ไม่สามารถอ่านไฟล์ได้" };
      url          = folderReturn.createFile(blob).setName(fullName).getUrl();
      action       = "USER_RETURN";
      statusToSave = "อยู่ระหว่างการส่งคืน";
      note        += (note ? "\n" : "") + "[หลักฐานการคืน]: " + url;
    }
    
    sheetData.appendRow([
      timestamp, obj.userId, obj.userName, obj.userType, obj.userRoom || "", obj.userSerial || "",
      action, note, statusToSave,
      obj.uploadType === 'borrow' ? url : "", "", "", "", ""
    ]);

    // อัปเดตสถานะลงชีตรายชื่อหลัก
    if (obj.source_sheet) {
      const targetSheet = ss.getSheetByName(obj.source_sheet);
      if (targetSheet) {
        const sheetValues = targetSheet.getDataRange().getValues();
        const isTeacher   = obj.source_sheet === SHEET_NAMES.TEACHERS;
        for (let i = 1; i < sheetValues.length; i++) {
          const rowId = isTeacher ? ('T-' + String(sheetValues[i][0])) : String(sheetValues[i][1]);
          if (rowId === String(obj.userId)) {
            const newBorrow = obj.uploadType === 'return' ? 'อยู่ระหว่างการส่งคืน' : 'ยืมอยู่';
            const newDoc    = obj.uploadType === 'borrow' ? 'รอตรวจสอบ' : (sheetValues[i][11] || 'รอตรวจสอบ');
            targetSheet.getRange(i + 1, 11).setValue(newBorrow);
            if (obj.uploadType === 'borrow') targetSheet.getRange(i + 1, 12).setValue(newDoc);
            break;
          }
        }
      }
    }

    invalidateSystemDataCache();
    // ── FIX #5: ลบ syncAllNamesToSheet ออก ให้ Time Trigger จัดการ ──
    return { success: true, message: "อัปโหลดเอกสารสำเร็จ" };
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() };
  }
}

function processForm(formObject) {
  if (!formObject || !formObject.userId || !formObject.userName) {
    return { success: false, message: "ข้อมูลไม่ครบถ้วน (userId, userName)" };
  }

  const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetData = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  const folder    = DriveApp.getFolderById(FOLDER_ID);
  
  try {
    const timestamp = new Date();
    const uploadFromBase64 = (b64Data, prefix) => {
      if (!b64Data) return "";
      const fullName = prefix + "_" + formObject.userName + "_" + timestamp.getTime();
      const blob = base64ToBlob(b64Data, fullName);
      if (!blob) return "";
      return folder.createFile(blob).setName(fullName).getUrl();
    };

    let url_agreement = uploadFromBase64(formObject.file_agreement_b64, "AGREEMENT");
    let url_return    = "";
    const folderReturn = DriveApp.getFolderById(FOLDER_RETURN_ID);
    
    if (formObject.file_return_b64) {
      const fullName = "RETURN_ห้อง" + (formObject.userRoom || "") + "_" + formObject.userName + "_" + timestamp.getTime();
      const blob = base64ToBlob(formObject.file_return_b64, fullName);
      if (blob) url_return = folderReturn.createFile(blob).setName(fullName).getUrl();
    }

    let statusToSave = formObject.statusSelect || 'ยืมอยู่';
    let action       = "USER_UPDATE";
    let note         = formObject.note || "";
    if (url_agreement !== "") statusToSave = statusToSave + " | รอตรวจสอบเอกสาร";
    if (url_return !== "") {
      statusToSave = "อยู่ระหว่างการส่งคืน";
      action       = "USER_RETURN";
      note        += (note ? "\n" : "") + "[หลักฐานการคืน]: " + url_return;
    }

    sheetData.appendRow([
      timestamp, formObject.userId, formObject.userName, formObject.userType,
      formObject.userRoom, formObject.userSerial, action, note, statusToSave,
      url_agreement, "", "", "", ""
    ]);
    invalidateSystemDataCache();
    // ── FIX #5: ลบ syncAllNamesToSheet ออก ──
    return { success: true, message: "บันทึกข้อมูลเรียบร้อย" };
  } catch (error) { 
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.toString() };
  }
}

function verifyAdmin(u, p) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ADMIN);
  if (!sheet) return { success: false, message: "No Admin Sheet" };
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(u).trim() && String(data[i][1]).trim() === String(p).trim()) {
      return { success: true, role: 'admin', name: data[i][2] ? String(data[i][2]).trim() : 'Admin' };
    }
  }
  return { success: false, message: "Login Failed" };
}

function verifyAdvisor(u, p) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ADVISOR);
  if (!sheet) return { success: false, message: "ไม่พบชีตครูที่ปรึกษา" };
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    let rowUser = String(data[i][0] || '').trim();
    let rowPass = String(data[i][1] || '').trim();
    if (rowUser === String(u).trim() && rowPass === String(p).trim()) {
      let level = data[i][2] || '';
      let room  = data[i][3] ? String(data[i][3]).trim() : '';
      let name  = data[i][4] ? String(data[i][4]).trim() : 'คุณครูที่ปรึกษา';
      return { success: true, role: 'advisor', level: level, room: room, name: name };
    }
  }
  return { success: false, message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" };
}

function adminUpdateData(data) {
  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetLog = ss.getSheetByName(SHEET_NAMES.DATA_DB); 
  
  try {
    const timestamp = new Date();
    let editor = data.editorRole === 'advisor' 
      ? ("ครูที่ปรึกษา: " + data.editorName) 
      : ("ADMIN: " + (data.editorName || "ผู้ดูแลระบบ"));
    
    if (sheetLog) {
      if (data.borrowStatusSelect) {
        sheetLog.appendRow([
          timestamp, data.userId, data.userName, data.userType, data.userRoom, data.userSerial,
          editor, data.note || "-", data.borrowStatusSelect, "", "", "", "", ""
        ]);
      }
      if (data.docStatusSelect && data.docStatusSelect !== "") {
        sheetLog.appendRow([
          timestamp, data.userId, data.userName, data.userType, data.userRoom, data.userSerial,
          editor + "|DOC", data.note || "-", data.docStatusSelect, "", "", "", "", ""
        ]);
      }
    }
    
    const targetSheetName = data.source_sheet; 
    if (targetSheetName) {
      const targetSheet = ss.getSheetByName(targetSheetName);
      if (targetSheet) {
        const sheetValues  = targetSheet.getDataRange().getValues();
        const isTeacherSheet = targetSheetName === SHEET_NAMES.TEACHERS;
        for (let i = 1; i < sheetValues.length; i++) {
          const rowMatches = isTeacherSheet
            ? ("T-" + String(sheetValues[i][0])) === String(data.userId)
            : String(sheetValues[i][1]) === String(data.userId);
          if (rowMatches) {
            if (data.note !== undefined) targetSheet.getRange(i + 1, 8).setValue(data.note);
            if (data.borrowStatusSelect)  targetSheet.getRange(i + 1, 11).setValue(data.borrowStatusSelect);
            if (data.docStatusSelect && data.docStatusSelect !== "") targetSheet.getRange(i + 1, 12).setValue(data.docStatusSelect);
            break;
          }
        }
      } else {
        throw new Error("หาแผ่นงาน '" + targetSheetName + "' ไม่เจอในไฟล์นี้");
      }
    }

    invalidateSystemDataCache();
    // ── FIX #5: ลบ syncAllNamesToSheet ออก ──
    return { success: true, message: "อัปเดตข้อมูลและสถานะใน Sheet สำเร็จ" };
  } catch (e) { 
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() }; 
  }
}

function getAdminSummary() {
  try {
    const data = getAllSystemData();
    let total = 0, borrowed = 0, notBorrow = 0, pending = 0, returned = 0, repair = 0;
    data.forEach(function(p) {
      total++;
      const s = p.borrowStatus || '';
      if (s === 'ยืมอยู่' || s === 'อยู่ระหว่างการส่งคืน') borrowed++;
      else if (s === 'ยังไม่ยืม') notBorrow++;
      else if (s === 'คืนแล้ว') returned++;
      else if (s === 'ส่งซ่อม' || s === 'สละสิทธิ์') repair++;
      if ((p.docStatus || '') === 'รอตรวจสอบ') pending++;
    });
    return { success: true, total, borrowed, notBorrow, pending, returned, repair };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function advisorApproveDoc(data) {
  if (!data || !data.userId || !data.source_sheet) {
    return { success: false, message: "ข้อมูลไม่ครบถ้วน" };
  }
  const ss              = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetLog        = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  const targetSheetName = data.source_sheet;
  const targetSheet     = ss.getSheetByName(targetSheetName);
  
  if (!targetSheet) return { success: false, message: "ไม่พบแผ่นงาน" };
  
  try {
    const timestamp = new Date();
    const editor    = "ครูที่ปรึกษา: " + (data.editorName || "คุณครู");
    
    if (sheetLog) {
      sheetLog.appendRow([
        timestamp, data.userId, data.userName || "-", data.userType || "student", 
        data.userRoom || "-", data.userSerial || "-",
        editor, "อนุมัติเอกสาร", "เอกสารผ่าน",
        "", "", "", "", ""
      ]);
    }
    
    const sheetValues    = targetSheet.getDataRange().getValues();
    const isTeacherSheet = targetSheetName === SHEET_NAMES.TEACHERS;
    
    for (let i = 1; i < sheetValues.length; i++) {
      const rowMatches = isTeacherSheet
        ? ("T-" + String(sheetValues[i][0])) === String(data.userId)
        : String(sheetValues[i][1]) === String(data.userId);
      if (rowMatches) {
        targetSheet.getRange(i + 1, 12).setValue("เอกสารผ่าน");
        break;
      }
    }
    
    invalidateSystemDataCache();
    // ── FIX #5: ลบ syncAllNamesToSheet ออก ──
    return { success: true, message: "อนุมัติเอกสารเรียบร้อย (สถานะเครื่องยังอยู่ระหว่างการส่งคืน จนกว่าแอดมินจะอนุมัติ)" };
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() };
  }
}

function adminDeleteUser(data) {
  if (data.editorRole === 'advisor') return { success: false, message: "ครูที่ปรึกษาไม่ได้รับอนุญาตให้ลบข้อมูล" };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.source_sheet);
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน" };
  try {
    const rows = sheet.getDataRange().getDisplayValues();
    let rowToDelete = -1;
    for (let i = 0; i < rows.length; i++) {
      if (data.source_sheet === SHEET_NAMES.TEACHERS) { 
        if (String(rows[i][1]).trim() === String(data.name).trim()) { rowToDelete = i + 1; break; } 
      } else { 
        if (String(rows[i][1]).trim() === String(data.id).trim()) { rowToDelete = i + 1; break; } 
      }
    }
    if (rowToDelete > -1) { 
      sheet.deleteRow(rowToDelete);
      invalidateSystemDataCache();
      // ── FIX #5: ลบ syncAllNamesToSheet ออก ──
      return { success: true, message: "ลบข้อมูลเรียบร้อยแล้ว" }; 
    } else { return { success: false, message: "ไม่พบข้อมูล" }; }
  } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
}

function adminAddUser(data) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.targetSheet); 
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน" };
  try {
    const nextNo = sheet.getLastRow(); 
    if (data.targetSheet === SHEET_NAMES.TEACHERS) {
      sheet.appendRow([nextNo, data.name]); 
    } else {
      sheet.appendRow([nextNo, data.id, data.name, data.room]);
    }
    invalidateSystemDataCache();
    // ── FIX #5: ลบ syncAllNamesToSheet ออก ──
    return { success: true, message: "เพิ่มรายชื่อเรียบร้อยแล้ว" };
  } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
}

// ==========================================
// 6. DASHBOARD STATS SYSTEM
// ==========================================

/**
 * ── FIX #3: ใช้ getAllSystemData() แทนการเปิด Sheet ซ้ำ ──
 * ประหยัด Quota และเร็วกว่ามากเพราะใช้ Cache
 */
function getDashboardStats() {
  try {
    const data = getAllSystemData();
    let borrowed = 0;
    for (let i = 0; i < data.length; i++) {
      const s = data[i].borrowStatus || '';
      if (s === 'ยืมอยู่' || s === 'อยู่ระหว่างการส่งคืน') borrowed++;
    }
    return { total: 2085, borrowed: borrowed, available: 2085 - borrowed };
  } catch(e) {
    return { total: 2085, borrowed: 0, available: 2085 };
  }
}

// ==========================================
// ฟังก์ชันสำหรับครูที่ปรึกษาทำเรื่องคืน iPad
// ==========================================
function processAdvisorReturn(formObject) {
  if (!formObject || !formObject.userId || !formObject.userName) {
    return { success: false, message: "ข้อมูลไม่ครบถ้วน (userId, userName)" };
  }

  const ss           = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetData    = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  const folderReturn = DriveApp.getFolderById(FOLDER_RETURN_ID);
  
  try {
    const timestamp = new Date();
    let url_return  = "";
    
    if (formObject.file_return_b64) {
      const fileName = "RETURN_ห้อง" + (formObject.userRoom || "") + "_" + formObject.userName + "_" + timestamp.getTime();
      const blob = base64ToBlob(formObject.file_return_b64, fileName);
      if (blob) url_return = folderReturn.createFile(blob).setName(fileName).getUrl();
    }
    
    let statusToSave = "อยู่ระหว่างการส่งคืน";
    let newNote = formObject.note || "";
    if (url_return !== "") {
      newNote += "\n[หลักฐานการคืน]: " + url_return;
    }

    sheetData.appendRow([
      timestamp, formObject.userId, formObject.userName, formObject.userType, 
      formObject.userRoom, formObject.userSerial, "ADVISOR_RETURN", newNote, statusToSave, 
      "", "", "", "", ""
    ]);

    if (formObject.source_sheet) {
      const targetSheet = ss.getSheetByName(formObject.source_sheet);
      if (targetSheet) {
        const sheetValues  = targetSheet.getDataRange().getValues();
        const isTeacher    = formObject.source_sheet === SHEET_NAMES.TEACHERS;
        for (let i = 1; i < sheetValues.length; i++) {
          const rowId = isTeacher ? ('T-' + String(sheetValues[i][0])) : String(sheetValues[i][1]);
          if (rowId === String(formObject.userId)) {
            targetSheet.getRange(i + 1, 11).setValue('อยู่ระหว่างการส่งคืน');
            break;
          }
        }
      }
    }

    invalidateSystemDataCache();
    // ── FIX #5: ลบ syncAllNamesToSheet ออก ──
    return { success: true, message: "ส่งเรื่องคืนเรียบร้อย รอแอดมินตรวจสอบ" };
  } catch (error) { 
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.toString() };
  }
}

// ==========================================
// 7. EXPORT & REPORT
// ==========================================

function getExportData(options) {
  try {
    let data = getAllSystemData();
    if (!options || options.scope !== 'filtered') return data;
    let filtered = data;
    if (options.levelFilter) {
      filtered = filtered.filter(function(r) { return String(r.source_sheet || '').indexOf(options.levelFilter) >= 0; });
    }
    if (options.roomFilter) {
      filtered = filtered.filter(function(r) { return String(r.room || '') === String(options.roomFilter); });
    }
    return filtered;
  } catch (e) {
    throw new Error("ไม่สามารถดึงข้อมูล Export ได้: " + e.toString());
  }
}

function getReportData() {
  try {
    var data  = getAllSystemData();
    var byKey = {};

    for (var i = 0; i < data.length; i++) {
      var p     = data[i];
      var level = p.source_sheet || '-';
      var room  = String(p.room || '-');
      var key   = level + '|' + room;

      if (!byKey[key]) {
        byKey[key] = { level: level, room: room, total: 0, borrowed: 0, docPassed: 0, notBorrowed: 0, returned: 0, repair: 0 };
      }

      byKey[key].total++;
      var status = p.borrowStatus || '';

      if (status === 'ยืมอยู่' || status === 'อยู่ระหว่างการส่งคืน') byKey[key].borrowed++;
      else if (status === 'ยังไม่ยืม') byKey[key].notBorrowed++;
      else if (status === 'คืนแล้ว') byKey[key].returned++;
      else if (status === 'ส่งซ่อม' || status === 'สละสิทธิ์') byKey[key].repair++;

      if ((p.docStatus || '') === 'เอกสารผ่าน') byKey[key].docPassed++;
    }

    return Object.keys(byKey).map(function(k) { return byKey[k]; }).sort(function(a, b) {
      if (a.level !== b.level) return String(a.level).localeCompare(b.level);
      return String(a.room).localeCompare(b.room);
    });
  } catch (e) {
    throw new Error("ไม่สามารถดึงข้อมูลรายงานได้: " + e.toString());
  }
}

// ==========================================
// 8. ซิงค์ชีต "รายชื่อทั้งหมด" (SYNC ALL NAMES SHEET)
// ==========================================

function syncAllNamesToSheet(data) {
  if (!data || !Array.isArray(data)) return;
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAMES.ALL_NAMES);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAMES.ALL_NAMES);
    
    const headers = ['No', 'รหัส', 'ชื่อ-นามสกุล', 'ประเภท', 'ระดับชั้น', 'ห้อง', 'Serial', 'สถานะเครื่อง', 'สถานะเอกสาร'];
    const rows = data.map(function(p, idx) {
      return [
        idx + 1, p.id || '', p.name || '',
        (p.type === 'teacher' ? 'ครู' : 'นักเรียน'),
        p.source_sheet || '', p.room || '',
        p.serial || '-', p.borrowStatus || 'ยังไม่ยืม', p.docStatus || 'ยังไม่ส่ง'
      ];
    });
    
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('syncAllNamesToSheet error: ' + e.toString());
  }
}

/**
 * อัปเดตข้อมูล + ซิงค์ชีตรายชื่อทั้งหมด (สำหรับปุ่ม "อัปเดต" และ Time Trigger)
 */
function refreshAndSyncData() {
  invalidateSystemDataCache();
  const data = getAllSystemData();
  syncAllNamesToSheet(data);
  return data;
}

// ==========================================
// 9. TIME-BASED TRIGGER SETUP
// ── FIX #5: สร้าง Trigger ให้ syncAllNamesToSheet ทำงานอัตโนมัติ
//    ทุก 5 นาที แทนการเรียกหลัง action ทุกครั้ง ──
// วิธีใช้: รัน createSyncTrigger() ครั้งเดียวจาก Apps Script Editor
// ==========================================

/**
 * สร้าง Time-Based Trigger สำหรับ sync ชีต "รายชื่อทั้งหมด" ทุก 5 นาที
 * รัน refreshAndSyncData ผ่าน Trigger แทนการเรียกหลัง action ทุกครั้ง
 */
function createSyncTrigger() {
  // ลบ Trigger เก่าออกก่อน (ป้องกัน duplicate)
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'refreshAndSyncData') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // สร้างใหม่ทุก 5 นาที
  ScriptApp.newTrigger('refreshAndSyncData')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('✅ Sync Trigger created: refreshAndSyncData ทุก 5 นาที');
}

/**
 * ลบ Trigger ทั้งหมด (ใช้เมื่อต้องการหยุด auto-sync)
 */
function deleteSyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'refreshAndSyncData') {
      ScriptApp.deleteTrigger(t);
    }
  });
  Logger.log('🗑️ Sync Trigger ถูกลบแล้ว');
}