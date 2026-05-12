const SHEET_NAME = 'ผู้สูงอายุ';
const HEADERS = [
  'รหัส', 'เลขบัตรประชาชน', 'คำนำหน้า', 'ชื่อ', 'นามสกุล',
  'วันเกิด', 'อายุ', 'เพศ', 'เบอร์โทร',
  'บ้านเลขที่', 'หมู่', 'ตำบล', 'อำเภอ', 'จังหวัด',
  'ผู้ดูแล', 'เบอร์โทรผู้ดูแล', 'ความสัมพันธ์',
  'โรคประจำตัว', 'ยาที่ใช้', 'สิทธิการรักษา',
  'สวัสดิการที่ได้รับ', 'อาชีพ', 'สถานะ', 'หมายเหตุ',
  'โหมดวันเกิด', 'อายุประมาณการ', 'วันที่บันทึก', 'ผู้บันทึก'
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบบันทึกข้อมูลผู้สูงอายุ อบต.โคกเพลาะ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function initSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setBackground('#1a73e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    ensureHeaders(sheet);
  }
  return sheet;
}

function ensureHeaders(sheet) {
  const existingHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  if (existingHeaders.join('|') === HEADERS.join('|')) return;

  const missingHeaders = HEADERS.filter(h => !existingHeaders.includes(h));
  if (missingHeaders.length > 0) {
    const startCol = existingHeaders.length + 1;
    sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]);
  }

  const finalHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const isSameOrder = HEADERS.every((h, i) => finalHeaders[i] === h);
  if (!isSameOrder) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }

  sheet.getRange(1, 1, 1, HEADERS.length)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function getAllData() {
  try {
    const sheet = initSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    return data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        if (row[i] instanceof Date) {
          obj[h] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          obj[h] = row[i];
        }
      });
      return obj;
    });

  } catch(e) {
    console.error('getAllData error:', e);
    return [];
  }
}

function saveData(formData) {
  try {
    const sheet = initSheet();
    const now = new Date();
    const headerIndex = getHeaderIndexMap(sheet);
    const parsedBirth = parseBirthInput(formData, now);

    if (formData['รหัส']) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][headerIndex['รหัส']] == formData['รหัส']) {
          const rowData = buildRow(
            formData,
            parsedBirth,
            formData['รหัส'],
            data[i][headerIndex['วันที่บันทึก']],
            data[i][headerIndex['ผู้บันทึก']]
          );
          sheet.getRange(i + 1, 1, 1, HEADERS.length).setValues([rowData]);
          return { success: true, message: 'แก้ไขข้อมูลเรียบร้อย' };
        }
      }
    } else {
      const lastRow = sheet.getLastRow();
      const newId = lastRow > 1 ? (sheet.getRange(lastRow, 1).getValue() + 1) : 1;
      const user = Session.getActiveUser().getEmail() || 'ไม่ระบุ';
      const rowData = buildRow(formData, parsedBirth, newId, now, user);
      sheet.appendRow(rowData);
      return { success: true, message: 'บันทึกข้อมูลเรียบร้อย' };
    }

  } catch(e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

function getHeaderIndexMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const map = {};
  headers.forEach((h, i) => map[h] = i);
  return map;
}

function parseBirthInput(formData, now) {
  const mode = formData['โหมดวันเกิด'] || 'AD';
  let birthDate = '';
  let age = 0;
  let approxAge = '';

  if (mode === 'UNKNOWN') {
    age = parseInt(formData['อายุประมาณการ'], 10);
    if (isNaN(age) || age < 0) throw new Error('กรุณากรอกอายุประมาณการให้ถูกต้อง');
    approxAge = age;
  } else {
    const source = mode === 'BE' ? formData['วันเกิด(พ.ศ.)'] : formData['วันเกิด'];
    const parts = String(source || '').split('-');
    if (parts.length !== 3) throw new Error('รูปแบบวันเกิดไม่ถูกต้อง');

    let year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);

    if ([year, month, day].some(v => isNaN(v))) throw new Error('รูปแบบวันเกิดไม่ถูกต้อง');
    if (mode === 'BE' || year > 2400) year -= 543;

    const dob = new Date(year, month, day);
    if (isNaN(dob.getTime())) throw new Error('ไม่สามารถแปลงวันเกิดได้');

    birthDate = Utilities.formatDate(dob, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    age = Math.floor((now - dob) / (365.25 * 24 * 60 * 60 * 1000));
  }

  return { mode, birthDate, age, approxAge };
}

function buildRow(f, birth, id, dateAdded, addedBy) {
  const rowObj = {
    'รหัส': id,
    'เลขบัตรประชาชน': "'" + String(f['เลขบัตรประชาชน'] || '').trim(),
    'คำนำหน้า': f['คำนำหน้า'],
    'ชื่อ': f['ชื่อ'],
    'นามสกุล': f['นามสกุล'],
    'วันเกิด': birth.birthDate,
    'อายุ': birth.age,
    'เพศ': f['เพศ'],
    'เบอร์โทร': "'" + String(f['เบอร์โทร'] || '').trim(),
    'บ้านเลขที่': f['บ้านเลขที่'],
    'หมู่': f['หมู่'],
    'ตำบล': f['ตำบล'],
    'อำเภอ': f['อำเภอ'],
    'จังหวัด': f['จังหวัด'],
    'ผู้ดูแล': f['ผู้ดูแล'],
    'เบอร์โทรผู้ดูแล': "'" + String(f['เบอร์โทรผู้ดูแล'] || '').trim(),
    'ความสัมพันธ์': f['ความสัมพันธ์'],
    'โรคประจำตัว': f['โรคประจำตัว'],
    'ยาที่ใช้': f['ยาที่ใช้'],
    'สิทธิการรักษา': f['สิทธิการรักษา'],
    'สวัสดิการที่ได้รับ': f['สวัสดิการที่ได้รับ'],
    'อาชีพ': f['อาชีพ'],
    'สถานะ': f['สถานะ'],
    'หมายเหตุ': f['หมายเหตุ'],
    'โหมดวันเกิด': birth.mode,
    'อายุประมาณการ': birth.approxAge,
    'วันที่บันทึก': dateAdded,
    'ผู้บันทึก': addedBy
  };

  return HEADERS.map(h => rowObj[h] !== undefined ? rowObj[h] : '');
}

function deleteData(id) {
  try {
    const sheet = initSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบข้อมูลเรียบร้อย' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูล' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getStats() {
  try {
    const data = getAllData();
    const total = data.length;
    const male = data.filter(r => r['เพศ'] === 'ชาย').length;
    const female = data.filter(r => r['เพศ'] === 'หญิง').length;

    const ageGroups = { '60-69': 0, '70-79': 0, '80-89': 0, '90+': 0 };
    data.forEach(r => {
      const age = parseInt(r['อายุ']);
      if (age >= 90) ageGroups['90+']++;
      else if (age >= 80) ageGroups['80-89']++;
      else if (age >= 70) ageGroups['70-79']++;
      else ageGroups['60-69']++;
    });

    const diseases = {};
    data.forEach(r => {
      (r['โรคประจำตัว'] || '').split(',').forEach(d => {
        d = d.trim();
        if (d) diseases[d] = (diseases[d] || 0) + 1;
      });
    });
    const topDiseases = Object.entries(diseases)
      .sort((a, b) => b[1] - a[1]).slice(0, 5);

    const welfare = {};
    data.forEach(r => {
      const w = r['สวัสดิการที่ได้รับ'] || 'ไม่ระบุ';
      welfare[w] = (welfare[w] || 0) + 1;
    });

    return { total, male, female, ageGroups, topDiseases, welfare };

  } catch(e) {
    console.error('getStats error:', e);
    return { total: 0, male: 0, female: 0, ageGroups: {}, topDiseases: [], welfare: {} };
  }
}

function exportToExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
    '/export?format=xlsx&sheet=' + encodeURIComponent(SHEET_NAME);
  return url;
}

function exportPersonPDF(id) {
  const data = getAllData();
  const person = data.find(r => r['รหัส'] == id);
  if (!person) return null;
  return generatePersonHTML(person);
}

function generatePersonHTML(p) {
  return '<!DOCTYPE html><html lang="th"><head>' +
    '<meta charset="UTF-8">' +
    '<link href="https://fonts.googleapis.com/css2?family=Sarabun&display=swap" rel="stylesheet">' +
    '<style>' +
    'body{font-family:"Sarabun",sans-serif;padding:20px;font-size:14px}' +
    'h2{text-align:center;color:#1a73e8}' +
    'table{width:100%;border-collapse:collapse}' +
    'td{padding:6px 10px;border:1px solid #ccc}' +
    '.label{background:#f0f4ff;font-weight:bold;width:35%}' +
    '.section{background:#1a73e8;color:white;font-weight:bold;padding:8px}' +
    '</style></head><body>' +
    '<h2>🏛️ ทะเบียนผู้สูงอายุ</h2>' +
    '<p style="text-align:center">องค์การบริหารส่วนตำบลโคกเพลาะ</p>' +
    '<table>' +
    '<tr><td colspan="2" class="section">ข้อมูลส่วนตัว</td></tr>' +
    '<tr><td class="label">เลขบัตรประชาชน</td><td>' + p['เลขบัตรประชาชน'] + '</td></tr>' +
    '<tr><td class="label">ชื่อ-นามสกุล</td><td>' + p['คำนำหน้า'] + p['ชื่อ'] + ' ' + p['นามสกุล'] + '</td></tr>' +
    '<tr><td class="label">วันเกิด</td><td>' + p['วันเกิด'] + '</td></tr>' +
    '<tr><td class="label">อายุ</td><td>' + p['อายุ'] + ' ปี</td></tr>' +
    '<tr><td class="label">เพศ</td><td>' + p['เพศ'] + '</td></tr>' +
    '<tr><td class="label">เบอร์โทร</td><td>' + p['เบอร์โทร'] + '</td></tr>' +
    '<tr><td colspan="2" class="section">ที่อยู่</td></tr>' +
    '<tr><td class="label">ที่อยู่</td><td>' + p['บ้านเลขที่'] + ' หมู่ ' + p['หมู่'] + ' ต.' + p['ตำบล'] + ' อ.' + p['อำเภอ'] + ' จ.' + p['จังหวัด'] + '</td></tr>' +
    '<tr><td colspan="2" class="section">ผู้ดูแล</td></tr>' +
    '<tr><td class="label">ชื่อผู้ดูแล</td><td>' + p['ผู้ดูแล'] + ' (' + p['ความสัมพันธ์'] + ')</td></tr>' +
    '<tr><td class="label">เบอร์โทรผู้ดูแล</td><td>' + p['เบอร์โทรผู้ดูแล'] + '</td></tr>' +
    '<tr><td colspan="2" class="section">ข้อมูลสุขภาพ</td></tr>' +
    '<tr><td class="label">โรคประจำตัว</td><td>' + p['โรคประจำตัว'] + '</td></tr>' +
    '<tr><td class="label">ยาที่ใช้</td><td>' + p['ยาที่ใช้'] + '</td></tr>' +
    '<tr><td class="label">สิทธิการรักษา</td><td>' + p['สิทธิการรักษา'] + '</td></tr>' +
    '<tr><td colspan="2" class="section">สวัสดิการ</td></tr>' +
    '<tr><td class="label">สวัสดิการที่ได้รับ</td><td>' + p['สวัสดิการที่ได้รับ'] + '</td></tr>' +
    '<tr><td class="label">อาชีพ</td><td>' + p['อาชีพ'] + '</td></tr>' +
    '<tr><td class="label">สถานะ</td><td>' + p['สถานะ'] + '</td></tr>' +
    '<tr><td class="label">หมายเหตุ</td><td>' + p['หมายเหตุ'] + '</td></tr>' +
    '</table>' +
    '<p style="text-align:right;margin-top:20px;font-size:12px">บันทึกเมื่อ: ' + p['วันที่บันทึก'] + ' | โดย: ' + p['ผู้บันทึก'] + '</p>' +
    '</body></html>';
}

function filterData(filters) {
  try {
    const data = getAllData(); // ดึงข้อมูลทั้งหมดที่แปลงเป็น Object แล้วมาใช้
    if (!data || data.length === 0) return [];

    const { searchText, gender, ageRange, status } = filters;

    return data.filter(r => {
      // 1. กรองจากคำค้นหา (ชื่อ, นามสกุล หรือ เลขบัตรประชาชน)
      const nameMatch = !searchText || 
                        String(r['ชื่อ']).includes(searchText) || 
                        String(r['นามสกุล']).includes(searchText) || 
                        String(r['เลขบัตรประชาชน']).includes(searchText);

      // 2. กรองจากเพศ
      const genderMatch = !gender || r['เพศ'] === gender;

      // 3. กรองจากสถานะ
      const statusMatch = !status || r['สถานะ'] === status;

      // 4. กรองจากช่วงอายุ
      let ageMatch = true;
      if (ageRange) {
        const age = parseInt(r['อายุ']);
        if (ageRange === "60-69") ageMatch = (age >= 60 && age <= 69);
        else if (ageRange === "70-79") ageMatch = (age >= 70 && age <= 79);
        else if (ageRange === "80-89") ageMatch = (age >= 80 && age <= 89);
        else if (ageRange === "90+") ageMatch = (age >= 90);
      }

      return nameMatch && genderMatch && statusMatch && ageMatch;
    });

  } catch (e) {
    console.error('filterData error:', e);
    return [];
  }
}
