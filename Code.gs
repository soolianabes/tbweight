var SHEET_ID = '1l5Pjw6OL0aJsGTKIo2SJtgyBzDyedXwFqFfURpKlgDM';
var SHEET_NAME = 'Students';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบบันทึกน้ำหนัก-ส่วนสูงออนไลน์')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['studentId', 'fullName', 'class', 'gender', 'age', 'weight', 'height']);
  }
  return sheet;
}

function fetchData() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    row.rowId = i + 1;
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    result.push(row);
  }
  return JSON.stringify(result);
}

function addStudent(data) {
  var sheet = getSheet();
  sheet.appendRow([
    data.studentId,
    data.fullName,
    data.class,
    data.gender,
    data.age,
    data.weight,
    data.height
  ]);
  return JSON.stringify({ status: 'success', message: 'บันทึกข้อมูลสำเร็จ' });
}

function updateStudent(data) {
  var sheet = getSheet();
  var row = parseInt(data.rowId);
  if (row < 2) return JSON.stringify({ status: 'error', message: 'ไม่พบข้อมูล' });
  sheet.getRange(row, 1).setValue(data.studentId);
  sheet.getRange(row, 2).setValue(data.fullName);
  sheet.getRange(row, 3).setValue(data.class);
  sheet.getRange(row, 4).setValue(data.gender);
  sheet.getRange(row, 5).setValue(data.age);
  sheet.getRange(row, 6).setValue(data.weight);
  sheet.getRange(row, 7).setValue(data.height);
  return JSON.stringify({ status: 'success', message: 'อัปเดตข้อมูลสำเร็จ' });
}

function deleteStudent(rowId) {
  var sheet = getSheet();
  var row = parseInt(rowId);
  if (row < 2) return JSON.stringify({ status: 'error', message: 'ไม่พบข้อมูล' });
  sheet.deleteRow(row);
  return JSON.stringify({ status: 'success', message: 'ลบข้อมูลสำเร็จ' });
}

function addMultipleStudents(studentsData) {
  var sheet = getSheet();
  var thaiToEn = {
    'รหัสนักเรียน': 'studentId',
    'ชื่อ-สกุล': 'fullName',
    'ชื่อ-นามสกุล': 'fullName',
    'ชื่อ สกุล': 'fullName',
    'ชั้น': 'class',
    'เพศ': 'gender',
    'อายุ': 'age',
    'น้ำหนัก': 'weight',
    'น้ำหนัก (kg)': 'weight',
    'น้ำหนัก(kg)': 'weight',
    'ส่วนสูง': 'height',
    'ส่วนสูง (cm)': 'height',
    'ส่วนสูง(cm)': 'height'
  };

  var count = 0;
  for (var i = 0; i < studentsData.length; i++) {
    var row = studentsData[i];
    var mapped = {};
    for (var key in row) {
      var engKey = thaiToEn[key] || key;
      mapped[engKey] = row[key];
    }
    if (mapped.studentId) {
      sheet.appendRow([
        mapped.studentId || '',
        mapped.fullName || '',
        mapped.class || '',
        mapped.gender || '',
        mapped.age || '',
        mapped.weight || '',
        mapped.height || ''
      ]);
      count++;
    }
  }
  return JSON.stringify({ status: 'success', message: 'นำเข้าข้อมูล ' + count + ' รายการสำเร็จ' });
}
