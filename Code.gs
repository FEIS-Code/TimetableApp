// ============================================================
// Timetable App — Google Apps Script Backend
// Five Elements International School
// Team Phoenix The Changer
// ============================================================

const SPREADSHEET_ID = '1l6Kpp5sATMqQihsiMySolk1jIRIrqIn6LHvEQUov53A';
const TIMETABLE_SHEET = 'Timetable';
const USERS_SHEET = 'Users';
const CONFIG_SHEET = 'Config';
const TEACHERS_SHEET = 'Teachers';

function getSheet(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  // Try case-insensitive match
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toLowerCase() === name.trim().toLowerCase()) {
      return sheets[i];
    }
  }
  return null;
}

// --- Web App Entry Points ---

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'all';
  var result;
  switch (action) {
    case 'classes':
      result = getClasses();
      break;
    case 'timetable':
      result = getTimetable(e.parameter.grade, e.parameter.section);
      break;
    case 'teachertt':
      result = getTeacherTimetable(e.parameter.teacher);
      break;
    case 'config':
      result = getConfig();
      break;
    case 'teachers':
      result = getTeachers();
      break;
    case 'users':
      result = getUsers(e.parameter.u, e.parameter.p);
      break;
    case 'all':
      result = getAllTimetables();
      break;
    default:
      result = { error: 'Unknown action' };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var data;
  if (e.postData && e.postData.type === 'application/x-www-form-urlencoded') {
    data = {};
    var params = e.parameter;
    for (var key in params) data[key] = params[key];
  } else {
    data = JSON.parse(e.postData.contents);
  }

  var result;
  switch (data.action) {
    case 'login':
      result = login(data.username, data.password);
      break;
    case 'saveTimetable':
      result = saveTimetable(data);
      break;
    case 'setupData':
      var authCheck = login(data.username, data.password);
      if (!authCheck.success || authCheck.role !== 'admin') {
        result = { success: false, message: 'Unauthorized' };
      } else {
        setupData();
        result = { success: true, message: 'Setup complete' };
      }
      break;
    case 'addClass':
      result = addClassTimetable(data);
      break;
    case 'saveConfig':
      result = saveConfig(data);
      break;
    case 'saveTeachers':
      result = saveTeachers(data);
      break;
    case 'saveUsers':
      result = saveUsers(data);
      break;
    default:
      result = { error: 'Unknown action' };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// --- Data Functions ---

function getAllTimetables() {
  var sheet = getSheet(TIMETABLE_SHEET);
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    rows.push(obj);
  }
  return rows;
}

function getClasses() {
  var all = getAllTimetables();
  var classMap = {};
  for (var i = 0; i < all.length; i++) {
    var key = all[i].Grade + '-' + all[i].Section;
    if (!classMap[key]) {
      classMap[key] = { grade: all[i].Grade, section: all[i].Section };
    }
  }
  var classes = [];
  for (var k in classMap) classes.push(classMap[k]);
  classes.sort(function(a, b) {
    return (parseInt(a.grade) || 0) - (parseInt(b.grade) || 0) || a.section.localeCompare(b.section);
  });
  return classes;
}

function getTimetable(grade, section) {
  var all = getAllTimetables();
  return all.filter(function(r) {
    return String(r.Grade) === String(grade) && String(r.Section) === String(section);
  });
}

function getTeacherTimetable(teacherName) {
  var all = getAllTimetables();
  return all.filter(function(r) {
    return r.Teacher && r.Teacher.toLowerCase().indexOf(teacherName.toLowerCase()) !== -1;
  });
}

function login(username, password) {
  var sheet = getSheet(USERS_SHEET);
  if (!sheet) return { success: false, message: 'Users sheet not found' };
  var data = sheet.getDataRange().getValues();
  // Headers: Username, Password, Role, DisplayName
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(username).trim() &&
        String(data[i][1]).trim() === String(password).trim()) {
      return {
        success: true,
        role: String(data[i][2]).trim(),
        displayName: String(data[i][3]).trim(),
        username: String(data[i][0]).trim()
      };
    }
  }
  return { success: false, message: 'Invalid credentials' };
}

function saveTimetable(data) {
  // data: { auth: {username, password}, entries: [{Grade, Section, Day, PeriodNum, PeriodStart, PeriodEnd, Subject, Teacher, Room}] }
  var authResult = login(data.auth.username, data.auth.password);
  if (!authResult.success || authResult.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  var sheet = getSheet(TIMETABLE_SHEET);
  var existing = sheet.getDataRange().getValues();
  var headers = existing[0];

  var gradeCol = headers.indexOf('Grade');
  var sectionCol = headers.indexOf('Section');
  var dayCol = headers.indexOf('Day');
  var periodCol = headers.indexOf('PeriodNum');
  var subjectCol = headers.indexOf('Subject');
  var teacherCol = headers.indexOf('Teacher');
  var roomCol = headers.indexOf('Room');
  var pStartCol = headers.indexOf('PeriodStart');
  var pEndCol = headers.indexOf('PeriodEnd');

  var entries = data.entries;
  for (var e = 0; e < entries.length; e++) {
    var entry = entries[e];
    var found = false;
    for (var i = 1; i < existing.length; i++) {
      if (String(existing[i][gradeCol]) === String(entry.Grade) &&
          String(existing[i][sectionCol]) === String(entry.Section) &&
          String(existing[i][dayCol]) === String(entry.Day) &&
          String(existing[i][periodCol]) === String(entry.PeriodNum)) {
        // Update existing row
        sheet.getRange(i + 1, subjectCol + 1).setValue(entry.Subject);
        sheet.getRange(i + 1, teacherCol + 1).setValue(entry.Teacher || '');
        sheet.getRange(i + 1, roomCol + 1).setValue(entry.Room || '');
        // Refresh in-memory array too
        existing[i][subjectCol] = entry.Subject;
        existing[i][teacherCol] = entry.Teacher || '';
        existing[i][roomCol] = entry.Room || '';
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([entry.Grade, entry.Section, entry.Day, entry.PeriodNum, entry.PeriodStart, entry.PeriodEnd, entry.Subject, entry.Teacher || '', entry.Room || '']);
      // Refresh existing array
      existing = sheet.getDataRange().getValues();
    }
  }
  return { success: true, message: 'Timetable saved' };
}

function addClassTimetable(data) {
  var authResult = login(data.auth.username, data.auth.password);
  if (!authResult.success || authResult.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }
  var sheet = getSheet(TIMETABLE_SHEET);
  if (!sheet) return { success: false, message: 'Timetable sheet not found' };

  var grade = data.grade;
  var section = data.section;
  var isUpper = parseInt(grade) >= 6;
  var periods = isUpper
    ? [['0','8:10','8:30'],['00','8:30','9:00'],['1','9:00','9:40'],['2','9:40','10:20'],['3','10:20','11:00'],['4','11:00','11:40'],['5','11:40','12:20'],['L','12:20','1:00'],['6','1:00','1:40'],['7','1:40','2:20'],['8','2:20','3:00']]
    : [['0','8:20','8:40'],['00','8:40','9:00'],['1','9:00','9:40'],['2','9:40','10:20'],['3','10:20','11:00'],['4','11:00','11:40'],['5','11:40','12:20'],['L','12:20','1:00'],['6','1:00','1:40'],['7','1:40','2:20'],['8','2:20','3:00']];
  var days = ['MON','TUE','WED','THU','FRI','SAT'];
  var defaults = {'0':'ASSEMBLY/YOGA','00':'SHORT BREAK','L':'LUNCH BREAK'};

  var rows = [];
  for (var d = 0; d < days.length; d++) {
    for (var p = 0; p < periods.length; p++) {
      var pn = periods[p][0];
      rows.push([grade, section, days[d], pn, periods[p][1], periods[p][2], defaults[pn] || '', '', '']);
    }
  }
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 9).setValues(rows);
  return { success: true, message: 'Class ' + grade + section + ' added' };
}

// --- Config & Teachers ---

function getConfig() {
  var sheet = getSheet(CONFIG_SHEET);
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var config = {};
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) config[String(data[i][0]).trim()] = String(data[i][1] || '').trim();
  }
  return config;
}

function getTeachers() {
  var sheet = getSheet(TEACHERS_SHEET);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      name: String(data[i][0]).trim(),
      subjects: String(data[i][1] || '').trim(),
      classes: String(data[i][2] || '').trim()
    });
  }
  return result;
}

function getUsers(username, password) {
  var auth = login(username, password);
  if (!auth.success || auth.role !== 'admin') return { error: 'Unauthorized' };
  var sheet = getSheet(USERS_SHEET);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({ username: String(data[i][0]), password: String(data[i][1]), role: String(data[i][2]), displayName: String(data[i][3]) });
  }
  return result;
}

function saveConfig(data) {
  var authResult = login(data.auth.username, data.auth.password);
  if (!authResult.success || authResult.role !== 'admin') return { success: false, message: 'Unauthorized' };
  var sheet = getSheet(CONFIG_SHEET);
  if (!sheet) {
    sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(CONFIG_SHEET);
  }
  sheet.clear();
  var entries = data.config; // {key: value, ...}
  for (var key in entries) {
    sheet.appendRow([key, entries[key]]);
  }
  return { success: true };
}

function saveTeachers(data) {
  var authResult = login(data.auth.username, data.auth.password);
  if (!authResult.success || authResult.role !== 'admin') return { success: false, message: 'Unauthorized' };
  var sheet = getSheet(TEACHERS_SHEET);
  if (!sheet) {
    sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(TEACHERS_SHEET);
  }
  sheet.clear();
  sheet.appendRow(['Name', 'Subjects', 'Classes']);
  var teachers = data.teachers; // [{name, subjects, classes}]
  for (var i = 0; i < teachers.length; i++) {
    sheet.appendRow([teachers[i].name, teachers[i].subjects, teachers[i].classes]);
  }
  return { success: true };
}

function saveUsers(data) {
  var authResult = login(data.auth.username, data.auth.password);
  if (!authResult.success || authResult.role !== 'admin') return { success: false, message: 'Unauthorized' };
  var sheet = getSheet(USERS_SHEET);
  if (!sheet) return { success: false, message: 'Users sheet not found' };
  sheet.clear();
  sheet.appendRow(['Username', 'Password', 'Role', 'DisplayName']);
  var users = data.users;
  for (var i = 0; i < users.length; i++) {
    sheet.appendRow([users[i].username, users[i].password, users[i].role, users[i].displayName]);
  }
  return { success: true };
}

// ============================================================
// setupData() — Run ONCE to create sheets and pre-populate
// ============================================================

function setupData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // --- Create Users sheet ---
  var usersSheet = ss.getSheetByName(USERS_SHEET);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(USERS_SHEET);
  } else {
    usersSheet.clear();
  }
  usersSheet.appendRow(['Username', 'Password', 'Role', 'DisplayName']);
  usersSheet.appendRow(['admin', 'admin123', 'admin', 'Administrator']);
  usersSheet.appendRow(['maneesha', 'teach123', 'teacher', 'Ms. Maneesha']);
  usersSheet.appendRow(['madhuri', 'teach123', 'teacher', 'Ms. Madhuri']);
  usersSheet.appendRow(['akanksha', 'teach123', 'teacher', 'Ms. Akanksha']);
  usersSheet.appendRow(['rajsree', 'teach123', 'teacher', 'Ms. Rajsree']);
  usersSheet.appendRow(['sailaja', 'teach123', 'teacher', 'Ms. Sailaja']);
  usersSheet.appendRow(['saritha', 'teach123', 'teacher', 'Ms. Saritha Ambekar']);
  usersSheet.appendRow(['madhavi', 'teach123', 'teacher', 'Ms. Madhavi']);
  usersSheet.appendRow(['vidya', 'teach123', 'teacher', 'Ms. Vidya']);
  usersSheet.appendRow(['sridevi', 'teach123', 'teacher', 'Ms. Sridevi']);
  usersSheet.appendRow(['usha', 'teach123', 'teacher', 'Ms. Usha']);
  usersSheet.appendRow(['amrutha', 'teach123', 'teacher', 'Ms. Amrutha']);

  // --- Create Config sheet ---
  var configSheet = ss.getSheetByName(CONFIG_SHEET);
  if (!configSheet) { configSheet = ss.insertSheet(CONFIG_SHEET); } else { configSheet.clear(); }
  configSheet.appendRow(['Grades', '3,4,5,6,7,8,9']);
  configSheet.appendRow(['Sections', 'A,B']);
  configSheet.appendRow(['Subjects', 'MATH,ENG,SCI,BIO,PHY/CHEM,SOC,ICT,IT,2L,3L,HINDI,TELUGU,EVS,MUS,ART,DANCE,SPORTS,KALARI,SW/CY,HOBBY,LIB,AF,VE,ACTIVITY']);
  configSheet.appendRow(['PeriodCount', '11']);

  // --- Create Teachers sheet ---
  var teachersSheet = ss.getSheetByName(TEACHERS_SHEET);
  if (!teachersSheet) { teachersSheet = ss.insertSheet(TEACHERS_SHEET); } else { teachersSheet.clear(); }
  teachersSheet.appendRow(['Name', 'Subjects', 'Classes']);
  teachersSheet.appendRow(['Ms. Maneesha', 'BIO,SCI', '9A']);
  teachersSheet.appendRow(['Ms. Madhuri', 'ICT', '8A']);
  teachersSheet.appendRow(['Ms. Akanksha', 'SOC', '7A']);
  teachersSheet.appendRow(['Ms. Rajsree', 'ENG', '6A']);
  teachersSheet.appendRow(['Ms. Sailaja', 'ENG,MATH', '5A']);
  teachersSheet.appendRow(['Ms. Saritha Ambekar', 'ENG,MATH', '5B']);
  teachersSheet.appendRow(['Ms. Madhavi', 'ENG,MATH', '4A']);
  teachersSheet.appendRow(['Ms. Vidya', 'ENG,MATH', '4B']);
  teachersSheet.appendRow(['Ms. Sridevi', 'ENG,MATH', '3A']);
  teachersSheet.appendRow(['Ms. Usha', 'ENG,MATH', '3B']);
  teachersSheet.appendRow(['Ms. Amrutha', 'SCI', '6A,7A,8A,9A']);
  teachersSheet.appendRow(['Ms. Rakhi', 'HINDI,2L,3L', '6A,7A,8A,9A']);
  teachersSheet.appendRow(['Ms. Srilatha', 'IT', '3A,3B,4A,4B,5A,5B']);
  teachersSheet.appendRow(['Ms. Seema', 'EVS', '3A,3B,4A,4B,5A,5B']);
  teachersSheet.appendRow(['Veena', 'TELUGU', '3A,3B,4A,4B,5A,5B']);
  teachersSheet.appendRow(['Ms. Chandrika', 'MUS', '3A,3B,4A,4B,5A,5B,6A,7A,8A,9A']);
  teachersSheet.appendRow(['Ms. Venkatesh', 'DANCE', '3A,3B,4A,4B,5A,5B,6A,7A,8A,9A']);
  teachersSheet.appendRow(['Ms. Manasa', 'LIB', '3A,3B,4A,4B,5A']);
  teachersSheet.appendRow(['Ms. Nageshwar Rao', 'SCI,PHY/CHEM', '6A,7A,8A,9A']);
  teachersSheet.appendRow(['Ms. Sai Kumar', 'SPORTS,KALARI', '3A,3B,4A,4B,5A,5B,6A,7A,8A,9A']);

  // --- Create Timetable sheet ---
  var ttSheet = ss.getSheetByName(TIMETABLE_SHEET);
  if (!ttSheet) {
    ttSheet = ss.insertSheet(TIMETABLE_SHEET);
  } else {
    ttSheet.clear();
  }
  ttSheet.appendRow(['Grade', 'Section', 'Day', 'PeriodNum', 'PeriodStart', 'PeriodEnd', 'Subject', 'Teacher', 'Room']);

  // Period definitions
  var periods69 = [
    ['0', '8:10', '8:30'],
    ['00', '8:30', '9:00'],
    ['1', '9:00', '9:40'],
    ['2', '9:40', '10:20'],
    ['3', '10:20', '11:00'],
    ['4', '11:00', '11:40'],
    ['5', '11:40', '12:20'],
    ['L', '12:20', '1:00'],
    ['6', '1:00', '1:40'],
    ['7', '1:40', '2:20'],
    ['8', '2:20', '3:00']
  ];

  var periods35 = [
    ['0', '8:20', '8:40'],
    ['00', '8:40', '9:00'],
    ['1', '9:00', '9:40'],
    ['2', '9:40', '10:20'],
    ['3', '10:20', '11:00'],
    ['4', '11:00', '11:40'],
    ['5', '11:40', '12:20'],
    ['L', '12:20', '1:00'],
    ['6', '1:00', '1:40'],
    ['7', '1:40', '2:20'],
    ['8', '2:20', '3:00']
  ];

  var days = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'];

  // --- Grade 9A ---
  var g9a = {
    MON: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HOBBY', 'BIO', '2L', 'MATH', 'ENG', 'LUNCH BREAK', 'ENG', 'MUS', 'AF'],
    TUE: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HOBBY', 'PHY/CHEM', 'SOC', 'MATH', 'ICT', 'LUNCH BREAK', 'PHY/CHEM', 'ART', 'ACTIVITY'],
    WED: ['ASSEMBLY/YOGA', 'SHORT BREAK', '2L', 'SOC', 'SOC', 'MATH', 'PHY/CHEM', 'LUNCH BREAK', 'ENG', 'DANCE', ''],
    THU: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ICT', 'ENG', 'SOC', 'PHY/CHEM', 'MATH', 'LUNCH BREAK', 'KALARI', 'SPORTS', ''],
    FRI: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'BIO', 'MATH', 'IT', '3L', 'SOC', 'LUNCH BREAK', 'LIB', 'SW/CY', 'SW/CY'],
    SAT: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'SOC', 'MATH', '2L', 'ENG', '2L', 'LUNCH BREAK', '', '', '']
  };

  // --- Grade 8A ---
  var g8a = {
    MON: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HOBBY', 'ICT', 'BIO', '2L', 'PHY/CHEM', 'LUNCH BREAK', 'SOC', 'MUS', 'AF'],
    TUE: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HOBBY', '2L', 'MATH', 'PHY/CHEM', 'SOC', 'LUNCH BREAK', 'BIO', 'ART', 'ACTIVITY'],
    WED: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ENG', 'MATH', 'ENG', '2L', 'IT', 'LUNCH BREAK', 'SOC', 'DANCE', ''],
    THU: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ENG', '3L', 'MATH', 'ENG', 'BIO', 'LUNCH BREAK', 'KALARI', 'SPORTS', ''],
    FRI: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'MATH', 'SOC', '3L', 'PHY/CHEM', 'MATH', 'LUNCH BREAK', 'LIB', 'SW/CY', 'SW/CY'],
    SAT: ['ASSEMBLY/YOGA', 'SHORT BREAK', '2L', 'ENG', 'MATH', 'ICT', 'SOC', 'LUNCH BREAK', '', '', '']
  };

  // --- Grade 7A ---
  var g7a = {
    MON: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HOBBY', 'SOC', 'ENG', 'SOC', 'MATH', 'LUNCH BREAK', '2L', 'SW/CY', 'SW/CY'],
    TUE: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HOBBY', 'SCI', 'ICT', '2L', 'SCI', 'LUNCH BREAK', 'ENG', 'MUS', 'ACTIVITY'],
    WED: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'SCI', 'ENG', 'IT', 'SOC', 'MATH', 'LUNCH BREAK', '3L', 'ART', ''],
    THU: ['ASSEMBLY/YOGA', 'SHORT BREAK', '2L', 'MATH', 'ICT', 'SCI', 'ENG', 'LUNCH BREAK', 'LIB', 'DANCE', ''],
    FRI: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ENG', 'SCI', 'MATH', 'SOC', '2L', 'LUNCH BREAK', 'SPORTS', 'KALARI', ''],
    SAT: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'SCI', '3L', 'SOC', 'MATH', 'MATH', 'LUNCH BREAK', '', '', '']
  };

  // --- Grade 6A ---
  var g6a = {
    MON: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HOBBY', 'ENG', 'MATH', 'SCI', 'SOC', 'LUNCH BREAK', 'MATH', 'SW/CY', 'SW/CY'],
    TUE: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HOBBY', 'ENG', 'SCI', 'SOC', 'MATH', 'LUNCH BREAK', '3L', 'DANCE', 'ACTIVITY'],
    WED: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'MATH', '2L', 'SCI', 'IT', 'ENG', 'LUNCH BREAK', 'ICT', 'MUS', ''],
    THU: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'SCI', 'SOC', 'ENG', '3L', '2L', 'LUNCH BREAK', 'ART', 'LIB', ''],
    FRI: ['ASSEMBLY/YOGA', 'SHORT BREAK', '2L', 'ICT', 'SOC', 'MATH', 'SCI', 'LUNCH BREAK', 'SPORTS', 'KALARI', ''],
    SAT: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'MATH', 'SOC', 'ENG', '2L', 'SCI', 'LUNCH BREAK', '', '', '']
  };

  // --- Grade 5A (from Personal TT - Ms. Sailaja) ---
  var g5a = {
    MON: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ENG', 'MATH', 'SCI', 'EVS', 'HINDI', 'LUNCH BREAK', 'TELUGU', 'SW/CY', 'SW/CY'],
    TUE: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'MATH', 'ENG', 'HINDI', 'SCI', 'EVS', 'LUNCH BREAK', 'ART', 'DANCE', 'ACTIVITY'],
    WED: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'SCI', 'HINDI', 'MATH', 'ENG', 'TELUGU', 'LUNCH BREAK', 'IT', 'MUS', ''],
    THU: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ENG', 'SCI', 'EVS', 'MATH', 'HINDI', 'LUNCH BREAK', 'LIB', 'SPORTS', ''],
    FRI: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'MATH', 'TELUGU', 'ENG', 'SCI', 'EVS', 'LUNCH BREAK', 'KALARI', 'ACTIVITY', ''],
    SAT: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HINDI', 'MATH', 'ENG', 'SCI', 'TELUGU', 'LUNCH BREAK', '', '', '']
  };

  // --- Grade 5B (from Personal TT - Ms. Saritha Ambekar) ---
  var g5b = {
    MON: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'MATH', 'ENG', 'EVS', 'SCI', 'HINDI', 'LUNCH BREAK', 'TELUGU', 'ACTIVITY', ''],
    TUE: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ENG', 'MATH', 'VE', 'HINDI', 'SCI', 'LUNCH BREAK', 'EVS', 'ACTIVITY', ''],
    WED: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'SCI', 'ENG', 'HINDI', 'TELUGU', 'EVS', 'LUNCH BREAK', 'MATH', 'SW/CY', ''],
    THU: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'MATH', 'SCI', 'ENG', 'HINDI', '3L', 'LUNCH BREAK', 'ACTIVITY', '', ''],
    FRI: ['ASSEMBLY/YOGA', 'SHORT BREAK', '3L', 'ENG', 'MATH', 'SCI', '3L', 'LUNCH BREAK', 'ACTIVITY', '', ''],
    SAT: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HINDI', 'TELUGU', '3L', 'EVS', 'MATH', 'LUNCH BREAK', '', '', '']
  };

  // --- Grade 4A (from Personal TT - Ms. Madhavi) ---
  var g4a = {
    MON: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ENG', 'MATH', 'SCI', 'HINDI', 'EVS', 'LUNCH BREAK', 'TELUGU', 'VE', ''],
    TUE: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'MATH', 'ENG', 'HINDI', 'SCI', 'EVS', 'LUNCH BREAK', 'TELUGU', 'SW/CY', 'ACTIVITY'],
    WED: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'SCI', 'HINDI', 'ENG', 'MATH', 'EVS', 'LUNCH BREAK', 'IT', 'AF', 'ACTIVITY'],
    THU: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HINDI', 'MATH', 'SCI', 'ENG', 'TELUGU', 'LUNCH BREAK', 'LIB', 'DANCE', 'ACTIVITY'],
    FRI: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'MATH', 'SCI', 'ENG', 'EVS', 'HINDI', 'LUNCH BREAK', 'MUS', 'SPORTS', 'ACTIVITY'],
    SAT: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'EVS', 'ENG', 'MATH', 'SCI', 'HINDI', 'LUNCH BREAK', '', '', '']
  };

  // --- Grade 4B (from Personal TT - Ms. Vidya) ---
  var g4b = {
    MON: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'MATH', 'ENG', 'HINDI', 'SCI', 'EVS', 'LUNCH BREAK', 'TELUGU', 'ACTIVITY', ''],
    TUE: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ENG', 'MATH', 'SCI', 'HINDI', 'EVS', 'LUNCH BREAK', 'TELUGU', 'SW/CY', 'ACTIVITY'],
    WED: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HINDI', 'ENG', 'MATH', 'EVS', 'SCI', 'LUNCH BREAK', 'IT', 'AF', ''],
    THU: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ENG', 'SCI', 'HINDI', 'EVS', 'MATH', 'LUNCH BREAK', 'DANCE', 'LIB', 'ACTIVITY'],
    FRI: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'MATH', 'ENG', 'LIB', 'SCI', 'EVS', 'LUNCH BREAK', 'MUS', 'SPORTS', 'ACTIVITY'],
    SAT: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'SCI', 'HINDI', 'ENG', 'EVS', 'MATH', 'LUNCH BREAK', '', '', '']
  };

  // --- Grade 3A (from Personal TT - Ms. Sridevi) ---
  var g3a = {
    MON: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ENG', 'MATH', 'HINDI', 'SCI', 'EVS', 'LUNCH BREAK', 'TELUGU', 'ACTIVITY', ''],
    TUE: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'MATH', 'ENG', 'SCI', 'EVS', 'HINDI', 'LUNCH BREAK', 'TELUGU', 'ACTIVITY', ''],
    WED: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'SCI', 'HINDI', 'ENG', 'MATH', 'EVS', 'LUNCH BREAK', 'IT', 'ACTIVITY', ''],
    THU: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ENG', 'MATH', 'SCI', 'HINDI', '3L', 'LUNCH BREAK', 'SW/CY', 'ACTIVITY', ''],
    FRI: ['ASSEMBLY/YOGA', 'SHORT BREAK', '3L', 'TELUGU', 'EVS', 'MATH', 'ENG', 'LUNCH BREAK', 'DANCE', 'MUS', 'AF'],
    SAT: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HINDI', 'SCI', 'MATH', 'ENG', 'EVS', 'LUNCH BREAK', '', '', '']
  };

  // --- Grade 3B (from Personal TT - Ms. Usha) ---
  var g3b = {
    MON: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'MATH', 'ENG', 'SCI', 'HINDI', 'EVS', 'LUNCH BREAK', 'TELUGU', 'VE', ''],
    TUE: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ENG', 'MATH', 'HINDI', 'EVS', 'SCI', 'LUNCH BREAK', 'IT', 'ACTIVITY', ''],
    WED: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'HINDI', 'SCI', 'ENG', 'EVS', 'MATH', 'LUNCH BREAK', 'TELUGU', 'ACTIVITY', ''],
    THU: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'SCI', 'ENG', 'HINDI', 'MATH', 'EVS', 'LUNCH BREAK', 'SW/CY', 'ACTIVITY', ''],
    FRI: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'ENG', 'MATH', 'SCI', 'EVS', 'HINDI', 'LUNCH BREAK', 'DANCE', 'MUS', ''],
    SAT: ['ASSEMBLY/YOGA', 'SHORT BREAK', 'EVS', 'MATH', 'HINDI', 'SCI', 'ENG', 'LUNCH BREAK', '', '', '']
  };

  // Map of all classes
  var allClasses = [
    { grade: '9', section: 'A', data: g9a, periods: periods69, teacher: 'Ms. Maneesha' },
    { grade: '8', section: 'A', data: g8a, periods: periods69, teacher: 'Ms. Madhuri' },
    { grade: '7', section: 'A', data: g7a, periods: periods69, teacher: 'Ms. Akanksha' },
    { grade: '6', section: 'A', data: g6a, periods: periods69, teacher: 'Ms. Rajsree' },
    { grade: '5', section: 'A', data: g5a, periods: periods35, teacher: 'Ms. Sailaja' },
    { grade: '5', section: 'B', data: g5b, periods: periods35, teacher: 'Ms. Saritha Ambekar' },
    { grade: '4', section: 'A', data: g4a, periods: periods35, teacher: 'Ms. Madhavi' },
    { grade: '4', section: 'B', data: g4b, periods: periods35, teacher: 'Ms. Vidya' },
    { grade: '3', section: 'A', data: g3a, periods: periods35, teacher: 'Ms. Sridevi' },
    { grade: '3', section: 'B', data: g3b, periods: periods35, teacher: 'Ms. Usha' }
  ];

  var rows = [];
  for (var c = 0; c < allClasses.length; c++) {
    var cls = allClasses[c];
    for (var d = 0; d < days.length; d++) {
      var dayName = days[d];
      var subjects = cls.data[dayName];
      if (!subjects) continue;
      for (var p = 0; p < cls.periods.length; p++) {
        var period = cls.periods[p];
        var subject = (p < subjects.length) ? subjects[p] : '';
        rows.push([cls.grade, cls.section, dayName, period[0], period[1], period[2], subject, '', '']);
      }
    }
  }

  // Batch write for performance
  if (rows.length > 0) {
    ttSheet.getRange(2, 1, rows.length, 9).setValues(rows);
  }

  Logger.log('Setup complete: ' + rows.length + ' timetable rows created.');
}
