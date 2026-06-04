// ==========================================
// 系統設定
// ==========================================
const PASSWORD_KEY = 'system_password';

function getSystemPassword() {
  return PropertiesService.getScriptProperties().getProperty(PASSWORD_KEY) || '25003';
}

function setSystemPassword(pwd) {
  PropertiesService.getScriptProperties().setProperty(PASSWORD_KEY, pwd);
}
const CACHE_TTL = 60; // 快取 TTL 60 秒 (1 分鐘)

// ==========================================
// CacheService — deep module
// Interface: get(key), put(key, value), invalidate(action, params), clearAll()
// Implementation hides: serialization, TTL, invalidation mapping, template resolution
// ==========================================
const CACHE = {
  _cache: null,
  _getCache: function() {
    if (!this._cache) this._cache = CacheService.getUserCache();
    return this._cache;
  },

  get: function(key) {
    var val = this._getCache().get(key);
    return val ? JSON.parse(val) : null;
  },

  put: function(key, value) {
    this._getCache().put(key, JSON.stringify(value), CACHE_TTL);
  },

  // Invalidation mapping: action name → key template list
  // {date} → resolved from params.date
  // static key → used as-is
  _invalidationMap: {
    'writeWorkout': ['prog_{date}', 'all_events', 'raw_data'],
    'moveSchedule': ['prog_{oldDate}', 'prog_{newDate}', 'all_events', 'raw_data'],
    'updateSettings': ['settings'],
    'generateProgram': ['all_events', 'raw_data']
  },

  invalidate: function(action, params) {
    var templates = this._invalidationMap[action];
    if (!templates) throw new Error('Unknown cache action: ' + action);

    for (var i = 0; i < templates.length; i++) {
      var tpl = templates[i];
      var key = tpl;
      // Resolve template variables: {fieldName} → params.fieldName
      var match = tpl.match(/\{(\w+)\}/);
      if (match) {
        var field = match[1];
        if (params[field] === undefined) {
          throw new Error('Missing param "' + field + '" for action "' + action + '"');
        }
        key = tpl.replace('{' + field + '}', params[field]);
      }
      this._getCache().remove(key);
    }
  },

  clearAll: function() {
    var cache = this._getCache();
    var keys = cache.getAll();
    for (var key in keys) {
      if (keys.hasOwnProperty(key)) cache.remove(key);
    }
    Logger.log("所有快取已清除（包含 " + Object.keys(keys).length + " 筆）");
  }
};

// 🧹 清除所有快取（後台改資料後手動執行）
function clearAllCache() {
  CACHE.clearAll();
}

// ==========================================
// DateUtil — deep module
// Interface: formatDate(val) → 'yyyy-MM-dd'
// Implementation hides: timezone query, Date vs string dispatch, format pattern
// ==========================================
const DateUtil = {
  _tz: Session.getScriptTimeZone(),
  _fmt: 'yyyy-MM-dd',

  formatDate: function(val) {
    if (val instanceof Date) {
      return Utilities.formatDate(val, this._tz, this._fmt);
    }
    // Already a string — normalize (trim whitespace, take first 10 chars)
    return String(val).trim().substring(0, 10);
  }
};

// ==========================================
// Compute — deep module
// Interface: est1RM(w, r), detectPR(w, r, history), nextMonday(dateStr), calcWeight(tm, percent)
// Implementation hides: Epley cap (1.25×), PR comparison logic, date arithmetic, rounding
// ==========================================
const Compute = {
  est1RM: function(weight, reps) {
    var epley = weight * (1 + reps / 30);
    return Math.min(epley, weight * 1.25);
  },

  detectPR: function(weight, reps, history) {
    var types = [];
    if (weight > history.bestWeight) types.push('重量');
    if (reps > history.bestReps) types.push('次數');
    if (weight * reps > history.bestVolume) types.push('總量');
    return types;
  },

  nextMonday: function(dateStr) {
    var base = new Date(dateStr);
    var day = base.getDay();
    var diff = base.getDate() - day + (day == 0 ? -6 : 1);
    return new Date(base.setDate(diff));
  },

  calcWeight: function(tm, percent) {
    return Math.round((tm * percent) / 2.5) * 2.5;
  }
};

// 讀取全表資料（帶快取）
function readAllProgramData() {
  let data = CACHE.get('raw_data');
  if (data) return data;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Program_Plan');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rawData = sheet.getRange(1, 1, lastRow, 12).getValues();

  // 日期欄位轉字串，避免 JSON 序列化後 Date 變成 ISO 字串導致比對失敗
  data = rawData.map((r) => {
    return [
      DateUtil.formatDate(r[0]),
      r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11]
    ];
  });

  CACHE.put('raw_data', data);
  return data;
}

function doGet(e) {
  return ContentService.createTextOutput("Camp 5/3/1 API is running normally.");
}

function doPost(e) {
  let response = { success: false, error: '未知錯誤', data: null };
  
  try {
    const params = JSON.parse(e.postData.contents);

    // 🔒 密碼驗證攔截器
    if (params.password !== getSystemPassword()) {
      response = { success: false, error: '密碼錯誤或未經授權的存取！', error_code: 'AUTH_ERROR', data: null };
      return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
    }

    let result;
    switch (params.action) {
      case 'init':
        initSystem();
        result = "系統初始化完成";
        break;
      case 'getSettings':
        result = getSettings();
        break;
      
      // 🚀 新增：三合一加速載入 (登入時一次拿完所有資料)
      case 'initLoad':
        result = {
          settings: getSettings(),
          events: getAllPrograms(),
          workouts: getProgramByDate(params.targetDate)
        };
        break;

      case 'saveSettings':
        result = saveSettings(params.data);
        break;
      case 'generateProgram':
        result = generateProgram(params.startDateStr, params.isIncrement, params.tmRatio, params.progType);
        break;
      case 'getProgramByDate':
        result = getProgramByDate(params.targetDate);
        break;
      case 'getAllPrograms':
        result = getAllPrograms();
        break;
      case 'logWorkout':
        result = logWorkout(params.logData);
        break;
      case 'addExtraSet':
        result = addExtraSet(params.logData);
        break;
      case 'calcNextMonday':
        result = DateUtil.formatDate(Compute.nextMonday(params.date));
        break;
      case 'moveWorkoutDate':
        result = moveWorkoutDate(params.oldDate, params.newDateStr);
        break;
      case 'getPRs':
        result = getPRs();
        break;
      case 'getStatistics':
        result = getStatistics();
        break;
      case 'saveBodyWeight':
        result = saveBodyWeight(params.bodyData);
        break;
      case 'getBodyWeight':
        result = getBodyWeight();
        break;
      case 'changePassword':
        if (params.oldPassword === getSystemPassword()) {
          setSystemPassword(params.newPassword);
          result = '密碼已更新';
        } else {
          var err = new Error('舊密碼錯誤');
          err.error_code = 'AUTH_ERROR';
          throw err;
        }
        break;
      default:
        var err2 = new Error("未知的 API 請求");
        err2.error_code = 'VALIDATION_ERROR';
        throw err2;
    }
    
    response.success = true;
    response.data = result;
    response.error = null;
    
  } catch (error) {
    response = {
      success: false,
      error: error.toString(),
      error_code: error.error_code || 'SERVER_ERROR',
      data: null
    };
  }
  
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 以下為原本的資料庫邏輯 (未做更動)
// ==========================================

function initSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
    settingsSheet.appendRow(['設定項目', '數值']);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e9ecef');
    const defaultSettings = [
      ['Squat_1RM', '100'], ['Squat_Added', '0'],
      ['Deadlift_1RM', '120'], ['Deadlift_Added', '0'],
      ['BenchPress_1RM', '80'], ['BenchPress_Added', '0'],
      ['OverheadPress_1RM', '50'], ['OverheadPress_Added', '0']
    ];
    settingsSheet.getRange(2, 1, 8, 2).setValues(defaultSettings);
  }

  let programSheet = ss.getSheetByName('Program_Plan');
  if (programSheet) {
    let headers = programSheet.getRange(1, 1, 1, programSheet.getLastColumn()).getValues()[0];
    if (headers[6] !== '百分比' && headers[6] === '已完成') {
      programSheet.insertColumnBefore(7);
      programSheet.getRange(1, 7).setValue('百分比');
      programSheet.getRange(1, 8).setValue('已完成');
    }
    headers = programSheet.getRange(1, 1, 1, programSheet.getLastColumn()).getValues()[0];
    if (headers.length < 10 || headers[8] !== '實際重量') {
      programSheet.getRange(1, 9).setValue('實際重量');
      programSheet.getRange(1, 10).setValue('實際次數');
    }
    // 新增暖身組欄位
    headers = programSheet.getRange(1, 1, 1, programSheet.getLastColumn()).getValues()[0];
    if (headers.length < 12 || headers[10] !== '暖身重量') {
      programSheet.getRange(1, 11).setValue('暖身重量');
      programSheet.getRange(1, 12).setValue('暖身次數');
    }
  } else {
    programSheet = ss.insertSheet('Program_Plan');
    programSheet.appendRow(['日期', '週次', '動作', '組數', '預定重量(kg)', '預定次數', '百分比', '已完成', '實際重量', '實際次數', '暖身重量', '暖身次數']);
    programSheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#e9ecef');
  }

  let logSheet = ss.getSheetByName('Workout_Log');
  if (!logSheet) {
    logSheet = ss.insertSheet('Workout_Log');
    logSheet.appendRow(['時間戳記', '日期', '動作', '實際重量(kg)', '實際次數', '備註', '估算 1RM']);
    logSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e9ecef');
  }

  // 新增 PR_Log 工作表
  let prSheet = ss.getSheetByName('PR_Log');
  if (!prSheet) {
    prSheet = ss.insertSheet('PR_Log');
    prSheet.appendRow(['日期', '動作', '重量(kg)', '次數', '類型']);
    prSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#fff3cd');
  }

  // 新增 Body_Weight 工作表
  let bwSheet = ss.getSheetByName('Body_Weight');
  if (!bwSheet) {
    bwSheet = ss.insertSheet('Body_Weight');
    bwSheet.appendRow(['日期', '體重(kg)']);
    bwSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#d1ecf1');
  }
}

function getSettings() {
  // 🚀 讀快取
  let settings = CACHE.get('settings');
  if (settings) return settings;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sheet) { initSystem(); return getSettings(); }
  const data = sheet.getDataRange().getValues();
  settings = {};
  for (let i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }
  // 🚀 寫回快取
  CACHE.put('settings', settings);
  return settings;
}

function saveSettings(settingsData) {
  initSystem();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  let existingKeys = {};
  for (let i = 1; i < data.length; i++) { existingKeys[data[i][0]] = i + 1; }
  let batchUpdates = [];
  let appendedKeys = [];
  for (let key in settingsData) {
    if (existingKeys[key]) {
      batchUpdates.push([existingKeys[key], 2, Number(settingsData[key])]);
    } else {
      appendedKeys.push([key, Number(settingsData[key])]);
    }
  }
  // Update existing rows
  batchUpdates.forEach(u => sheet.getRange(u[0], u[1]).setValue(u[2]));
  // Append new keys (still one-by-one as they're new rows)
  appendedKeys.forEach(k => sheet.appendRow(k));

  // 🚀 寫入後清除設定快取
  CACHE.invalidate('updateSettings');

  return "設定已儲存";
}

function generateProgram(startDateStr, isIncrement, tmRatio, progType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const programSheet = ss.getSheetByName('Program_Plan');
  let settings = getSettings();

  if (isIncrement) {
    settings['Squat_Added'] = Number(settings['Squat_Added'] || 0) + 5;
    settings['Deadlift_Added'] = Number(settings['Deadlift_Added'] || 0) + 5;
    settings['BenchPress_Added'] = Number(settings['BenchPress_Added'] || 0) + 2.5;
    settings['OverheadPress_Added'] = Number(settings['OverheadPress_Added'] || 0) + 2.5;
    saveSettings(settings); 
  }

  const multiplier = parseFloat(tmRatio) || 0.9;
  const tm = {
    '深蹲 Squat': (Number(settings['Squat_1RM']) + Number(settings['Squat_Added'] || 0)) * multiplier,
    '肩推 Press': (Number(settings['OverheadPress_1RM']) + Number(settings['OverheadPress_Added'] || 0)) * multiplier,
    '硬舉 Deadlift': (Number(settings['Deadlift_1RM']) + Number(settings['Deadlift_Added'] || 0)) * multiplier,
    '臥推 Bench': (Number(settings['BenchPress_1RM']) + Number(settings['BenchPress_Added'] || 0)) * multiplier
  };

  let cycleScheme;
  if (progType === '5x5') {
    cycleScheme = [
      { week: 1, percents: [0.75, 0.75, 0.75, 0.75, 0.75], reps: ['5', '5', '5', '5', '5'] },
      { week: 2, percents: [0.80, 0.80, 0.80, 0.80, 0.80], reps: ['5', '5', '5', '5', '5'] },
      { week: 3, percents: [0.85, 0.85, 0.85, 0.85, 0.85], reps: ['5', '5', '5', '5', '5'] },
      { week: 4, percents: [0.40, 0.50, 0.60], reps: ['10', '10', '10'] }
    ];
  } else {
    cycleScheme = [
      { week: 1, percents: [0.65, 0.75, 0.85], reps: ['5', '5', '5+'] },
      { week: 2, percents: [0.70, 0.80, 0.90], reps: ['3', '3', '3+'] },
      { week: 3, percents: [0.75, 0.85, 0.95], reps: ['5', '3', '1+'] },
      { week: 4, percents: [0.40, 0.50, 0.60], reps: ['10', '10', '10'] }
    ];
  }

  const typeLabel = progType === '5x5' ? '5x5' : '5/3/1';
  const movements = ['深蹲 Squat', '肩推 Press', '硬舉 Deadlift', '臥推 Bench'];
  let mondayDate = Compute.nextMonday(startDateStr);

  let newRows = [];
  let targetDatesSet = new Set(); 

  for (let w = 0; w < 4; w++) {
    for (let d = 0; d < 4; d++) {
      let trainingDate = new Date(mondayDate);
      trainingDate.setDate(mondayDate.getDate() + (w * 7) + d);
      let dateStr = DateUtil.formatDate(trainingDate);
      
      targetDatesSet.add(dateStr); 
      let move = movements[d];
      let scheme = cycleScheme[w];
      
      for (let s = 0; s < scheme.reps.length; s++) {
        let weight = Compute.calcWeight(tm[move], scheme.percents[s]);
        let percentStr = Math.round(scheme.percents[s] * 100) + '%';
        // 第一組預填暖身組（50%×5、60%×3）
        let warmupW = '', warmupR = '';
        if (s === 0) {
          let w1 = Math.round((tm[move] * 0.50) / 2.5) * 2.5;
          let w2 = Math.round((tm[move] * 0.60) / 2.5) * 2.5;
          warmupW = `${w1}kg×5 / ${w2}kg×3`;
          warmupR = 'warmup';
        }
        newRows.push([dateStr, `Week ${w+1} (${typeLabel})`, move, `Set ${s+1}`, weight, scheme.reps[s], percentStr, '', '', '', warmupW, warmupR]);
      }
    }
  }

  let existingData = [];
  if (programSheet.getLastRow() > 1) {
    let lastCol = programSheet.getLastColumn() > 11 ? programSheet.getLastColumn() : 12;
    existingData = programSheet.getRange(2, 1, programSheet.getLastRow() - 1, lastCol).getValues();
  }

  let keptData = [];
  for (let i = 0; i < existingData.length; i++) {
    let row = existingData[i];
    let rowDateStr = DateUtil.formatDate(row[0]); 

    if (!targetDatesSet.has(rowDateStr)) {
      while(row.length < 12) row.push('');
      keptData.push(row);
    }
  }

  let finalData = keptData.concat(newRows);
  finalData.sort((a, b) => {
    let da = (a[0] instanceof Date) ? a[0] : new Date(a[0]);
    let db = (b[0] instanceof Date) ? b[0] : new Date(b[0]);
    return da - db;
  });

  if (programSheet.getLastRow() > 1) {
    programSheet.getRange(2, 1, programSheet.getLastRow() - 1, 12).clearContent();
  }
  if (finalData.length > 0) {
    programSheet.getRange(2, 1, finalData.length, 12).setValues(finalData);
  }

  // 🚀 清除課表快取
  _invalidateProgramCaches();

  return "新週期課表已生成！（已自動覆寫衝突日期並完成排序）";
}

// 清除課表快取
function _invalidateProgramCaches() {
  CACHE.invalidate('generateProgram');
}

function getProgramByDate(targetDate) {
  // 🚀 讀快取
  const cached = CACHE.get('prog_' + targetDate);
  if (cached) return cached;

  // 從全表資料過濾，不再讀 Sheets
  const data = readAllProgramData();
  const result = data.reduce((result, r, idx) => {
    if (r[0] === targetDate || (r[0] instanceof Date && DateUtil.formatDate(r[0]) === targetDate)) {
      let actualW = (r[8] !== undefined && r[8] !== '') ? String(r[8]) : String(r[4]);
      let actualR = (r[9] !== undefined && r[9] !== '') ? String(r[9]) : String(r[5]);
      // 計算估算 1RM（Epley 公式，上限 1.25 倍）
      var w = Number(actualW);
      var rp = Number(actualR.replace('+', ''));
      var est1RM = (w > 0 && rp > 0) ? Compute.est1RM(w, rp) : 0;

      result.push({
        rowIdx: idx + 1,
        date: DateUtil.formatDate(r[0]),
        week: String(r[1]), movement: String(r[2]), set: String(r[3]),
        targetWeight: String(r[4]), targetReps: String(r[5]), percent: String(r[6]), isLogged: String(r[7]),
        actualWeight: actualW, actualReps: actualR,
        warmupWeight: r[10] || '', warmupReps: r[11] || '',
        est1RM: est1RM
      });
    }
    return result;
  }, []);
  // 🚀 寫回快取
  CACHE.put('prog_' + targetDate, result);
  return result;
}

function getAllPrograms() {
  // 🚀 讀快取
  let events = CACHE.get('all_events');
  if (events) return events;

  // 從全表資料計算，不再讀 Sheets
  const data = readAllProgramData();
  events = {};
  for(let i = 1; i < data.length; i++) {
    let r = data[i];
    let date = DateUtil.formatDate(r[0]);
    let week = String(r[1]);
    let move = String(r[2]);
    let weight = String(r[4]);
    let isDone = String(r[7]) === 'V';

    if(!events[date]) {
      events[date] = { move: move, week: week, maxWeight: weight, sets: 1, done: isDone, doneCount: isDone ? 1 : 0 };
    } else {
      if (Number(weight) > Number(events[date].maxWeight)) events[date].maxWeight = weight;
      events[date].sets += 1;
      if (isDone) events[date].doneCount += 1;
      if (!isDone) events[date].done = false;
    }
  }
  // 修正：doneCount === sets 時，done 應為 true
  for(let d in events) {
    if(events[d].doneCount === events[d].sets) events[d].done = true;
  }
  // 🚀 寫回快取
  CACHE.put('all_events', events);
  return events;
}

function logWorkout(logData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Workout_Log');
  const progSheet = ss.getSheetByName('Program_Plan');
  const rowIdx = logData.rowIdx;
  const dateStr = logData.date;
  const move = logData.movement;
  const set = logData.set || '';
  const actW = Number(logData.actualWeight);
  const actR = Number(logData.actualReps);

  // 計算估算 1RM（Epley 公式，上限 1.25 倍）
  let est1RM = 0;
  if (actW > 0 && actR > 0) {
    est1RM = Compute.est1RM(actW, actR);
  }

  // 判斷是否為追加組（3D/3E/3F...）
  // Set 3 = 原訂，Set 3A/3B/3C... = 追加組
  var isExtra = set.indexOf('Set 3') === 0 && set.length > 6;

  if (isExtra) {
    // 追加組：一律新增一筆 Workout_Log（不覆蓋前一組）
    logSheet.appendRow([new Date(), dateStr, move, actW, actR, '', est1RM]);
  } else {
    // 原訂組：更新最後一筆同日期/動作的紀錄（不新增重複）
    const logRows = logSheet.getDataRange().getValues();
    let lastLogIdx = -1;
    for (let i = logRows.length - 1; i >= 1; i--) {
      if (String(logRows[i][1]) === dateStr && String(logRows[i][2]) === move) {
        lastLogIdx = i;
        break;
      }
    }
    if (lastLogIdx >= 0) {
      logSheet.getRange(lastLogIdx + 1, 4, 1, 3).setValues([[actW, actR, est1RM]]);
    } else {
      logSheet.appendRow([new Date(), dateStr, move, actW, actR, '', est1RM]);
    }
  }

  // 更新 Program_Plan：先用前端 rowIdx，失敗則搜尋正確列
  try {
    progSheet.getRange(rowIdx, 8, 1, 3).setValues([['V', actW, actR]]);
  } catch (e) {
    // rowIdx 不正確時，搜尋該日期/動作/組數的正確列
    const progData = progSheet.getDataRange().getValues();
    for (let i = progData.length - 1; i >= 1; i--) {
      if (DateUtil.formatDate(progData[i][0]) === dateStr &&
          String(progData[i][2]) === move &&
          String(progData[i][3]) === set) {
        progSheet.getRange(i + 1, 8, 1, 3).setValues([['V', actW, actR]]);
        break;
      }
    }
  }

  // 檢查是否創 PR
  let prTypes = checkAndRecordPR(dateStr, move, actW, actR);

  // 🚀 寫入後清除相關快取
  CACHE.invalidate('writeWorkout', { date: dateStr });

  return { message: "記錄完成", est1RM: est1RM, pr: prTypes };
}

// 為 Set 3 追加一組（3D/3E/3F...）
function addExtraSet(logData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progSheet = ss.getSheetByName('Program_Plan');
  const logSheet = ss.getSheetByName('Workout_Log');
  const dateStr = logData.date;
  const move = logData.movement;
  const actW = Number(logData.actualWeight);
  const actR = Number(logData.actualReps);

  if (actW <= 0 || actR <= 0) {
    throw new Error('重量和次數必須大於 0');
  }

  // 找出該日期/動作最後一組 Set 3（3A/3B/3C/3D...）
  // Set 3 = 原訂最後一組，Set 3A/3B/3C... = 追加組
  const data = readAllProgramData();
  var lastSet3Data = null;
  var hasSet3 = false; // 是否有原訂 Set 3
  var maxLetterIdx = -1; // 追加組字母：A=0, B=1, C=2...
  var extraCount = 0;
  for (var i = data.length - 1; i >= 0; i--) {
    var r = data[i];
    var setStr = String(r[3]);
    if (String(r[0]) === dateStr && String(r[2]) === move && setStr.indexOf('Set 3') === 0) {
      if (setStr === 'Set 3') {
        hasSet3 = true;
        lastSet3Data = r;
      } else if (setStr.length > 6) {
        // 追加組：Set 3A, 3B, 3C...
        var letter = setStr.charCodeAt(setStr.length - 1) - 65; // A=0, B=1, C=2...
        if (letter > maxLetterIdx) {
          maxLetterIdx = letter;
          lastSet3Data = r;
        }
        extraCount++;
      }
    }
  }

  if (!hasSet3) {
    throw new Error('找不到該日期的 Set 3。請確認課表已生成且該日期/動作存在 Set 3。資料筆數：' + data.length);
  }

  // 限制最多 5 組追加（3D/3E/3F/3G/3H）
  if (extraCount >= 5) {
    throw new Error('已達追加組上限（5 組）');
  }

  // 計算下一組字母
  var nextLetter = String.fromCharCode(65 + maxLetterIdx + 1); // C=2 -> next = D
  var nextSet = 'Set 3' + nextLetter;

  // 從最後一組 Set 3 複製資料，填入新的重量/次數
  var newRow = [
    lastSet3Data[0],   // date
    lastSet3Data[1],   // week
    lastSet3Data[2],   // movement
    nextSet,           // set (3D/3E/...)
    '',                // targetWeight (追加組無原訂重量)
    '',                // targetReps (追加組無原訂次數)
    lastSet3Data[6],   // percent
    '',                // isLogged
    String(actW),      // actualWeight
    String(actR),      // actualReps
    lastSet3Data[10],  // warmupWeight
    lastSet3Data[11]   // warmupReps
  ];

  // appendRow 會加在最後一列
  progSheet.appendRow(newRow);

  // 計算估算 1RM
  var est1RM = Compute.est1RM(actW, actR);

  // 寫入 Workout_Log
  logSheet.appendRow([new Date(), dateStr, move, actW, actR, '', est1RM]);

  // 檢查是否創 PR
  var prTypes = checkAndRecordPR(dateStr, move, actW, actR);

  // 🚀 寫入後清除相關快取
  CACHE.invalidate('writeWorkout', { date: dateStr });

  return { message: '已新增 ' + nextSet, est1RM: est1RM, pr: prTypes, nextSet: nextSet };
}

// 檢查並記錄 PR，回傳創下的 PR 類型
function checkAndRecordPR(dateStr, move, weight, reps) {
  const prSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PR_Log');
  if (!prSheet) return [];

  const prData = prSheet.getDataRange().getValues();
  let bestWeight = 0;
  let bestReps = 0;
  let bestVolume = 0;

  for (let i = 1; i < prData.length; i++) {
    if (String(prData[i][1]) === move) {
      const w = Number(prData[i][2]);
      const r = Number(prData[i][3]);
      if (w > bestWeight) bestWeight = w;
      if (r > bestReps) bestReps = r;
      const vol = w * r;
      if (vol > bestVolume) bestVolume = vol;
    }
  }

  let prTypes = Compute.detectPR(weight, reps, { bestWeight, bestReps, bestVolume });

  if (prTypes.length > 0) {
    const dateFormatted = DateUtil.formatDate(new Date(dateStr));
    prTypes.forEach(type => {
      prSheet.appendRow([dateFormatted, move, weight, reps, type + ' PR']);
    });
  }
  return prTypes;
}

function getPRs() {
  const prSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PR_Log');
  if (!prSheet || prSheet.getLastRow() < 2) return [];

  const data = prSheet.getRange(2, 1, prSheet.getLastRow() - 1, 5).getValues();
  return data.map(r => ({
    date: DateUtil.formatDate(r[0]),
    movement: String(r[1]),
    weight: Number(r[2]),
    reps: Number(r[3]),
    type: String(r[4])
  }));
}

// 訓練統計
function getStatistics() {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Workout_Log');
  if (!logSheet || logSheet.getLastRow() < 2) return { summary: {}, volumeTrends: {}, weightTrends: {} };

  const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 7).getValues();

  // 過濾有效紀錄（有實際重量和次數的）
  const validLogs = logData.filter(r => Number(r[3]) > 0 && Number(r[4]) > 0);

  // 各動作統計
  const movementStats = {};
  let totalSessions = 0;
  let totalVolume = 0;
  const dateSet = new Set();

  validLogs.forEach(r => {
    const dateStr = DateUtil.formatDate(r[1])
    const move = String(r[2]);
    const weight = Number(r[3]);
    const reps = Number(r[4]);
    const est1RM = Number(r[6]) || 0;
    const vol = weight * reps;

    if (!movementStats[move]) {
      movementStats[move] = { bestWeight: 0, bestReps: 0, bestVolume: 0, sessions: 0, est1RMs: [], volumeData: [] };
    }

    movementStats[move].sessions++;
    if (weight > movementStats[move].bestWeight) movementStats[move].bestWeight = weight;
    if (reps > movementStats[move].bestReps) movementStats[move].bestReps = reps;
    if (vol > movementStats[move].bestVolume) movementStats[move].bestVolume = vol;
    if (est1RM > 0) movementStats[move].est1RMs.push(est1RM);
    movementStats[move].volumeData.push({ date: dateStr, volume: vol, weight: weight, reps: reps });

    totalSessions++;
    totalVolume += vol;
    dateSet.add(dateStr);
  });

  // 各動作每日訓練量（重量 × 次數，每天加總）
  const volumeTrends = {};
  validLogs.forEach(r => {
    const dateStr = DateUtil.formatDate(r[1])
    const move = String(r[2]);
    const weight = Number(r[3]);
    const reps = Number(r[4]);
    if (!volumeTrends[move]) volumeTrends[move] = {};
    if (!volumeTrends[move][dateStr]) volumeTrends[move][dateStr] = 0;
    volumeTrends[move][dateStr] += weight * reps;
  });

  // 各動作每日最大重量（取每天最重的一筆）
  const weightTrends = {};
  validLogs.forEach(r => {
    const dateStr = DateUtil.formatDate(r[1])
    const move = String(r[2]);
    const weight = Number(r[3]);
    if (!weightTrends[move]) weightTrends[move] = {};
    if (!weightTrends[move][dateStr] || weight > weightTrends[move][dateStr]) {
      weightTrends[move][dateStr] = weight;
    }
  });

  // 最近 30 天訓練天數
  const now = new Date();
  let recentDays = 0;
  for (let i = 0; i < 30; i++) {
    const d = new Date(now);
    d.setDate(d.getDate() - i);
    const ds = DateUtil.formatDate(d);
    if (dateSet.has(ds)) recentDays++;
  }

  // 各動作最佳估算 1RM
  const bestEst1RM = {};
  for (const move in movementStats) {
    const est1RMs = movementStats[move].est1RMs;
    bestEst1RM[move] = est1RMs.length > 0 ? Math.round(Math.max(...est1RMs)) : 0;
  }

  return {
    summary: {
      totalSessions,
      totalVolume,
      recentDays,
      bestEst1RM
    },
    movementStats,
    volumeTrends,
    weightTrends
  };
}

// 體重紀錄
function saveBodyWeight(bodyData) {
  const bwSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Body_Weight');
  if (!bwSheet) return '找不到 Body_Weight 工作表';

  const data = bwSheet.getDataRange().getValues();
  // 檢查該日期是否已存在
  for (let i = 1; i < data.length; i++) {
    const rowDate = DateUtil.formatDate(data[i][0]);
    if (rowDate === bodyData.date) {
      bwSheet.getRange(i + 1, 2).setValue(bodyData.weight);
      return '體重已更新';
    }
  }

  bwSheet.appendRow([bodyData.date, bodyData.weight]);
  return '體重已紀錄';
}

function getBodyWeight() {
  const bwSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Body_Weight');
  if (!bwSheet || bwSheet.getLastRow() < 2) return [];

  const data = bwSheet.getRange(2, 1, bwSheet.getLastRow() - 1, 2).getValues();
  return data.map(r => ({
    date: DateUtil.formatDate(r[0]),
    weight: Number(r[1])
  })).reverse(); // 最新的在最前面
}

function moveWorkoutDate(oldDate, newDateStr) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Program_Plan');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return "找不到該日期的課表";
  const data = sheet.getRange(1, 1, lastRow, 12).getValues();
  let found = false;
  for(let i = 1; i < data.length; i++) {
    let rowDate = DateUtil.formatDate(data[i][0]);
    if(rowDate === oldDate) {
      sheet.getRange(i + 1, 1).setValue(newDateStr);
      found = true;
    }
  }
  if(found) {
    // 🚀 寫入後清除相關快取
    CACHE.invalidate('moveSchedule', { oldDate: oldDate, newDate: newDateStr });
    return "課表已成功移動至 " + newDateStr;
  }
  return "找不到該日期的課表";
}