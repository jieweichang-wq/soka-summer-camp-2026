/**
 * 2026 創價歡樂夏令營 - 運營本部管理平台 Google Apps Script 後端資料庫及日曆同步服務
 * 
 * 本程式主要處理三件事：
 * 1. 下載並解析公開的 Google 行事曆 iCal (.ics) 檔案以同步活動資訊。
 * 2. 進度回報：提供讀取、寫入及刪除功能，資料儲存於試算表的「進度回報」分頁中。
 * 3. 題目統計：提供讀取、寫入及刪除功能，資料儲存於試算表的「題目統計」分頁中。
 */

function doGet(e) {
  // 防禦性檢查是否傳入參數
  if (!e || !e.parameter) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "無傳入任何參數" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 1. 如果有傳入 calendarId 參數，則下載並解析 Google 行事曆的公開 ICS 檔案
  if (e.parameter.calendarId) {
    return getGoogleCalendarEvents(e.parameter.calendarId);
  }
  
  // 2. 如果 type === 'daimoku'，則讀取「題目統計」分頁的紀錄
  if (e.parameter.type === 'daimoku') {
    var headers = ["id", "dept", "name", "minutes", "daimokuCount", "timestamp"];
    var data = readSheetData("題目統計", headers);
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 3. 預設：讀取「進度回報」分頁的會議與進度紀錄
  var headers = ["id", "meetingTime", "attendees", "dept", "progress", "content", "pendingItems", "other", "fileName", "timestamp"];
  var data = readSheetData("進度回報", headers);
  return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    
    // 檢查是否為「題目統計」操作
    if (params.type === 'daimoku') {
      var headers = ["id", "dept", "name", "minutes", "daimokuCount", "timestamp"];
      var sheet = getOrCreateSheet("題目統計", headers);
      
      if (params.action === 'delete') {
        var success = deleteRowById(sheet, params.id);
        return ContentService.createTextOutput(JSON.stringify({ status: success ? "success" : "not_found" }))
          .setMimeType(ContentService.MimeType.JSON);
      } else {
        // 新增題目統計資料
        var newRow = [];
        for (var i = 0; i < headers.length; i++) {
          var key = headers[i];
          var val = params[key] !== undefined ? params[key] : "";
          if (key === "minutes" || key === "daimokuCount") {
            val = Number(val);
          }
          newRow.push(val);
        }
        sheet.appendRow(newRow);
        return ContentService.createTextOutput(JSON.stringify({ status: "success", data: params }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    } else {
      // 預設：進度回報操作
      var headers = ["id", "meetingTime", "attendees", "dept", "progress", "content", "pendingItems", "other", "fileName", "timestamp"];
      var sheet = getOrCreateSheet("進度回報", headers);
      
      if (params.action === 'delete') {
        var success = deleteRowById(sheet, params.id);
        return ContentService.createTextOutput(JSON.stringify({ status: success ? "success" : "not_found" }))
          .setMimeType(ContentService.MimeType.JSON);
      } else {
        // 新增進度回報資料
        var newRow = [];
        for (var i = 0; i < headers.length; i++) {
          var key = headers[i];
          var val = params[key] !== undefined ? params[key] : "";
          if (key === "progress") {
            val = Number(val);
          }
          newRow.push(val);
        }
        sheet.appendRow(newRow);
        return ContentService.createTextOutput(JSON.stringify({ status: "success", data: params }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --- 輔助函式：取得或建立工作表頁面 ---
function getOrCreateSheet(sheetName, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    if (sheetName === "進度回報") {
      var sheets = ss.getSheets();
      if (sheets.length > 0) {
        sheet = sheets[0];
        var currentName = sheet.getName();
        // 如果第一個分頁是預設空白名稱，就直接重命名
        if (currentName === "工作表1" || currentName === "Sheet1" || currentName === "無標題試算表" || currentName.indexOf("Sheet") === 0) {
          sheet.setName("進度回報");
        }
      } else {
        sheet = ss.insertSheet("進度回報");
      }
    } else {
      sheet = ss.insertSheet(sheetName);
    }
  }
  
  // 若為剛新建的空白工作表，寫入表頭
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  
  return sheet;
}

// --- 輔助函式：讀取工作表資料 ---
function readSheetData(sheetName, headers) {
  var sheet = getOrCreateSheet(sheetName, headers);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }
  
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var result = [];
  
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var item = {};
    var hasData = false;
    for (var j = 0; j < headers.length; j++) {
      var val = row[j];
      item[headers[j]] = val !== undefined ? val : "";
      if (val !== "") {
        hasData = true;
      }
    }
    // 必須包含 id 欄位才加入結果
    if (hasData && item.id) {
      result.push(item);
    }
  }
  
  return result;
}

// --- 輔助函式：根據 ID 刪除該列資料 ---
function deleteRowById(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;
  
  // 假設 id 永遠在第一列
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 2); // Apps Script 列數以 1 開始，且要跳過第一列標頭
      return true;
    }
  }
  return false;
}

// --- 行事曆同步核心邏輯 ---
function getGoogleCalendarEvents(calendarId) {
  try {
    var url = "https://calendar.google.com/calendar/ical/" + encodeURIComponent(calendarId) + "/public/basic.ics";
    var response = UrlFetchApp.fetch(url);
    var icsContent = response.getContentText();
    var events = parseICS(icsContent);
    
    // 設定抓取時間範圍：2026年5月1日至2026年8月31日 (滿足從今天 5/23 開始呈現的需求)
    var startLimit = new Date("2026-05-01T00:00:00+08:00").getTime();
    var endLimit = new Date("2026-08-31T23:59:59+08:00").getTime();
    
    var filteredEvents = events.filter(function(ev) {
      var t = new Date(ev.startTime).getTime();
      return t >= startLimit && t <= endLimit;
    });
    
    // 按時間由近到遠排序 (由小到大)
    filteredEvents.sort(function(a, b) {
      return new Date(a.startTime) - new Date(b.startTime);
    });
    
    return ContentService.createTextOutput(JSON.stringify(filteredEvents))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 解析 ICS 格式
function parseICS(icsStr) {
  var events = [];
  var lines = icsStr.split(/\r?\n/);
  
  var currentEvent = null;
  var inEvent = false;

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    
    // 處理折行 (Folded lines: 以空格或 Tab 開頭的代表上一列的延續)
    while (i + 1 < lines.length && (lines[i + 1].indexOf(" ") === 0 || lines[i + 1].indexOf("\t") === 0)) {
      line += lines[i + 1].substring(1);
      i++;
    }

    if (line.indexOf("BEGIN:VEVENT") === 0) {
      currentEvent = {
        title: "",
        startTime: "",
        endTime: "",
        location: "高雄六龜國小",
        desc: "無描述資訊",
        owner: "主辦者"
      };
      inEvent = true;
    } else if (line.indexOf("END:VEVENT") === 0) {
      if (currentEvent) {
        events.push(currentEvent);
      }
      inEvent = false;
      currentEvent = null;
    } else if (inEvent && currentEvent) {
      var colonIdx = line.indexOf(":");
      if (colonIdx > -1) {
        var key = line.substring(0, colonIdx);
        var val = line.substring(colonIdx + 1);
        
        // 清除 iCal 中的字串逸出字元
        val = val.replace(/\\,/g, ",").replace(/\\;/g, ";").replace(/\\n/g, "\n").replace(/\\/g, "");
        
        if (key.indexOf("SUMMARY") === 0) {
          currentEvent.title = val;
        } else if (key.indexOf("LOCATION") === 0) {
          currentEvent.location = val || "高雄六龜國小";
        } else if (key.indexOf("DESCRIPTION") === 0) {
          currentEvent.desc = val || "無描述資訊";
        } else if (key.indexOf("DTSTART") === 0) {
          currentEvent.startTime = parseICSDate(val);
        } else if (key.indexOf("DTEND") === 0) {
          currentEvent.endTime = parseICSDate(val);
        }
      }
    }
  }
  
  return events;
}

// 解析 ICS 日期格式
function parseICSDate(icsDateStr) {
  if (icsDateStr.indexOf(":") > -1) {
    icsDateStr = icsDateStr.substring(icsDateStr.indexOf(":") + 1);
  }
  
  var cleanStr = icsDateStr.replace(/[^0-9TZ]/g, "");
  
  var y = cleanStr.substring(0, 4);
  var m = cleanStr.substring(4, 6);
  var d = cleanStr.substring(6, 8);
  
  if (cleanStr.indexOf("T") > -1) {
    var h = cleanStr.substring(9, 11);
    var min = cleanStr.substring(11, 13);
    var s = cleanStr.substring(13, 15);
    
    if (cleanStr.endsWith("Z")) {
      return new Date(Date.UTC(parseInt(y, 10), parseInt(m, 10) - 1, parseInt(d, 10), parseInt(h, 10), parseInt(min, 10), parseInt(s, 10))).toISOString();
    } else {
      // 本地時間，預設為台北時區 (+08:00)
      return new Date(y + "-" + m + "-" + d + "T" + h + ":" + min + ":" + s + "+08:00").toISOString();
    }
  } else {
    // 整天活動，預設加時區以防跨日偏移
    return new Date(y + "-" + m + "-" + d + "T00:00:00+08:00").toISOString();
  }
}

// 測試用函數：在 Apps Script 編輯器中選擇並手動點擊「執行」此函數，以觸發 Google Calendar 及外部連線權限授權
function testCalendar() {
  CalendarApp.getCalendarById("test");
  UrlFetchApp.fetch("https://www.google.com");
}
