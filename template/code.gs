// --- 請修改以下設定 ---
// 1. 如果您更換了新的試算表檔案，請務必更新為新的 SPREADSHEET_ID。
const SPREADSHEET_ID = '1wtE7ZNq2Oip9LxAtPynMqCP7i2SZxiSXM93fpeTFKlc';
// 2. 請將 '工作表1' 替換成您的回報資料所在的工作表名稱。
const SHEET_NAME = '工作表1';
// 3. 這是您的影城總表工作表名稱，請確保與您建立的相符。
const CINEMA_LIST_SHEET_NAME = '影城列表';
// --- 設定結束 ---

// --- 自訂地區排序 ---
const REGION_ORDER = ['北北基', '桃竹苗', '中彰投', '雲嘉南', '高屏', '宜花東', '離島'];
// --- 排序設定結束 ---

// --- 快取設定 (單位：秒) ---
const CACHE_EXPIRATION_SECONDS = 60;
// --- 快取設定結束 ---

/**
 * 當使用者透過 GET 請求訪問網頁應用程式時，執行此函數。
 * 現在這個函數會根據 action 參數來處理 API 請求，而不是返回 HTML。
 */
function doGet(e) {
  // 取得 URL 參數
  const params = e.parameter;
  const action = params.action;
  
  let result;
  
  try {
    // 根據 action 參數調用對應的函數
    switch(action) {
      case 'getAvailableWeeks':
        result = getAvailableWeeks();
        break;
      case 'getLastUpdateTime':
        result = getLastUpdateTime();
        break;
      case 'getDataForWeek':
        result = getDataForWeek(params.week);
        break;
      default:
        result = { error: '未知的 action 參數' };
    }
  } catch (error) {
    result = { error: error.toString() };
  }
  
  // 返回 JSON 格式的資料，並設定 CORS headers
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * [新增] 獲取最後一筆回報的提交時間。
 * @returns {string|null} 格式化後的時間字串，或在找不到時回傳 null。
 */
function getLastUpdateTime() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    // 如果工作表不存在或少於2列（只有標題列），則返回 null
    if (!sheet || sheet.getLastRow() < 2) {
      return null;
    }
    // 「Submission time」位於第 2 欄 (B欄)
    const lastTimestamp = sheet.getRange(sheet.getLastRow(), 2).getValue();
    
    if (lastTimestamp instanceof Date) {
      // 格式化日期為 YYYY/MM/DD HH:MM
      const year = lastTimestamp.getFullYear();
      const month = (lastTimestamp.getMonth() + 1).toString().padStart(2, '0');
      const day = lastTimestamp.getDate().toString().padStart(2, '0');
      const hours = lastTimestamp.getHours().toString().padStart(2, '0');
      const minutes = lastTimestamp.getMinutes().toString().padStart(2, '0');
      return `${year}/${month}/${day} ${hours}:${minutes}`;
    }
    return null;
  } catch (e) {
    Logger.log(`獲取最後更新時間失敗: ${e.message}`);
    // 在前端靜默失敗，不顯示錯誤訊息
    return null;
  }
}


/**
 * 獲取所有不重複的特典週次
 * @returns {Array<string>|Object} 一個包含所有週次的陣列，或是一個包含錯誤訊息的物件。
 */
function getAvailableWeeks() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`在試算表中找不到名為 "${SHEET_NAME}" 的工作表，請檢查名稱是否完全相符。`);
    }

    if (sheet.getLastRow() < 2) {
      return [];
    }
    
    const weekData = sheet.getRange(2, 11, sheet.getLastRow() - 1, 1).getValues(); 
    
    const weeks = new Set();
    weekData.forEach(row => {
      if (row[0]) {
        weeks.add(row[0].toString().trim());
      }
    });
    
    return Array.from(weeks).sort((a, b) => {
      const numA = parseInt(a.replace(/[^0-9]/g, ''), 10);
      const numB = parseInt(b.replace(/[^0-9]/g, ''), 10);
      return numA - numB;
    });

  } catch (e) {
    Logger.log(`獲取週次列表失敗: ${e.message}`);
    return { error: e.message };
  }
}

/**
 * [核心優化函數] 根據指定的週次，從快取或試算表中獲取處理過的資料。
 * @param {string} week - 使用者選擇的週次，例如 '第1週'。
 * @returns {Object} 一個包含所有地區、影城和該週次最新狀態的物件。
 */
function getProcessedDataForWeek_(week) {
  if (!week) return null;

  const cache = CacheService.getScriptCache();
  const CACHE_KEY = `WEEK_DATA_V2_${week}`;
  
  const cachedData = cache.get(CACHE_KEY);
  if (cachedData != null) {
    return JSON.parse(cachedData);
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const cinemaListSheet = ss.getSheetByName(CINEMA_LIST_SHEET_NAME);
  if (!cinemaListSheet) {
    throw new Error(`在試算表中找不到名為 "${CINEMA_LIST_SHEET_NAME}" 的工作表，請檢查名稱是否完全相符。`);
  }
  const cinemaListData = cinemaListSheet.getDataRange().getValues();
  
  const regions = new Set();
  const cinemasByRegion = {};
  
  for (let i = 1; i < cinemaListData.length; i++) {
    const region = cinemaListData[i][0];
    const cinema = cinemaListData[i][1];
    if (region && cinema) {
      regions.add(region);
      if (!cinemasByRegion[region]) cinemasByRegion[region] = [];
      cinemasByRegion[region].push(cinema);
    }
  }

  const responseSheet = ss.getSheetByName(SHEET_NAME);
  if (!responseSheet) {
    throw new Error(`在試算表中找不到名為 "${SHEET_NAME}" 的工作表，請檢查名稱是否完全相符。`);
  }
  const responseData = responseSheet.getDataRange().getValues();
  
  const theaterReports = {};

  const COL_WEEK = 10;
  const COL_PICKUP_TIME = 11;
  const COL_STATUS = 12;

  for (let i = 1; i < responseData.length; i++) {
    const row = responseData[i];
    const rowWeek = row[COL_WEEK] ? row[COL_WEEK].toString().trim() : '';

    if (rowWeek === week) {
      let theaterName = '';
      for (let j = 3; j <= 9; j++) {
        if (row[j]) {
          theaterName = row[j].toString().trim();
          break;
        }
      }

      const pickupTime = row[COL_PICKUP_TIME];
      
      if (theaterName && pickupTime) {
        if (!theaterReports[theaterName]) {
          theaterReports[theaterName] = [];
        }
        
        const statusResponse = row[COL_STATUS] ? row[COL_STATUS].toString() : '';
        const hasStock = statusResponse.includes('是') || statusResponse.includes('有');
        const statusSymbol = hasStock ? '⭕️' : '❌';
        
        theaterReports[theaterName].push({
          time: new Date(pickupTime),
          status: statusSymbol
        });
      }
    }
  }

  const latestStatus = {};
  for (const theaterName in theaterReports) {
    const sortedReports = theaterReports[theaterName].sort((a, b) => b.time.getTime() - a.time.getTime());
    
    if (sortedReports.length > 0) {
      const latestReport = sortedReports[0];
      const timestamp = latestReport.time;
      const hours = timestamp.getHours().toString().padStart(2, '0');
      const minutes = timestamp.getMinutes().toString().padStart(2, '0');
      const timeString = `${hours}:${minutes}`;
      latestStatus[theaterName] = `${timeString} ${latestReport.status}`;
    }
  }

  const sortedRegions = Array.from(regions).sort((a, b) => {
    const indexA = REGION_ORDER.indexOf(a);
    const indexB = REGION_ORDER.indexOf(b);
    if (indexA === -1) return 1;
    if (indexB === -1) return -1;
    return indexA - indexB;
  });

  const processedData = {
    regions: sortedRegions,
    cinemasByRegion: cinemasByRegion,
    latestStatus: latestStatus
  };

  cache.put(CACHE_KEY, JSON.stringify(processedData), CACHE_EXPIRATION_SECONDS);
  return processedData;
}

/**
 * 獲取指定週次的所有地區及其影城的最新狀態。
 * @param {string} week - 使用者選擇的週次。
 * @returns {Array<Object>|Object} 整理好的資料結構，或是一個包含錯誤訊息的物件。
 */
function getDataForWeek(week) {
  try {
    const data = getProcessedDataForWeek_(week);
    if (!data) return [];
    
    if (data.error) {
        return data;
    }
    
    const allRegions = data.regions;
    
    return allRegions.map(region => {
      const cinemasInRegion = data.cinemasByRegion[region] || [];
      cinemasInRegion.sort((a, b) => a.localeCompare(b, 'zh-Hant'));
      
      const cinemaStatuses = cinemasInRegion.map(cinemaName => ({
        theater: cinemaName,
        status: data.latestStatus[cinemaName] || ''
      }));

      return { region: region, cinemas: cinemaStatuses };
    });

  } catch (e) {
    Logger.log(`獲取週次資料失敗 (${week}): ${e.message}`);
    return { error: e.message };
  }
}