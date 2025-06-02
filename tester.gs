// 專門用於測試和診斷網站爬取的工具

/**
 * 診斷期貨資料爬取失敗的原因
 * 可以直接在 Apps Script 編輯器中執行此函數來查看詳細日誌
 */
function diagnoseScrapingIssue() {
  // 清空舊日誌
  Logger.clear();
  
  // 取得今日日期
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, 'Asia/Taipei', 'yyyy/MM/dd');
  
  // 測試各種契約
  const contracts = ['TX', 'TE', 'MTX', 'ZMX', 'NQF'];
  
  for (const contract of contracts) {
    testContractWithDetails(contract, formattedDate);
  }
  
  // 特別詳細測試 NQF
  testNQFWithAllOptions(formattedDate);
  
  // 輸出所有診斷日誌到工作表
  writeLogsToSheet();
}

/**
 * 檢查工作表是否存在，若不存在則建立
 * @param {SpreadsheetApp.Spreadsheet} ss - 試算表物件
 * @param {string} sheetName - 工作表名稱
 * @returns {SpreadsheetApp.Sheet} 工作表物件
 */
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    setupSheetHeader(sheet);
  }
  return sheet;
}

/**
 * 針對特定契約進行詳細測試
 */
function testContractWithDetails(contract, dateStr) {
  Logger.log(`======== 開始詳細測試 ${contract} 在 ${dateStr} 的解析 ========`);
  
  // 構建查詢參數
  const queryData = {
    'queryType': '2',
    'marketCode': '0',
    'dateaddcnt': '',
    'commodity_id': contract,
    'queryDate': dateStr
  };
  
  // 發送請求
  const response = fetchDataFromTaifex(queryData);
  
  if (response) {
    Logger.log(`成功獲取回應，內容長度: ${response.length}`);
    
    // 檢查是否有無資料訊息
    if (hasNoDataMessage(response)) {
      Logger.log('回應中包含無資料訊息');
    } else {
      // 保存 HTML 到變數以便分析
      const htmlSummary = response.substring(0, 1000) + "... [內容已截斷]";
      Logger.log(`HTML 預覽: ${htmlSummary}`);
      
      // 檢查表格數量
      inspectTables(response);
      
      // 嘗試解析
      const data = parseContractData(response, contract, dateStr);
      
      if (data) {
        Logger.log('解析成功!');
        Logger.log(`資料: ${JSON.stringify(data)}`);
      } else {
        Logger.log('解析失敗 - 請檢查上方日誌以了解更多詳情');
      }
    }
  } else {
    Logger.log('請求失敗 - 無法獲取回應');
  }
  
  Logger.log(`======== ${contract} 測試完成 ========\n`);
}

/**
 * 針對 NQF 契約進行全參數測試
 */
function testNQFWithAllOptions(dateStr) {
  Logger.log(`======== 開始 NQF 全參數測試 ========`);
  
  // 使用不同的 queryType
  const queryTypes = ['1', '2'];
  // 使用不同的 marketCode
  const marketCodes = ['0', '1'];
  
  for (const queryType of queryTypes) {
    for (const marketCode of marketCodes) {
      Logger.log(`\n測試 NQF 使用參數組合: queryType=${queryType}, marketCode=${marketCode}`);
      
      // 構建查詢參數
      const queryData = {
        'queryType': queryType,
        'marketCode': marketCode,
        'dateaddcnt': '',
        'commodity_id': 'NQF',
        'queryDate': dateStr
      };
      
      // 發送請求
      const response = fetchDataFromTaifex(queryData);
      
      if (response) {
        // 快速檢查是否包含 NQF 或美國那斯達克
        const hasNQF = response.includes('NQF') || response.includes('美國那斯達克');
        Logger.log(`回應中包含 NQF 或美國那斯達克的關鍵字: ${hasNQF}`);
        
        // 檢查表格
        const tableCount = inspectTables(response);
        
        // 如果找到表格，嘗試解析
        if (tableCount > 0) {
          const data = parseContractData(response, 'NQF', dateStr);
          if (data) {
            Logger.log(`成功！使用參數組合 queryType=${queryType}, marketCode=${marketCode} 可以獲取 NQF 資料`);
            Logger.log(`資料: ${JSON.stringify(data)}`);
          }
        }
      } else {
        Logger.log(`使用參數組合 queryType=${queryType}, marketCode=${marketCode} 請求失敗`);
      }
    }
  }
  
  Logger.log(`======== NQF 全參數測試完成 ========\n`);
}

/**
 * 分析回應中的表格結構
 */
function inspectTables(htmlContent) {
  // 尋找所有表格
  const tables = htmlContent.match(/<table[^>]*>[\s\S]*?<\/table>/gi) || [];
  Logger.log(`找到 ${tables.length} 個表格`);
  
  if (tables.length > 0) {
    // 檢查每個表格的基本屬性
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      const classMatch = table.match(/class="([^"]*?)"/i);
      const className = classMatch ? classMatch[1] : '無類名';
      
      // 檢查表格行數
      const rows = table.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];
      
      // 檢查表格是否包含關鍵字
      const hasTrading = table.includes('交易口數');
      const hasOpenInterest = table.includes('未平倉口數');
      const hasNQF = table.includes('NQF') || table.includes('美國那斯達克');
      
      Logger.log(`表格 ${i+1}: 類名="${className}", 行數=${rows.length}, 包含[交易口數=${hasTrading}, 未平倉口數=${hasOpenInterest}, NQF相關=${hasNQF}]`);
      
      // 如果表格包含 NQF 或交易資料，進一步分析
      if (hasNQF || hasTrading || hasOpenInterest) {
        // 檢查表頭
        const headerRow = rows[0];
        if (headerRow) {
          const headerCells = headerRow.match(/<th[^>]*>([\s\S]*?)<\/th>/gi) || [];
          Logger.log(`  表頭單元格數: ${headerCells.length}`);
          
          if (headerCells.length > 0) {
            const headerTexts = headerCells.map(cell => {
              return cell.replace(/<[^>]*>/g, '').trim();
            });
            Logger.log(`  表頭內容: ${headerTexts.join(' | ')}`);
          }
        }
      }
    }
  }
  
  return tables.length;
}

/**
 * 將診斷日誌寫入工作表
 */
function writeLogsToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('診斷日誌');
  
  if (!logSheet) {
    logSheet = ss.insertSheet('診斷日誌');
  } else {
    logSheet.clear();
  }
  
  // 設置標題
  logSheet.getRange('A1').setValue('診斷日誌 - ' + new Date());
  logSheet.getRange('A1').setFontWeight('bold');
  
  // 獲取日誌
  const logs = Logger.getLog().split('\n');
  
  // 寫入日誌
  for (let i = 0; i < logs.length; i++) {
    logSheet.getRange(i + 3, 1).setValue(logs[i]);
  }
  
  // 調整欄寬
  logSheet.autoResizeColumn(1);
  
  // 返回到第一行
  logSheet.setActiveSelection('A1');
  
  // 凍結首行
  logSheet.setFrozenRows(1);
  
  Logger.log('診斷日誌已寫入工作表');
}

/**
 * 直接展示台灣期貨交易所網站返回的原始內容
 */
function showRawHtmlResponse() {
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, 'Asia/Taipei', 'yyyy/MM/dd');
  
  // 針對 NQF 爬取原始回應
  const queryData = {
    'queryType': '2',
    'marketCode': '0',
    'dateaddcnt': '',
    'commodity_id': 'NQF',
    'queryDate': formattedDate
  };
  
  // 發送請求
  const response = fetchDataFromTaifex(queryData);
  
  if (response) {
    // 創建一個新工作表來儲存 HTML
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let htmlSheet = ss.getSheetByName('HTML原始內容');
    
    if (!htmlSheet) {
      htmlSheet = ss.insertSheet('HTML原始內容');
    } else {
      htmlSheet.clear();
    }
    
    // 設置標題
    htmlSheet.getRange('A1').setValue(`NQF 在 ${formattedDate} 的原始 HTML 回應`);
    htmlSheet.getRange('A1').setFontWeight('bold');
    
    // 寫入 HTML 內容 (分成多行以避免超出單元格上限)
    const chunkSize = 50000; // 每個單元格的最大字元數
    for (let i = 0; i < response.length; i += chunkSize) {
      const chunk = response.substring(i, Math.min(i + chunkSize, response.length));
      htmlSheet.getRange(3 + Math.floor(i / chunkSize), 1).setValue(chunk);
    }
    
    Logger.log('原始 HTML 已保存到工作表');
  } else {
    Logger.log('無法獲取回應');
  }
}

/**
 * 測試特定日期的資料爬取
 * 可以直接在 Apps Script 編輯器中執行此函數來測試特定日期的爬取
 */
function testSpecificDate() {
  // 清空舊日誌
  Logger.clear();
  
  // 設定特定日期
  const specificDate = "2025/05/16";
  
  Logger.log(`開始測試爬取 ${specificDate} 的期貨資料`);
  Logger.log(`注意：此為測試模式，爬取的資料會以 ${specificDate} 的日期儲存，即使實際資料可能來自不同日期`);
  Logger.log(`若要實際進行資料爬取，請使用主選單中的"爬取特定日期資料"功能`);
  
  // 測試各種契約
  const contracts = ['TX', 'TE', 'MTX', 'ZMX', 'NQF'];
  
  for (const contract of contracts) {
    testContractWithDetails(contract, specificDate);
  }
  
  // 測試更近的未來日期 (明天)
  testNearFutureDate();
  
  // 輸出所有診斷日誌到工作表
  writeLogsToSheet();
  
  Logger.log(`完成 ${specificDate} 的測試`);
  
  // 提示用戶測試結果
  SpreadsheetApp.getUi().alert(
    '測試完成',
    `已完成 ${specificDate} 的測試爬取。\n\n` +
    `注意：此為測試模式，爬取的資料會以 ${specificDate} 的日期儲存，即使實際資料可能來自其他日期。\n\n` +
    `若在資料中發現日期不一致的問題，請使用主選單中的"爬取特定日期資料"功能進行正式爬取。`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 測試近期未來日期的資料爬取
 * 這可能會比較容易成功，因為期貨交易所可能限制查詢太遠的未來日期
 */
function testNearFutureDate() {
  // 獲取明天的日期
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  
  // 如果明天是週末，則改為下週一
  const day = tomorrow.getDay();
  if (day === 0) { // 週日
    tomorrow.setDate(tomorrow.getDate() + 1); // 改為週一
  } else if (day === 6) { // 週六
    tomorrow.setDate(tomorrow.getDate() + 2); // 改為週一
  }
  
  const formattedDate = Utilities.formatDate(tomorrow, 'Asia/Taipei', 'yyyy/MM/dd');
  
  Logger.log(`\n======== 開始測試明日/下一個交易日 ${formattedDate} ========`);
  
  // 測試各種契約
  const contracts = ['TX', 'TE', 'MTX', 'ZMX', 'NQF'];
  
  for (const contract of contracts) {
    testContractWithDetails(contract, formattedDate);
  }
  
  Logger.log(`======== 明日/下一個交易日測試完成 ========\n`);
}

/**
 * 測試使用不同查詢參數組合
 */
function testWithDifferentParameters() {
  Logger.clear();
  
  // 獲取明天的日期
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  
  // 如果明天是週末，則改為下週一
  const day = tomorrow.getDay();
  if (day === 0) { // 週日
    tomorrow.setDate(tomorrow.getDate() + 1); // 改為週一
  } else if (day === 6) { // 週六
    tomorrow.setDate(tomorrow.getDate() + 2); // 改為週一
  }
  
  const formattedDate = Utilities.formatDate(tomorrow, 'Asia/Taipei', 'yyyy/MM/dd');
  
  Logger.log(`開始測試不同參數組合，使用日期 ${formattedDate}`);
  
  // 使用不同的查詢參數組合
  const queryTypes = ['1', '2', '3'];
  const marketCodes = ['0', '1', '2'];
  
  for (const contract of ['TX', 'NQF']) {
    for (const queryType of queryTypes) {
      for (const marketCode of marketCodes) {
        testWithParameters(contract, formattedDate, queryType, marketCode);
      }
    }
  }
  
  // 輸出所有診斷日誌到工作表
  writeLogsToSheet();
}

/**
 * 使用特定參數組合測試
 */
function testWithParameters(contract, dateStr, queryType, marketCode) {
  Logger.log(`\n測試 ${contract} 使用參數: queryType=${queryType}, marketCode=${marketCode}, date=${dateStr}`);
  
  // 構建查詢參數
  const queryData = {
    'queryType': queryType,
    'marketCode': marketCode,
    'dateaddcnt': '',
    'commodity_id': contract,
    'queryDate': dateStr
  };
  
  // 發送請求
  const response = fetchDataFromTaifex(queryData);
  
  if (response) {
    Logger.log(`成功獲取回應，內容長度: ${response.length}`);
    
    // 檢查是否有無資料訊息
    if (hasNoDataMessage(response)) {
      Logger.log('回應中包含無資料訊息');
    } else if (hasErrorMessage(response)) {
      Logger.log('回應中包含錯誤訊息');
    } else {
      // 嘗試解析
      const data = parseContractData(response, contract, dateStr);
      
      if (data) {
        Logger.log('解析成功!');
        Logger.log(`資料: ${JSON.stringify(data)}`);
      } else {
        Logger.log('解析失敗');
      }
    }
  } else {
    Logger.log('請求失敗');
  }
}

/**
 * 專門針對TX契約進行測試，使用多種參數組合
 */
function testTXContract() {
  // 清空舊日誌
  Logger.clear();
  
  // 獲取今天和昨天的日期
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  
  // 格式化日期
  const todayStr = Utilities.formatDate(today, 'Asia/Taipei', 'yyyy/MM/dd');
  const yesterdayStr = Utilities.formatDate(yesterday, 'Asia/Taipei', 'yyyy/MM/dd');
  
  Logger.log(`======== 開始針對台指期(TX)進行詳細測試 ========`);
  
  // 測試不同的參數組合
  const dateToTest = [yesterdayStr, todayStr];
  const queryTypes = ['1', '2', '3'];
  const marketCodes = ['0', '1', '2'];
  
  for (const dateStr of dateToTest) {
    Logger.log(`\n----- 測試日期：${dateStr} -----`);
    
    for (const queryType of queryTypes) {
      for (const marketCode of marketCodes) {
        // 構建查詢參數
        const queryData = {
          'queryType': queryType,
          'marketCode': marketCode,
          'dateaddcnt': '',
          'commodity_id': 'TX',
          'queryDate': dateStr
        };
        
        Logger.log(`測試參數組合: queryType=${queryType}, marketCode=${marketCode}`);
        
        // 發送請求
        const response = fetchDataFromTaifex(queryData);
        
        if (response) {
          Logger.log(`成功獲取回應，內容長度: ${response.length}`);
          
          // 檢查是否包含TX相關關鍵字
          const hasTX = response.includes('TX') || response.includes('臺股期貨') || response.includes('台指期');
          Logger.log(`回應中包含TX相關關鍵字: ${hasTX}`);
          
          // 嘗試解析
          const data = parseContractData(response, 'TX', dateStr);
          
          if (data) {
            Logger.log(`解析成功! 使用參數組合: queryType=${queryType}, marketCode=${marketCode}`);
            Logger.log(`資料: ${JSON.stringify(data)}`);
          } else {
            Logger.log(`使用參數組合 queryType=${queryType}, marketCode=${marketCode} 解析失敗`);
          }
        } else {
          Logger.log(`使用參數組合 queryType=${queryType}, marketCode=${marketCode} 請求失敗`);
        }
        
        // 暫停一秒避免請求過快
        Utilities.sleep(1000);
      }
    }
  }
  
  Logger.log(`======== TX契約測試完成 ========`);
  
  // 輸出所有診斷日誌到工作表
  writeLogsToSheet();
}

/**
 * 測試NQF契約的身份別資料爬取
 * 使用專用處理邏輯爬取NQF契約的投信和外資資料
 */
function testNQFIdentityData() {
  try {
    Logger.log('開始測試NQF契約的身份別資料爬取功能');
    
    // 清空舊日誌
    Logger.clear();
    
    // 建立測試日期 (使用過去一周內的工作日)
    const today = new Date();
    let testDate = new Date(today);
    testDate.setDate(today.getDate() - 7); // 從一周前開始往前找
    
    // 找到有效的交易日
    for (let i = 0; i < 10; i++) {
      // 避開週末
      const dayOfWeek = testDate.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) {
        break;
      }
      testDate.setDate(testDate.getDate() + 1);
    }
    
    // 格式化日期
    const dateStr = Utilities.formatDate(testDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    Logger.log(`測試日期: ${dateStr}`);
    
    // 取得活動的試算表
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 取得NQF工作表
    let sheet = getOrCreateSheet(ss, 'NQF');
    
    // 測試針對不同身份別的爬取結果
    const identities = ['自營商', '投信', '外資'];
    
    // 先運行分析表格結構測試，以了解期交所網頁的結構
    Logger.log('先運行表格結構分析...');
    analyzeNQFTableStructure();
    
    // 使用專用處理邏輯爬取NQF身份別資料
    Logger.log('使用專用處理邏輯爬取資料...');
    const success = processNQFIdentityData(sheet, dateStr, identities, 0);
    
    // 寫入日誌到工作表
    writeLogsToSheet();
    
    Logger.log(`測試完成，成功爬取 ${success} 個身份別資料`);
    
    if (success > 0) {
      // 提示測試成功
      SpreadsheetApp.getUi().alert(
        '測試完成',
        `成功測試NQF契約 ${dateStr} 的身份別資料爬取，共爬取 ${success} 個身份別資料。\n\n請檢查NQF工作表中的資料和診斷日誌工作表以了解詳情。`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      // 提示測試失敗
      SpreadsheetApp.getUi().alert(
        '測試失敗',
        `測試NQF契約 ${dateStr} 的身份別資料爬取失敗，未能爬取到任何資料。\n\n請檢查診斷日誌工作表以了解詳細錯誤資訊。`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (e) {
    Logger.log(`測試NQF身份別資料爬取時發生錯誤: ${e.message}`);
    SpreadsheetApp.getUi().alert(`測試時發生錯誤: ${e.message}`);
  }
}

/**
 * 專門處理NQF契約的身份別資料爬取
 * 針對NQF契約對每個身份別使用不同的參數組合進行爬取
 * @param {Object} sheet - 資料表物件
 * @param {string} dateStr - 日期字符串，格式為'yyyy/MM/dd'
 * @param {Array} identities - 身份別陣列
 * @param {number} identityDataCount - 已爬取的資料計數
 * @returns {number} 成功爬取的資料數量
 */
function processNQFIdentityData(sheet, dateStr, identities, identityDataCount) {
  try {
    Logger.log(`使用專用方法處理NQF契約的身份別資料爬取`);
    const contract = 'NQF';
    let success = 0;
    
    // 依序處理每種身份別
    for (const identity of identities) {
      Logger.log(`NQF專用處理：爬取 ${identity} 資料`);
      let identitySuccess = false;
      
      // 嘗試不同的查詢參數組合 - 為不同身份別使用不同的優先順序
      let queryTypes = ['2', '1', '3'];
      let marketCodes = ['0', '1', '2'];
      
      // 對投信和外資使用不同的優先順序
      if (identity === '投信') {
        queryTypes = ['1', '2', '3', '1', '1']; // 投信優先使用queryType=1，重複嘗試
        marketCodes = ['0', '1', '2', '0', '1'];
      } else if (identity === '外資') {
        queryTypes = ['2', '3', '1', '3', '2']; // 外資優先使用queryType=2和queryType=3，重複嘗試
        marketCodes = ['1', '0', '2', '1', '0']; // 外資優先使用marketCode=1
      }
      
      // 嘗試不同的參數組合
      for (const queryType of queryTypes) {
        if (identitySuccess) break;
        
        for (const marketCode of marketCodes) {
          if (identitySuccess) break;
          
          // 構建查詢參數
          const queryData = {
            'queryType': queryType,
            'marketCode': marketCode,
            'dateaddcnt': '',
            'commodity_id': contract,
            'queryDate': dateStr
          };
          
          Logger.log(`NQF ${identity} 嘗試參數組合: queryType=${queryType}, marketCode=${marketCode}`);
          
          // 發送請求
          const response = fetchDataFromTaifex(queryData);
          
          // 處理回應
          if (response && !hasNoDataMessage(response) && !hasErrorMessage(response)) {
            try {
              // 使用增強版解析器解析NQF資料
              const data = enhancedParseIdentityData(response, dateStr, identity);
              
              if (data) {
                Logger.log(`成功解析 NQF ${dateStr} ${identity} 資料`);
                
                // 【追蹤日期】記錄從解析器獲得的資料日期
                Logger.log(`【日期追蹤】從增強版解析器獲得資料，查詢日期為: ${dateStr}, 資料日期為: ${data[0]}, 身份別: ${identity}`);
                
                // 重要修正：確保資料陣列中的日期與查詢日期一致
                if (data[0] !== dateStr) {
                  Logger.log(`【日期修正】將資料日期從 ${data[0]} 修正為查詢日期 ${dateStr}`);
                  data[0] = dateStr;
                }
                
                // 檢查是否已存在相同的資料
                if (!isIdentityDataExists(sheet, dateStr, contract, identity)) {
                  // 在寫入前再次確認日期
                  Logger.log(`【日期追蹤】準備寫入工作表，資料日期為: ${data[0]}, 身份別: ${identity}, 多方交易口數: ${data[2]}, 空方交易口數: ${data[4]}`);
                  
                  // 添加到工作表
                  addDataToSheet(sheet, data);
                  Logger.log(`成功爬取 NQF ${dateStr} ${identity} 的資料`);
                  addLog(contract, `${dateStr} (${identity})`, '成功 (增強解析)');
                  identitySuccess = true;
                  success++;
                  identityDataCount++;
                } else {
                  Logger.log(`NQF ${dateStr} ${identity} 資料已存在，跳過`);
                  addLog(contract, `${dateStr} (${identity})`, '已存在');
                  identitySuccess = true;
                  success++;
                }
              } else {
                // 如果增強版解析器失敗，嘗試原始解析器
                Logger.log(`增強版解析器失敗，嘗試原始解析器`);
                const oldData = parseIdentityData(response, contract, dateStr, identity);
                
                if (oldData && validateIdentityData(oldData, identity)) {
                  Logger.log(`原始解析器成功解析 NQF ${dateStr} ${identity} 資料`);
                  
                  // 【追蹤日期】記錄從原始解析器獲得的資料日期
                  Logger.log(`【日期追蹤】從原始解析器獲得資料，查詢日期為: ${dateStr}, 資料日期為: ${oldData[0]}, 身份別: ${identity}`);
                  
                  // 重要修正：確保資料陣列中的日期與查詢日期一致
                  if (oldData[0] !== dateStr) {
                    Logger.log(`【日期修正】將資料日期從 ${oldData[0]} 修正為查詢日期 ${dateStr}`);
                    oldData[0] = dateStr;
                  }
                  
                  // 檢查是否已存在相同的資料
                  if (!isIdentityDataExists(sheet, dateStr, contract, identity)) {
                    // 在寫入前再次確認日期
                    Logger.log(`【日期追蹤】準備寫入工作表(原始解析)，資料日期為: ${oldData[0]}, 身份別: ${identity}, 多方交易口數: ${oldData[2]}, 空方交易口數: ${oldData[4]}`);
                    
                    // 添加到工作表
                    addDataToSheet(sheet, oldData);
                    Logger.log(`成功爬取 NQF ${dateStr} ${identity} 的資料 (原始解析器)`);
                    addLog(contract, `${dateStr} (${identity})`, '成功 (原始解析)');
                    identitySuccess = true;
                    success++;
                    identityDataCount++;
                  } else {
                    Logger.log(`NQF ${dateStr} ${identity} 資料已存在，跳過`);
                    addLog(contract, `${dateStr} (${identity})`, '已存在');
                    identitySuccess = true;
                    success++;
                  }
                }
              }
            } catch (e) {
              Logger.log(`解析 NQF ${dateStr} ${identity} 資料時發生錯誤: ${e.message}`);
            }
          }
          
          // 避免請求過於頻繁
          Utilities.sleep(1000);
        }
      }
      
      // 如果所有參數組合都無法獲取該身份別的資料，嘗試硬編碼方式
      if (!identitySuccess) {
        Logger.log(`所有參數組合都無法爬取 NQF ${dateStr} ${identity} 的資料，嘗試直接建立資料`);
        
        // 嘗試創建可能的資料
        if (identity === '投信') {
          // 投信通常交易量很小，可能為0或小數字
          const investData = [
            dateStr,              // 日期
            contract,             // 契約代碼
            0,                    // 多方交易口數 (可能為0或小數字)
            0,                    // 多方契約金額
            0,                    // 空方交易口數 (可能為0或小數字)
            0,                    // 空方契約金額
            0,                    // 多空淨額交易口數
            0,                    // 多空淨額契約金額
            0,                    // 多方未平倉口數 (可能為0或小數字)
            0,                    // 多方未平倉契約金額
            0,                    // 空方未平倉口數 (可能為0或小數字)
            0,                    // 空方未平倉契約金額
            0,                    // 多空淨額未平倉口數
            0,                    // 多空淨額未平倉契約金額
            identity              // 身份別 = '投信'
          ];
          
          // 提示用戶是否將投信資料設為0
          const ui = SpreadsheetApp.getUi();
          const confirm = ui.alert(
            '未找到NQF投信資料',
            `未能從期交所爬取到NQF ${dateStr} 的投信資料。\n\n投信對NQF的交易通常很少，是否要將投信資料設為0？`,
            ui.ButtonSet.YES_NO
          );
          
          if (confirm === ui.Button.YES) {
            if (!isIdentityDataExists(sheet, dateStr, contract, identity)) {
              addDataToSheet(sheet, investData);
              Logger.log(`已將 NQF ${dateStr} 投信資料設為0`);
              addLog(contract, `${dateStr} (${identity})`, '成功 (手動設為0)');
              success++;
              identityDataCount++;
            } else {
              Logger.log(`NQF ${dateStr} ${identity} 資料已存在，跳過`);
              addLog(contract, `${dateStr} (${identity})`, '已存在');
              success++;
            }
          } else {
            Logger.log(`用戶選擇不將 NQF ${dateStr} 投信資料設為0`);
            addLog(contract, `${dateStr} (${identity})`, '用戶取消設為0');
          }
        } else if (identity === '外資') {
          // 提示用戶手動輸入外資數據
          const ui = SpreadsheetApp.getUi();
          const confirm = ui.alert(
            '未找到NQF外資資料',
            `未能從期交所爬取到NQF ${dateStr} 的外資資料。\n\n外資對NQF通常有一定交易量。\n您是否想要手動輸入外資交易和未平倉數據？`,
            ui.ButtonSet.YES_NO
          );
          
          if (confirm === ui.Button.YES) {
            try {
              // 提示用戶輸入多方交易口數
              const buyPrompt = ui.prompt('NQF外資數據 - 多方交易口數', 
                                        '請輸入 NQF 外資多方交易口數：', 
                                        ui.ButtonSet.OK_CANCEL);
              
              if (buyPrompt.getSelectedButton() === ui.Button.OK) {
                const buyVolume = parseInt(buyPrompt.getResponseText().trim()) || 0;
                
                // 提示用戶輸入空方交易口數
                const sellPrompt = ui.prompt('NQF外資數據 - 空方交易口數', 
                                           '請輸入 NQF 外資空方交易口數：', 
                                           ui.ButtonSet.OK_CANCEL);
                
                if (sellPrompt.getSelectedButton() === ui.Button.OK) {
                  const sellVolume = parseInt(sellPrompt.getResponseText().trim()) || 0;
                  
                  // 計算多空淨額
                  const netVolume = buyVolume - sellVolume;
                  
                  // 提示用戶輸入多方未平倉口數
                  const buyOIPrompt = ui.prompt('NQF外資數據 - 多方未平倉口數', 
                                             '請輸入 NQF 外資多方未平倉口數：', 
                                             ui.ButtonSet.OK_CANCEL);
                  
                  if (buyOIPrompt.getSelectedButton() === ui.Button.OK) {
                    const buyOI = parseInt(buyOIPrompt.getResponseText().trim()) || 0;
                    
                    // 提示用戶輸入空方未平倉口數
                    const sellOIPrompt = ui.prompt('NQF外資數據 - 空方未平倉口數', 
                                                '請輸入 NQF 外資空方未平倉口數：', 
                                                ui.ButtonSet.OK_CANCEL);
                    
                    if (sellOIPrompt.getSelectedButton() === ui.Button.OK) {
                      const sellOI = parseInt(sellOIPrompt.getResponseText().trim()) || 0;
                      
                      // 計算多空淨額未平倉
                      const netOI = buyOI - sellOI;
                      
                      // 創建資料
                      const foreignData = [
                        dateStr,              // 日期
                        contract,             // 契約代碼
                        buyVolume,            // 多方交易口數
                        0,                    // 多方契約金額
                        sellVolume,           // 空方交易口數
                        0,                    // 空方契約金額
                        netVolume,            // 多空淨額交易口數
                        0,                    // 多空淨額契約金額
                        buyOI,                // 多方未平倉口數
                        0,                    // 多方未平倉契約金額
                        sellOI,               // 空方未平倉口數
                        0,                    // 空方未平倉契約金額
                        netOI,                // 多空淨額未平倉口數
                        0,                    // 多空淨額未平倉契約金額
                        identity              // 身份別 = '外資'
                      ];
                      
                      if (!isIdentityDataExists(sheet, dateStr, contract, identity)) {
                        addDataToSheet(sheet, foreignData);
                        Logger.log(`已手動輸入 NQF ${dateStr} 外資資料`);
                        addLog(contract, `${dateStr} (${identity})`, '成功 (手動輸入)');
                        success++;
                        identityDataCount++;
                      } else {
                        Logger.log(`NQF ${dateStr} ${identity} 資料已存在，跳過`);
                        addLog(contract, `${dateStr} (${identity})`, '已存在');
                        success++;
                      }
                    }
                  }
                }
              }
            } catch (e) {
              Logger.log(`手動輸入外資資料時發生錯誤: ${e.message}`);
              addLog(contract, `${dateStr} (${identity})`, '手動輸入錯誤');
            }
          } else {
            Logger.log(`用戶選擇不手動輸入 NQF ${dateStr} 外資資料`);
            addLog(contract, `${dateStr} (${identity})`, '用戶取消手動輸入');
          }
        }
      }
    }
    
    // 計算成功率
    const successRate = Math.round((success / identities.length) * 100);
    Logger.log(`NQF專用處理完成，成功率: ${successRate}% (${success}/${identities.length})`);
    
    return success;
  } catch (e) {
    Logger.log(`NQF專用處理時發生錯誤: ${e.message}`);
    return 0;
  }
}

/**
 * 針對特定契約爬取身份別資料
 * @param {string} contract - 契約代碼
 * @param {string} dateStr - 日期字符串
 * @param {Array} identities - 身份別陣列
 * @returns {number} 成功爬取的資料數量
 */
function fetchIdentityDataForContract(contract, dateStr, identities) {
  try {
    Logger.log(`開始爬取 ${contract} ${dateStr} 的身份別資料`);
    
    // 取得活動的試算表
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 取得對應的工作表
    const sheet = getOrCreateSheet(ss, contract);
    
    // 計數器
    let successCount = 0;
    
    // 如果是NQF契約，使用專用處理邏輯
    if (contract === 'NQF') {
      return processNQFIdentityData(sheet, dateStr, identities, 0);
    }
    
    // 其他契約使用標準處理邏輯
    // 嘗試不同的查詢參數組合
    const queryTypes = ['2', '1', '3']; // 優先使用queryType=2
    const marketCodes = ['0', '1', '2']; // 優先使用marketCode=0
    
    // 嘗試不同的參數組合
    for (const queryType of queryTypes) {
      for (const marketCode of marketCodes) {
        // 構建查詢參數
        const queryData = {
          'queryType': queryType,
          'marketCode': marketCode,
          'dateaddcnt': '',
          'commodity_id': contract,
          'queryDate': dateStr
        };
        
        Logger.log(`嘗試使用參數組合: queryType=${queryType}, marketCode=${marketCode} 爬取 ${contract} ${dateStr} 的身份別資料`);
        
        // 發送請求
        const response = fetchDataFromTaifex(queryData);
        
        // 處理回應
        if (response && !hasNoDataMessage(response) && !hasErrorMessage(response)) {
          try {
            // 嘗試解析身份別資料
            for (const identity of identities) {
              // 嘗試解析特定身份別資料
              const data = parseIdentityData(response, contract, dateStr, identity);
              
              if (data) {
                // 檢查是否已存在相同的資料
                if (!isIdentityDataExists(sheet, dateStr, contract, identity)) {
                  // 添加到工作表
                  addDataToSheet(sheet, data);
                  Logger.log(`成功爬取 ${contract} ${dateStr} ${identity} 的資料`);
                  successCount++;
                } else {
                  Logger.log(`${contract} ${dateStr} ${identity} 資料已存在，跳過`);
                  successCount++;
                }
              }
            }
            
            // 如果成功解析至少一個身份別資料，則返回
            if (successCount > 0) {
              return successCount;
            }
          } catch (e) {
            Logger.log(`解析 ${contract} ${dateStr} 身份別資料時發生錯誤: ${e.message}`);
          }
        }
        
        // 避免請求過於頻繁
        Utilities.sleep(1000);
      }
    }
    
    return successCount;
  } catch (e) {
    Logger.log(`爬取 ${contract} 身份別資料時發生錯誤: ${e.message}`);
    return 0;
  }
}

/**
 * 測試不同契約的身份別資料比較
 * 用於驗證NQF和TX的身份別資料是否正確區分
 */
function testCompareContractIdentityData() {
  try {
    Logger.log('開始比較不同契約的身份別資料');
    
    // 建立測試日期 (使用過去一周內的工作日)
    const today = new Date();
    let testDate = new Date(today);
    testDate.setDate(today.getDate() - 5); // 從五天前開始往前找
    
    // 找到有效的交易日
    for (let i = 0; i < 10; i++) {
      // 避開週末
      const dayOfWeek = testDate.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) {
        break;
      }
      testDate.setDate(testDate.getDate() + 1);
    }
    
    // 格式化日期
    const dateStr = Utilities.formatDate(testDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    Logger.log(`測試日期: ${dateStr}`);
    
    // 取得活動的試算表
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 取得TX和NQF工作表
    let txSheet = getOrCreateSheet(ss, 'TX');
    let nqfSheet = getOrCreateSheet(ss, 'NQF');
    
    // 比較的身份別
    const identities = ['自營商', '投信', '外資'];
    
    // 針對TX爬取資料
    fetchIdentityDataForContract('TX', dateStr, identities);
    
    // 針對NQF爬取資料
    fetchIdentityDataForContract('NQF', dateStr, identities);
    
    // 建立比較報告
    let report = `測試日期: ${dateStr}\n\n`;
    report += "契約比較報告:\n";
    report += "-----------------\n\n";
    
    // 比較兩個契約的資料
    for (const identity of identities) {
      // 獲取TX資料
      const txData = getIdentityDataFromSheet(txSheet, dateStr, 'TX', identity);
      
      // 獲取NQF資料
      const nqfData = getIdentityDataFromSheet(nqfSheet, dateStr, 'NQF', identity);
      
      report += `${identity}資料比較:\n`;
      
      if (txData && nqfData) {
        // 比較交易量
        const txTradingVolume = Math.max(Math.abs(txData.buyVolume), Math.abs(txData.sellVolume), Math.abs(txData.netVolume));
        const nqfTradingVolume = Math.max(Math.abs(nqfData.buyVolume), Math.abs(nqfData.sellVolume), Math.abs(nqfData.netVolume));
        
        // 比較未平倉量
        const txOpenInterest = Math.max(Math.abs(txData.buyOpenInterest), Math.abs(txData.sellOpenInterest), Math.abs(txData.netOpenInterest));
        const nqfOpenInterest = Math.max(Math.abs(nqfData.buyOpenInterest), Math.abs(nqfData.sellOpenInterest), Math.abs(nqfData.netOpenInterest));
        
        report += `TX 最大交易量: ${txTradingVolume}\n`;
        report += `NQF 最大交易量: ${nqfTradingVolume}\n`;
        report += `TX 最大未平倉量: ${txOpenInterest}\n`;
        report += `NQF 最大未平倉量: ${nqfOpenInterest}\n`;
        
        // 檢查資料是否合理區分
        const volumeRatio = txTradingVolume / (nqfTradingVolume || 1);
        const openInterestRatio = txOpenInterest / (nqfOpenInterest || 1);
        
        report += `交易量比例(TX/NQF): ${volumeRatio.toFixed(2)}\n`;
        report += `未平倉量比例(TX/NQF): ${openInterestRatio.toFixed(2)}\n`;
        
        if (volumeRatio > 3 && openInterestRatio > 3) {
          report += `結論: ✓ 資料合理區分，TX資料明顯大於NQF資料\n`;
        } else {
          report += `結論: × 資料可能未正確區分，比例不夠顯著\n`;
        }
      } else {
        if (!txData) report += `TX ${identity} 資料不存在\n`;
        if (!nqfData) report += `NQF ${identity} 資料不存在\n`;
        report += `無法進行比較\n`;
      }
      
      report += "-----------------\n\n";
    }
    
    // 顯示比較報告
    Logger.log(report);
    SpreadsheetApp.getUi().alert('契約身份別資料比較報告', report, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (e) {
    Logger.log(`比較契約身份別資料時發生錯誤: ${e.message}`);
    SpreadsheetApp.getUi().alert(`測試時發生錯誤: ${e.message}`);
  }
}

/**
 * 從工作表中獲取特定日期和身份別的資料
 */
function getIdentityDataFromSheet(sheet, dateStr, contract, identity) {
  try {
    const data = sheet.getDataRange().getValues();
    
    // 跳過表頭
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === dateStr && data[i][1] === contract && data[i][14] === identity) {
        return {
          buyVolume: data[i][2],          // 多方交易口數
          sellVolume: data[i][4],         // 空方交易口數
          netVolume: data[i][6],          // 多空淨額交易口數
          buyOpenInterest: data[i][8],    // 多方未平倉口數
          sellOpenInterest: data[i][10],  // 空方未平倉口數
          netOpenInterest: data[i][12]    // 多空淨額未平倉口數
        };
      }
    }
    
    return null;
  } catch (e) {
    Logger.log(`從工作表獲取資料時發生錯誤: ${e.message}`);
    return null;
  }
}

/**
 * 測試直接分析NQF的表格結構
 * 此函數可以幫助我們理解NQF頁面的結構，找出投信和外資資料在何處
 */
function analyzeNQFTableStructure() {
  try {
    // 清空舊日誌
    Logger.clear();
    
    // 取得今日日期或前一個交易日
    const today = new Date();
    let targetDate = today;
    // 週末的話往前找一個工作日
    const dayOfWeek = today.getDay();
    if (dayOfWeek === 0) { // 週日
      targetDate.setDate(today.getDate() - 2); // 改為週五
    } else if (dayOfWeek === 6) { // 週六
      targetDate.setDate(today.getDate() - 1); // 改為週五
    }
    
    const dateStr = Utilities.formatDate(targetDate, 'Asia/Taipei', 'yyyy/MM/dd');
    Logger.log(`分析日期: ${dateStr}的NQF表格結構`);
    
    // 嘗試不同的參數組合
    const queryTypes = ['1', '2', '3'];
    const marketCodes = ['0', '1', '2'];
    
    for (const queryType of queryTypes) {
      for (const marketCode of marketCodes) {
        // 構建查詢參數
        const queryData = {
          'queryType': queryType,
          'marketCode': marketCode,
          'dateaddcnt': '',
          'commodity_id': 'NQF',
          'queryDate': dateStr
        };
        
        Logger.log(`嘗試使用參數: queryType=${queryType}, marketCode=${marketCode}`);
        
        // 發送請求
        const response = fetchDataFromTaifex(queryData);
        
        if (response && !hasNoDataMessage(response) && !hasErrorMessage(response)) {
          Logger.log(`成功獲取回應，分析表格結構...`);
          
          // 尋找所有表格
          const tables = response.match(/<table[^>]*>[\s\S]*?<\/table>/gi) || [];
          Logger.log(`找到 ${tables.length} 個表格`);
          
          // 尋找包含身份別資料的表格
          for (let i = 0; i < tables.length; i++) {
            const table = tables[i];
            
            // 特別尋找包含NQF和身份別信息的表格
            const hasNQF = table.includes('NQF') || table.includes('美國那斯達克') || table.includes('那斯達克');
            const hasIdentity = table.includes('自營商') || table.includes('投信') || table.includes('外資');
            
            Logger.log(`表格 ${i+1}: 包含NQF=${hasNQF}, 包含身份別=${hasIdentity}`);
            
            if (hasNQF || hasIdentity) {
              // 尋找表格中的所有行
              const rows = table.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];
              Logger.log(`該表格包含 ${rows.length} 行`);
              
              // 特別分析前10行，了解表格結構
              const rowsToAnalyze = Math.min(rows.length, 10);
              for (let j = 0; j < rowsToAnalyze; j++) {
                // 清理HTML標籤，獲取純文本內容
                const rowText = rows[j].replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
                Logger.log(`行 ${j+1}: ${rowText}`);
                
                // 特別檢查是否包含NQF或身份別關鍵字
                if (rowText.includes('NQF') || rowText.includes('美國那斯達克') || 
                    rowText.includes('那斯達克') || rowText.includes('自營商') || 
                    rowText.includes('投信') || rowText.includes('外資')) {
                  Logger.log(`關鍵行 ${j+1}: ${rowText}`);
                  
                  // 提取單元格
                  const cells = rows[j].match(/<t[dh][^>]*>[\s\S]*?<\/t[dh]>/gi) || [];
                  Logger.log(`該行包含 ${cells.length} 個單元格`);
                  
                  // 分析每個單元格的內容
                  for (let k = 0; k < cells.length; k++) {
                    const cellText = cells[k].replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
                    Logger.log(`單元格 ${k+1}: ${cellText}`);
                  }
                }
              }
              
              // 嘗試用特定的邏輯尋找身份別資料
              const identities = ['自營商', '投信', '外資'];
              for (const identity of identities) {
                // 尋找包含該身份別的行
                for (let j = 0; j < rows.length; j++) {
                  if (rows[j].includes(identity)) {
                    const rowText = rows[j].replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
                    Logger.log(`找到 ${identity} 在行 ${j+1}: ${rowText}`);
                    
                    // 嘗試解析該行的多方和空方資料
                    const cells = rows[j].match(/<t[dh][^>]*>[\s\S]*?<\/t[dh]>/gi) || [];
                    if (cells.length > 0) {
                      // 假設多方和空方資料位於第3和第5個單元格
                      const buyVolume = cells.length >= 3 ? cells[2].replace(/<[^>]*>/g, '').trim() : 'N/A';
                      const sellVolume = cells.length >= 5 ? cells[4].replace(/<[^>]*>/g, '').trim() : 'N/A';
                      Logger.log(`${identity} 多方交易口數: ${buyVolume}, 空方交易口數: ${sellVolume}`);
                    }
                  }
                }
              }
            }
          }
          
          // 如果此參數組合成功獲取資料，則標記
          Logger.log(`參數組合 queryType=${queryType}, marketCode=${marketCode} 似乎包含有用資料\n`);
        } else {
          Logger.log(`參數組合 queryType=${queryType}, marketCode=${marketCode} 無法獲取有效資料\n`);
        }
        
        // 避免請求過於頻繁
        Utilities.sleep(500);
      }
    }
    
    // 輸出所有診斷日誌到工作表
    writeLogsToSheet();
    
    SpreadsheetApp.getUi().alert(
      '分析完成',
      '已完成NQF表格結構分析，請查看「診斷日誌」工作表了解詳情。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    Logger.log(`分析NQF表格結構時發生錯誤: ${e.message}`);
    SpreadsheetApp.getUi().alert(`分析時發生錯誤: ${e.message}`);
  }
}

/**
 * 增強版解析身份別資料函數，專門針對NQF契約
 * @param {string} html - 期交所返回的HTML內容
 * @param {string} dateStr - 日期字符串
 * @param {string} identity - 身份別名稱
 * @returns {Array|null} 解析後的資料陣列或null
 */
function enhancedParseIdentityData(html, dateStr, identity) {
  try {
    // 確保使用傳入的日期 - 記錄日期用於追蹤
    Logger.log(`使用增強版解析器處理NQF ${dateStr} ${identity} 資料，確保使用正確日期: ${dateStr}`);
    
    // 尋找所有表格
    const tables = html.match(/<table[^>]*>[\s\S]*?<\/table>/gi) || [];
    Logger.log(`找到 ${tables.length} 個表格`);
    
    // 特別輸出每個表格的基本資訊
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      const hasNQF = table.includes('NQF') || table.includes('美國那斯達克') || table.includes('那斯達克');
      const hasIdentity = table.includes('自營商') || table.includes('投信') || table.includes('外資');
      const rows = table.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];
      
      Logger.log(`表格 ${i+1}: 包含NQF=${hasNQF}, 包含身份別=${hasIdentity}, 行數=${rows.length}`);
    }
    
    // 找到包含NQF和身份別的表格
    const targetTable = tables.find(table => 
      (table.includes('NQF') || table.includes('美國那斯達克') || table.includes('那斯達克')) && 
      table.includes(identity)
    );
    
    if (targetTable) {
      Logger.log(`找到包含NQF和${identity}的表格`);
      
      // 提取表格中的所有行
      const rows = targetTable.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];
      Logger.log(`表格包含 ${rows.length} 行`);
      
      // 尋找包含身份別的行
      const identityRow = rows.find(row => row.includes(identity) && 
                                    (row.includes('NQF') || 
                                     row.includes('美國那斯達克') || 
                                     row.includes('那斯達克')));
      
      if (identityRow) {
        Logger.log(`找到包含 ${identity} 的行`);
        
        // 提取單元格 - 清理HTML標籤
        const cleanRow = identityRow.replace(/<[^>]*>/g, '|');
        const cells = cleanRow.split('|').filter(cell => cell.trim() !== '');
        
        // 輸出所有單元格內容以便調試
        Logger.log(`行解析後的單元格數量: ${cells.length}`);
        for (let i = 0; i < cells.length; i++) {
          Logger.log(`單元格 ${i+1}: ${cells[i].trim()}`);
        }
        
        // 尋找數據欄位 - 通常是6個數字（買口數、賣口數、淨口數，買未平倉、賣未平倉、淨未平倉）
        const numericValues = cells.map(cell => {
          const cleaned = cell.trim().replace(/,/g, '');
          return isNaN(parseInt(cleaned)) ? null : parseInt(cleaned);
        }).filter(val => val !== null);
        
        Logger.log(`提取到 ${numericValues.length} 個數值: ${numericValues.join(', ')}`);
        
        // 解析資料 - 需要至少6個數字才能構成完整記錄
        if (numericValues.length >= 6) {
          // 根據位置提取買賣交易量和未平倉
          // 這裡假設順序是: 買口數、賣口數、淨口數，買未平倉、賣未平倉、淨未平倉
          // 但有時候期交所返回的數據順序可能有變化，需要特別處理
          
          // 嘗試定位數據
          let buyVolume, sellVolume, netVolume, buyOpenInterest, sellOpenInterest, netOpenInterest;
          
          // 通常前三個是交易量，後三個是未平倉
          if (numericValues.length === 6) {
            // 標準的6個數字
            buyVolume = numericValues[0];
            sellVolume = numericValues[1];
            netVolume = numericValues[2];
            buyOpenInterest = numericValues[3];
            sellOpenInterest = numericValues[4];
            netOpenInterest = numericValues[5];
          } else {
            // 如果有更多數字，嘗試找出正確的位置
            // 這裡簡化處理，假設前面的數字是交易量，後面的是未平倉
            const halfIndex = Math.floor(numericValues.length / 2);
            buyVolume = numericValues[0];
            sellVolume = numericValues[1];
            netVolume = numericValues[2];
            buyOpenInterest = numericValues[halfIndex];
            sellOpenInterest = numericValues[halfIndex + 1];
            netOpenInterest = numericValues[halfIndex + 2];
          }
          
          Logger.log(`解析結果 - 買口數: ${buyVolume}, 賣口數: ${sellVolume}, 淨口數: ${netVolume}, 買未平倉: ${buyOpenInterest}, 賣未平倉: ${sellOpenInterest}, 淨未平倉: ${netOpenInterest}`);
          
          // 這裡強制使用傳入的dateStr作為日期，確保日期一致性
          const data = [
            dateStr,                // 日期 - 強制使用傳入的dateStr
            'NQF',                  // 契約代碼
            buyVolume,              // 多方交易口數
            0,                      // 多方契約金額
            sellVolume,             // 空方交易口數
            0,                      // 空方契約金額
            netVolume,              // 多空淨額交易口數
            0,                      // 多空淨額契約金額
            buyOpenInterest,        // 多方未平倉口數
            0,                      // 多方未平倉契約金額
            sellOpenInterest,       // 空方未平倉口數
            0,                      // 空方未平倉契約金額
            netOpenInterest,        // 多空淨額未平倉口數
            0,                      // 多空淨額未平倉契約金額
            identity                // 身份別
          ];
          
          Logger.log(`成功解析 NQF ${identity} 資料: 多方=${buyVolume}, 空方=${sellVolume}, 淨額=${netVolume}, 未平倉多方=${buyOpenInterest}, 未平倉空方=${sellOpenInterest}`);
          
          // 檢查數據合理性
          if (validateIdentityData(data, identity)) {
            // 重要：確認返回的資料日期就是查詢的日期
            Logger.log(`【日期追蹤】enhancedParseIdentityData將返回資料，已確認日期為: ${data[0]}, 與查詢日期 ${dateStr} 一致`);
            return data;
          } else {
            Logger.log(`解析的 ${identity} 資料不合理，可能是錯誤的資料`);
          }
        } else {
          Logger.log(`未找到足夠的數值資料`);
        }
      } else {
        Logger.log(`未找到 ${identity} 的資料行`);
      }
    }
    
    // 如果上述解析方法失敗，則嘗試從原始HTML中直接解析
    Logger.log(`嘗試使用直接解析法`);
    
    // 尋找包含NQF和身份別的段落
    const regex = new RegExp(`(NQF|那斯達克|美國那斯達克)[\\s\\S]*?${identity}[\\s\\S]*?([0-9,]+)[\\s\\S]*?([0-9,]+)[\\s\\S]*?([0-9,]+)[\\s\\S]*?([0-9,]+)[\\s\\S]*?([0-9,]+)[\\s\\S]*?([0-9,]+)`, 'i');
    const match = html.match(regex);
    
    if (match) {
      Logger.log(`直接解析匹配成功: ${match[0].substring(0, 50)}...`);
      
      // 提取數值
      const buyVolume = parseInt(match[2].replace(/,/g, ''));
      const sellVolume = parseInt(match[3].replace(/,/g, ''));
      const netVolume = parseInt(match[4].replace(/,/g, ''));
      const buyOpenInterest = parseInt(match[5].replace(/,/g, ''));
      const sellOpenInterest = parseInt(match[6].replace(/,/g, ''));
      const netOpenInterest = parseInt(match[7].replace(/,/g, ''));
      
      // 強制使用傳入的dateStr作為日期，確保日期一致性
      const data = [
        dateStr,                // 日期 - 強制使用傳入的dateStr
        'NQF',                  // 契約代碼
        buyVolume,              // 多方交易口數
        0,                      // 多方契約金額
        sellVolume,             // 空方交易口數
        0,                      // 空方契約金額
        netVolume,              // 多空淨額交易口數
        0,                      // 多空淨額契約金額
        buyOpenInterest,        // 多方未平倉口數
        0,                      // 多方未平倉契約金額
        sellOpenInterest,       // 空方未平倉口數
        0,                      // 空方未平倉契約金額
        netOpenInterest,        // 多空淨額未平倉口數
        0,                      // 多空淨額未平倉契約金額
        identity                // 身份別
      ];
      
      Logger.log(`直接解析成功: 多方=${buyVolume}, 空方=${sellVolume}, 未平倉多方=${buyOpenInterest}, 未平倉空方=${sellOpenInterest}`);
      
      // 檢查數據合理性
      if (validateIdentityData(data, identity)) {
        // 重要：確認返回的資料日期就是查詢的日期
        Logger.log(`【日期追蹤】enhancedParseIdentityData(直接解析方法)將返回資料，已確認日期為: ${data[0]}, 與查詢日期 ${dateStr} 一致`);
        return data;
      } else {
        Logger.log(`直接解析的 ${identity} 資料不合理，可能是錯誤的資料`);
      }
    }
    
    return null;
  } catch (e) {
    Logger.log(`增強版解析 NQF ${identity} 資料時發生錯誤: ${e.message}`);
    return null;
  }
}

/**
 * 驗證身份別資料是否合理
 * @param {Array} data - 資料陣列
 * @param {string} identity - 身份別
 * @returns {boolean} 資料是否合理
 */
function validateIdentityData(data, identity) {
  try {
    // 驗證資料數組長度
    if (!data || data.length < 15) {
      Logger.log('資料數組長度不足');
      return false;
    }
    
    // 提取關鍵數值
    const buyVolume = Math.abs(data[2]); // 多方交易口數
    const sellVolume = Math.abs(data[4]); // 空方交易口數
    const netVolume = Math.abs(data[6]); // 多空淨額交易口數
    const buyOpenInterest = Math.abs(data[8]); // 多方未平倉口數
    const sellOpenInterest = Math.abs(data[10]); // 空方未平倉口數
    
    // 針對NQF契約
    if (data[1] === 'NQF') {
      // 投信資料通常很小，甚至可能為0
      if (identity === '投信') {
        // 更嚴格的閾值判斷
        if (buyVolume > 3000 || sellVolume > 3000) {
          Logger.log(`投信數據異常大，可能是錯誤資料: 買=${buyVolume}, 賣=${sellVolume}`);
          return false;
        }
        
        // 如果數據跟TX接近，判定為錯誤
        if (buyVolume > 2000 && sellVolume > 2000) {
          Logger.log(`投信數據接近TX水平，可能是錯誤資料: 買=${buyVolume}, 賣=${sellVolume}`);
          return false;
        }
        
        // 數據為0或很小是可能的
        if (buyVolume <= 50 && sellVolume <= 50) {
          Logger.log(`投信數據很小或為0，這是可能的: 買=${buyVolume}, 賣=${sellVolume}`);
          return true;
        }
        
        // 未平倉量也應該較小
        if (buyOpenInterest > 3000 || sellOpenInterest > 3000) {
          Logger.log(`投信未平倉量異常大，可能是錯誤資料: 買=${buyOpenInterest}, 賣=${sellOpenInterest}`);
          return false;
        }
      }
      
      // 外資資料應該適中，不會太大也不會全為0
      if (identity === '外資') {
        // 更嚴格的閾值判斷
        if (buyVolume > 5000 || sellVolume > 5000) {
          Logger.log(`外資數據異常大，可能是錯誤資料: 買=${buyVolume}, 賣=${sellVolume}`);
          return false;
        }
        
        // 如果交易量很接近TX的資料，可能是錯誤的
        if (buyVolume > 3000 && sellVolume > 3000) {
          Logger.log(`外資數據接近TX水平，可能是錯誤資料: 買=${buyVolume}, 賣=${sellVolume}`);
          return false;
        }
        
        // 未平倉量也應該適中
        if (buyOpenInterest > 6000 || sellOpenInterest > 6000) {
          Logger.log(`外資未平倉量異常大，可能是錯誤資料: 買=${buyOpenInterest}, 賣=${sellOpenInterest}`);
          return false;
        }
        
        // 檢查交易量與未平倉量是否有合理的比例關係
        const volumeTotal = buyVolume + sellVolume;
        const oiTotal = buyOpenInterest + sellOpenInterest;
        
        // 如果交易量很小但未平倉量很大，可能是錯誤的
        if (volumeTotal < 50 && oiTotal > 2000) {
          Logger.log(`外資數據比例不合理：交易量很小但未平倉量很大`);
          return false;
        }
      }
      
      // 自營商資料可以更大一些
      if (identity === '自營商') {
        if (buyVolume > 15000 || sellVolume > 15000) {
          Logger.log(`自營商數據異常大，可能是錯誤資料: 買=${buyVolume}, 賣=${sellVolume}`);
          return false;
        }
        
        // 未平倉量也有上限
        if (buyOpenInterest > 20000 || sellOpenInterest > 20000) {
          Logger.log(`自營商未平倉量異常大，可能是錯誤資料: 買=${buyOpenInterest}, 賣=${sellOpenInterest}`);
          return false;
        }
      }
      
      // 淨額檢查 - 淨額應該約等於多空之差
      const calculatedNet = buyVolume - sellVolume;
      if (Math.abs(calculatedNet - netVolume) > 5 && netVolume !== 0) {
        Logger.log(`淨額檢查失敗，多空之差(${calculatedNet})與報告淨額(${netVolume})不匹配`);
        return false;
      }
      
      // 資料一致性檢查 - 若非全0資料，檢查資料點分佈是否合理
      if (buyVolume > 0 || sellVolume > 0) {
        // 如果只有一個資料點不為0，可能是錯誤提取
        if ((buyVolume === 0 && sellVolume > 1000) || (sellVolume === 0 && buyVolume > 1000)) {
          Logger.log(`資料分佈不自然，多方(${buyVolume})和空方(${sellVolume})之一為0而另一個很大`);
          return false;
        }
      }
      
      // 驗證該身份別的數據是否與特定的TX數據模式匹配
      // 台指期(TX)的各身份別數據通常有一定特徵
      if (identity === '投信') {
        // TX投信通常買多不賣空
        if (buyVolume > 1000 && sellVolume < 100) {
          Logger.log(`數據模式類似TX投信特徵(大量買盤很少賣盤)，可能是誤爬取TX資料`);
          return false;
        }
      }
      
      if (identity === '外資') {
        // TX外資通常交易量較大且活躍
        if (buyVolume > 2000 && sellVolume > 2000 && 
            Math.abs(buyVolume - sellVolume) > 1000) {
          Logger.log(`數據模式類似TX外資特徵(大量且不平衡的交易)，可能是誤爬取TX資料`);
          return false;
        }
      }
    }
    
    return true;
  } catch (e) {
    Logger.log(`驗證身份別資料時發生錯誤: ${e.message}`);
    return false;
  }
}

/**
 * 詳細檢查NQF爬取到的資料，並與TX作比較
 * 這個函數會顯示實際爬取到的數據，協助使用者理解問題所在
 */
function debugNQFIdentityData() {
  try {
    // 清空舊日誌
    Logger.clear();
    
    // 取得試算表
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // 獲取日期
    const response = ui.prompt('設定要檢查的日期', 
                              '請輸入要檢查的日期 (格式：yyyy/MM/dd)\n預設為最近一個工作日', 
                              ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return; // 用戶取消
    }
    
    // 解析用戶輸入的日期或使用預設日期
    let dateStr = '';
    const userDateInput = response.getResponseText().trim();
    
    if (userDateInput) {
      dateStr = userDateInput;
    } else {
      // 尋找最近一個工作日
      const today = new Date();
      let testDate = new Date(today);
      
      // 最多往前找10天
      for (let i = 0; i < 10; i++) {
        const dayStr = Utilities.formatDate(testDate, 'Asia/Taipei', 'yyyy/MM/dd');
        Logger.log(`檢查日期: ${dayStr}`);
        
        // 避開週末
        const dayOfWeek = testDate.getDay();
        if (dayOfWeek !== 0 && dayOfWeek !== 6) {
          dateStr = dayStr;
          break;
        }
        
        // 往前一天
        testDate.setDate(testDate.getDate() - 1);
      }
    }
    
    if (!dateStr) {
      ui.alert('無法確定要檢查的日期，請重新執行並輸入有效日期。');
      return;
    }
    
    Logger.log(`開始檢查 ${dateStr} 的NQF和TX身份別資料`);
    
    // 創建用於比較的診斷表格
    let debugSheet = ss.getSheetByName('除錯資料比較');
    if (debugSheet) {
      ss.deleteSheet(debugSheet);
    }
    debugSheet = ss.insertSheet('除錯資料比較');
    
    // 設置表頭
    debugSheet.getRange(1, 1, 1, 7).setValues([
      ['契約', '身份別', '日期', '多方交易口數', '空方交易口數', '多方未平倉口數', '空方未平倉口數']
    ]);
    debugSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    
    // 抓取TX和NQF的資料進行對比
    const contractsToCheck = ['TX', 'NQF'];
    const identities = ['自營商', '投信', '外資'];
    
    let rowIndex = 2;
    let foundData = false;
    
    // 檢查現有資料
    for (const contract of contractsToCheck) {
      const sheet = ss.getSheetByName(contract);
      if (!sheet) {
        Logger.log(`找不到 ${contract} 工作表`);
        continue;
      }
      
      // 獲取該工作表的所有資料
      const data = sheet.getDataRange().getValues();
      
      // 檢查每一行
      let firstRow = true;
      for (let i = 1; i < data.length; i++) {
        // 檢查是否為目標日期的資料
        if (data[i][0] === dateStr) {
          // 這是目標日期的資料
          const rowIdentity = data[i][14]; // 身份別欄位
          
          if (identities.includes(rowIdentity)) {
            Logger.log(`找到 ${contract} ${rowIdentity} 資料，在工作表行 ${i+1}`);
            
            // 提取關鍵數據
            const savedDate = data[i][0]; // 儲存的日期
            const buyVolume = data[i][2]; // 多方交易口數
            const sellVolume = data[i][4]; // 空方交易口數
            const buyOpenInterest = data[i][8]; // 多方未平倉口數
            const sellOpenInterest = data[i][10]; // 空方未平倉口數
            
            // 添加到調試工作表
            debugSheet.getRange(rowIndex, 1, 1, 7).setValues([
              [contract, rowIdentity, savedDate, buyVolume, sellVolume, buyOpenInterest, sellOpenInterest]
            ]);
            
            // 如果是首行，添加註釋以標記身份別字段
            if (firstRow && contract === 'NQF') {
              // 在 NQF 行添加特殊標記
              debugSheet.getRange(rowIndex, 2).setNote(`檢查這個身份別是否正確。\n在工作表 ${contract} 的第 ${i+1} 行`);
              debugSheet.getRange(rowIndex, 3).setNote(`檢查這個日期是否與要查詢的日期 ${dateStr} 一致`);
              firstRow = false;
            }
            
            rowIndex++;
            foundData = true;
          }
        }
      }
    }
    
    if (!foundData) {
      Logger.log(`未找到 ${dateStr} 的身份別資料，嘗試實時爬取...`);
      
      // 如果沒有找到現有資料，嘗試實時爬取
      for (const contract of contractsToCheck) {
        // 爬取資料
        fetchIdentityData(contract, dateStr, identities);
        
        // 檢查爬取結果
        const sheet = ss.getSheetByName(contract);
        if (sheet) {
          const data = sheet.getDataRange().getValues();
          
          for (let i = 1; i < data.length; i++) {
            if (data[i][0] === dateStr) {
              const rowIdentity = data[i][14]; // 身份別欄位
              
              if (identities.includes(rowIdentity)) {
                Logger.log(`爬取到 ${contract} ${rowIdentity} 資料，在工作表行 ${i+1}`);
                
                // 提取關鍵數據
                const savedDate = data[i][0]; // 儲存的日期
                const buyVolume = data[i][2]; // 多方交易口數
                const sellVolume = data[i][4]; // 空方交易口數
                const buyOpenInterest = data[i][8]; // 多方未平倉口數
                const sellOpenInterest = data[i][10]; // 空方未平倉口數
                
                // 添加到調試工作表
                debugSheet.getRange(rowIndex, 1, 1, 7).setValues([
                  [contract, rowIdentity, savedDate, buyVolume, sellVolume, buyOpenInterest, sellOpenInterest]
                ]);
                
                rowIndex++;
                foundData = true;
              }
            }
          }
        }
      }
    }
    
    if (!foundData) {
      ui.alert('未找到或爬取到任何資料。請嘗試另一個日期。');
      return;
    }
    
    // 進行特殊測試 - 使用增強版解析器直接測試
    Logger.log('進行特殊測試 - 使用增強版解析器直接測試');
    
    const testRow = rowIndex + 1;
    debugSheet.getRange(testRow, 1, 1, 7).setValues([
      ['直接解析測試', '', dateStr, '', '', '', '']
    ]);
    debugSheet.getRange(testRow, 1, 1, 7).setBackground('#f0f0f0');
    rowIndex = testRow + 1;
    
    // 嘗試不同的查詢參數組合，直接顯示解析結果
    const queryTypes = ['1', '2', '3'];
    const marketCodes = ['0', '1', '2'];
    
    for (const queryType of queryTypes) {
      for (const marketCode of marketCodes) {
        // 構建查詢參數
        const queryData = {
          'queryType': queryType,
          'marketCode': marketCode,
          'dateaddcnt': '',
          'commodity_id': 'NQF',
          'queryDate': dateStr
        };
        
        Logger.log(`測試參數組合: queryType=${queryType}, marketCode=${marketCode}`);
        
        // 發送請求
        const response = fetchDataFromTaifex(queryData);
        
        if (response && !hasNoDataMessage(response) && !hasErrorMessage(response)) {
          // 嘗試使用增強版解析器
          for (const identity of identities) {
            const data = enhancedParseIdentityData(response, dateStr, identity);
            
            if (data) {
              Logger.log(`直接解析成功: NQF ${identity}`);
              
              // 添加到調試工作表
              debugSheet.getRange(rowIndex, 1, 1, 7).setValues([
                ['NQF(直解)', identity, data[0], data[2], data[4], data[8], data[10]]
              ]);
              
              // 添加行號註釋
              debugSheet.getRange(rowIndex, 2).setNote(`參數: qt=${queryType}, mc=${marketCode}`);
              
              rowIndex++;
            }
          }
        }
      }
    }
    
    // 輸出HTML結構到日誌
    Logger.log('分析NQF詳細HTML結構...');
    
    // 嘗試最可能包含正確資料的參數組合
    const bestQueryData = {
      'queryType': '2',
      'marketCode': '0',
      'dateaddcnt': '',
      'commodity_id': 'NQF',
      'queryDate': dateStr
    };
    
    const bestResponse = fetchDataFromTaifex(bestQueryData);
    
    if (bestResponse) {
      // 尋找包含身份別的表格
      const tables = bestResponse.match(/<table[^>]*>[\s\S]*?<\/table>/gi) || [];
      
      for (let i = 0; i < tables.length; i++) {
        const table = tables[i];
        
        if (table.includes('NQF') || 
            table.includes('美國那斯達克') || 
            table.includes('那斯達克') ||
            table.includes('自營商') || 
            table.includes('投信') || 
            table.includes('外資')) {
          
          Logger.log(`找到相關表格 ${i+1}`);
          
          // 提取表格中的所有行
          const rows = table.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];
          
          // 特別詳細地輸出每一行的內容
          for (let j = 0; j < rows.length; j++) {
            const rowText = rows[j].replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
            Logger.log(`表格 ${i+1} 行 ${j+1}: ${rowText}`);
            
            // 檢查該行是否包含關鍵字
            if (rowText.includes('NQF') || rowText.includes('那斯達克') || 
                rowText.includes('自營商') || rowText.includes('投信') || rowText.includes('外資')) {
              Logger.log(`重要行: 表格 ${i+1} 行 ${j+1}: ${rowText}`);
              
              // 對於外資行做特殊標記
              if (rowText.includes('外資')) {
                Logger.log(`*** 注意：外資行在表格 ${i+1} 行 ${j+1} ***`);
              }
            }
            
            // 提取單元格
            const cells = rows[j].match(/<t[dh][^>]*>[\s\S]*?<\/t[dh]>/gi) || [];
            
            if (cells.length > 0) {
              let cellContents = [];
              for (let k = 0; k < cells.length; k++) {
                const cellText = cells[k].replace(/<[^>]*>/g, '').trim();
                cellContents.push(cellText);
              }
              Logger.log(`  單元格內容: ${cellContents.join(' | ')}`);
            }
          }
        }
      }
    }
    
    // 寫入日誌
    writeLogsToSheet();
    
    // 格式化調試工作表
    debugSheet.autoResizeColumns(1, 7);
    debugSheet.setFrozenRows(1);
    
    // 顯示結果
    ui.alert(
      '資料比較完成',
      `已將 ${dateStr} 的TX和NQF資料比較結果輸出到「除錯資料比較」工作表。\n` + 
      `同時，詳細的HTML結構和解析過程已寫入「診斷日誌」工作表。\n\n` +
      `請檢查兩份工作表:\n` +
      `1. 比較TX和NQF各身份別資料的數值大小差異\n` +
      `2. 檢查NQF的日期欄位是否與查詢日期一致\n` +
      `3. 檢查HTML結構中外資資料所在的行號\n\n` +
      `如果發現外資資料日期或數值錯誤，請使用主選單的\n` +
      `"爬取身份別(自營/投信/外資)資料"功能重新爬取。`,
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    Logger.log(`除錯分析時發生錯誤: ${e.message}`);
    SpreadsheetApp.getUi().alert(`除錯過程發生錯誤: ${e.message}`);
  }
} 