// 期貨資料圖表工具

/**
 * 為指定的契約創建交易量和未平倉量的圖表
 * @param {string} contract - 契約代碼 (TX, TE, MTX, ZMX, NQF)
 * @param {number} days - 需要顯示的天數（從今天向前推算）
 */
function createContractChart(contract, days) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    Logger.log(`開始為 ${contract} 創建圖表，資料範圍：最近 ${days} 天`);
    Logger.log(`試算表名稱：${ss.getName()}`);
    
    // 靈活查找資料工作表
    let sheet = null;
    const possibleSheetNames = [
      ALL_CONTRACTS_SHEET_NAME,  // 主要工作表名稱
      '所有期貨資料',             // 備用名稱1
      'All_Contracts',          // 備用名稱2
      'FuturesData'             // 備用名稱3
    ];
    
    // 先記錄所有工作表名稱以供診斷
    const allSheets = ss.getSheets();
    const allSheetNames = allSheets.map(s => s.getName());
    Logger.log(`試算表中的所有工作表：${allSheetNames.join(', ')}`);
    
    // 嘗試找到資料工作表
    for (const sheetName of possibleSheetNames) {
      sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        Logger.log(`找到資料工作表：${sheetName}`);
        break;
      }
    }
    
    // 如果還是找不到，嘗試找包含期貨資料的工作表
    if (!sheet) {
      Logger.log('未找到標準名稱的工作表，嘗試搜尋包含期貨相關關鍵字的工作表');
      
      for (const s of allSheets) {
        const name = s.getName().toLowerCase();
        if (name.includes('期貨') || name.includes('futures') || name.includes('contracts') || name.includes('data')) {
          sheet = s;
          Logger.log(`找到可能的資料工作表：${s.getName()}`);
          break;
        }
      }
    }
    
    // 如果仍然找不到，使用第一個有資料的工作表
    if (!sheet && allSheets.length > 0) {
      // 檢查每個工作表是否有足夠的資料
      for (const s of allSheets) {
        if (s.getLastRow() > 5 && s.getLastColumn() > 5) { // 至少有一些資料
          sheet = s;
          Logger.log(`使用第一個有資料的工作表：${s.getName()}`);
          break;
        }
      }
    }
    
    if (!sheet) {
      const message = `找不到期貨資料工作表！\n\n` +
        `請確認以下事項：\n` +
        `• 是否已執行資料爬取\n` +
        `• 工作表是否名為「${ALL_CONTRACTS_SHEET_NAME}」\n` +
        `• 或包含「期貨」關鍵字\n\n` +
        `目前的工作表：${allSheetNames.join(', ')}`;
      SpreadsheetApp.getUi().alert('找不到資料工作表', message, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 詳細檢查工作表資料
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    Logger.log(`工作表「${sheet.getName()}」資訊：`);
    Logger.log(`- 最後一行：${lastRow}`);
    Logger.log(`- 最後一列：${lastCol}`);
    
    // 如果只有表頭，則沒有數據
    if (lastRow <= 1) {
      const message = `工作表「${sheet.getName()}」中沒有資料！\n\n` +
        `最後一行：${lastRow}（應該至少為2才有資料）\n\n` +
        `請先執行資料爬取。`;
      SpreadsheetApp.getUi().alert('工作表中沒有資料', message, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 檢查表頭格式
    const headers = sheet.getRange(1, 1, 1, Math.min(lastCol, 15)).getValues()[0];
    Logger.log(`表頭：${headers.join(' | ')}`);
    
    // 獲取所有資料
    const data = sheet.getDataRange().getValues();
    Logger.log(`總資料筆數（包含表頭）：${data.length}`);
    
    // 檢查資料格式，找出正確的欄位索引
    const sampleRows = data.slice(1, Math.min(6, data.length)); // 取前5筆資料作為範例
    Logger.log(`資料範例（前${Math.min(5, sampleRows.length)}筆）：`);
    sampleRows.forEach((row, index) => {
      Logger.log(`第${index + 1}筆：${row.slice(0, Math.min(8, row.length)).join(' | ')}`);
    });
    
    // 智能識別欄位位置
    let dateColIndex = 0;      // 日期欄位索引
    let contractColIndex = 1;  // 契約代碼欄位索引
    let identityColIndex = 2;  // 身份別欄位索引
    let volumeColIndex = 3;    // 成交量欄位索引
    let openInterestColIndex = 4; // 未平倉欄位索引
    
    // 嘗試從表頭識別欄位
    for (let i = 0; i < headers.length; i++) {
      const header = (headers[i] || '').toString().toLowerCase();
      
      if (header.includes('日期') || header.includes('date')) {
        dateColIndex = i;
      } else if (header.includes('契約') || header.includes('代碼') || header.includes('contract')) {
        contractColIndex = i;
      } else if (header.includes('身份') || header.includes('身分') || header.includes('identity')) {
        identityColIndex = i;
      } else if (header.includes('交易口數') || header.includes('成交量') || header.includes('volume')) {
        volumeColIndex = i;
      } else if (header.includes('未平倉') || header.includes('open') && header.includes('interest')) {
        openInterestColIndex = i;
      }
    }
    
    Logger.log(`欄位索引識別結果：`);
    Logger.log(`- 日期欄位：${dateColIndex} (${headers[dateColIndex]})`);
    Logger.log(`- 契約欄位：${contractColIndex} (${headers[contractColIndex]})`);
    Logger.log(`- 身份別欄位：${identityColIndex} (${headers[identityColIndex]})`);
    Logger.log(`- 成交量欄位：${volumeColIndex} (${headers[volumeColIndex]})`);
    Logger.log(`- 未平倉欄位：${openInterestColIndex} (${headers[openInterestColIndex]})`);
    
    // 過濾出指定契約的所有資料
    const allContractData = data.slice(1).filter(row => {
      const rowContract = row[contractColIndex];
      const isTargetContract = rowContract === contract;
      return isTargetContract;
    });
    
    Logger.log(`${contract} 總資料筆數：${allContractData.length}`);
    
    if (allContractData.length === 0) {
      // 顯示實際存在的契約代碼以供除錯
      const existingContracts = [...new Set(data.slice(1).map(row => row[contractColIndex]).filter(c => c))];
      
      const message = `合約 ${contract} 沒有任何資料！\n\n` +
        `工作表中存在的契約代碼：\n${existingContracts.join(', ')}\n\n` +
        `請確認：\n` +
        `• 契約代碼是否正確（區分大小寫）\n` +
        `• 是否已爬取該契約的資料\n\n` +
        `建議：先執行「快速爬取(含所有身分別)」或「爬取今日基本資料」`;
      SpreadsheetApp.getUi().alert('無法創建圖表', message, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 分析身份別分布
    const identityBreakdown = {};
    allContractData.forEach(row => {
      const identity = row[identityColIndex] || '空白';
      identityBreakdown[identity] = (identityBreakdown[identity] || 0) + 1;
    });
    
    Logger.log(`${contract} 身份別分布：`, identityBreakdown);
    
    // 過濾出總計資料（用於圖表生成）
    const contractData = allContractData.filter(row => {
      const identity = row[identityColIndex] || '';
      const isTotalData = identity === '' || identity === '總計' || !['自營商', '投信', '外資'].includes(identity);
      return isTotalData;
    });
    
    Logger.log(`${contract} 總計資料筆數：${contractData.length}`);
    
    if (contractData.length === 0) {
      let identityList = Object.keys(identityBreakdown).map(key => `"${key}": ${identityBreakdown[key]} 筆`).join('\n');
      
      const message = `合約 ${contract} 沒有可用於圖表的總計資料！\n\n` +
        `目前資料的身份別分布：\n${identityList}\n\n` +
        `圖表生成需要：\n` +
        `• 身份別欄位為空白，或\n` +
        `• 身份別欄位為"總計"\n\n` +
        `建議解決方案：\n` +
        `• 使用「強制重新爬取基本資料」功能\n` +
        `• 選擇合約代碼：${contract}`;
      
      SpreadsheetApp.getUi().alert('無法創建圖表', message, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 按日期排序
    contractData.sort((a, b) => {
      const dateA = a[dateColIndex] instanceof Date ? a[dateColIndex] : new Date(a[dateColIndex]);
      const dateB = b[dateColIndex] instanceof Date ? b[dateColIndex] : new Date(b[dateColIndex]);
      return dateA - dateB;
    });
    
    // 取最近指定天數的資料
    const recentData = contractData.slice(-days);
    
    if (recentData.length === 0) {
      SpreadsheetApp.getUi().alert(`合約 ${contract} 在最近 ${days} 天沒有資料！`);
      return;
    }
    
    Logger.log(`${contract} 最近 ${days} 天的資料筆數：${recentData.length}`);
    
    // 檢查是否有有效的數值資料
    const validData = recentData.filter(row => {
      const volume = parseFloat(row[volumeColIndex]) || 0; // 成交量
      const openInterest = parseFloat(row[openInterestColIndex]) || 0; // 未平倉
      return volume > 0 || openInterest > 0;
    });
    
    if (validData.length === 0) {
      const message = `合約 ${contract} 的資料中沒有有效的交易量或未平倉量數據！\n\n` +
        `請檢查：\n` +
        `• 成交量欄位（第${volumeColIndex + 1}欄）：${headers[volumeColIndex]}\n` +
        `• 未平倉欄位（第${openInterestColIndex + 1}欄）：${headers[openInterestColIndex]}\n\n` +
        `如果欄位不正確，請檢查資料格式。`;
      SpreadsheetApp.getUi().alert('無法創建圖表', message, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    Logger.log(`${contract} 有效資料筆數：${validData.length}`);
    
    // 創建圖表工作表名稱
    const chartSheetName = `圖表_${contract}`;
    
    // 刪除現有的圖表工作表（如果存在）
    const existingSheet = ss.getSheetByName(chartSheetName);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }
    
    // 創建新的圖表工作表
    const chartSheet = ss.insertSheet(chartSheetName);
    
    // 設置表頭
    const chartHeaders = ['日期', '成交量', '未平倉量'];
    chartSheet.getRange(1, 1, 1, chartHeaders.length).setValues([chartHeaders]);
    
    // 格式化表頭
    const headerRange = chartSheet.getRange(1, 1, 1, chartHeaders.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    
    // 填入數據
    const chartData = validData.map(row => {
      const date = row[dateColIndex] instanceof Date ? row[dateColIndex] : new Date(row[dateColIndex]);
      const volume = parseFloat(row[volumeColIndex]) || 0;
      const openInterest = parseFloat(row[openInterestColIndex]) || 0;
      return [date, volume, openInterest];
    });
    
    // 將數據寫入工作表
    if (chartData.length > 0) {
      chartSheet.getRange(2, 1, chartData.length, 3).setValues(chartData);
      
      // 格式化日期欄
      chartSheet.getRange(2, 1, chartData.length, 1).setNumberFormat('mm/dd');
      
      // 格式化數值欄
      chartSheet.getRange(2, 2, chartData.length, 2).setNumberFormat('#,##0');
    }
    
    // 創建圖表
    const dataRange = chartSheet.getRange(1, 1, chartData.length + 1, 3);
    
    const chart = chartSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRange)
      .setPosition(1, 5, 0, 0) // 從第E欄開始放置圖表
      .setOption('title', `${CONTRACT_NAMES[contract]} (${contract}) - 交易量與未平倉量趨勢`)
      .setOption('hAxis.title', '日期')
      .setOption('vAxes.0.title', '成交量')
      .setOption('vAxes.1.title', '未平倉量')
      .setOption('series.0.targetAxisIndex', 0) // 成交量使用左軸
      .setOption('series.1.targetAxisIndex', 1) // 未平倉量使用右軸
      .setOption('series.0.color', '#ff6b6b') // 成交量為紅色
      .setOption('series.1.color', '#4ecdc4') // 未平倉量為綠色
      .setOption('legend.position', 'top')
      .setOption('legend.alignment', 'center')
      .setOption('width', 800)
      .setOption('height', 400)
      .build();
    
    chartSheet.insertChart(chart);
    
    // 自動調整欄寬
    chartSheet.autoResizeColumns(1, 3);
    
    // 切換到圖表工作表
    ss.setActiveSheet(chartSheet);
    
    // 顯示成功訊息
    const successMessage = `成功為 ${CONTRACT_NAMES[contract]} (${contract}) 創建圖表！\n\n` +
      `資料來源：「${sheet.getName()}」工作表\n` +
      `資料範圍：最近 ${days} 天\n` +
      `有效資料點：${validData.length} 個\n` +
      `圖表已顯示在「${chartSheetName}」工作表中`;
    
    SpreadsheetApp.getUi().alert('圖表創建成功', successMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
    Logger.log(`成功為 ${contract} 創建圖表，共 ${validData.length} 個有效資料點`);
    
  } catch (e) {
    Logger.log(`創建 ${contract} 圖表時發生錯誤: ${e.message}`);
    Logger.log(`錯誤詳情: ${e.stack}`);
    SpreadsheetApp.getUi().alert(`創建圖表時發生錯誤: ${e.message}\n\n請查看執行日誌獲取詳細資訊。`);
  }
}

/**
 * 創建並發送圖表到通訊平台
 * @param {string} contract - 契約代碼
 * @param {number} days - 顯示天數
 * @param {boolean} sendToTelegram - 是否發送到Telegram
 * @param {boolean} sendToLine - 是否發送到LINE
 */
function createAndSendChart(contract, days = 30, sendToTelegram = true, sendToLine = false) {
  try {
    Logger.log(`開始為 ${contract} 生成並發送圖表，天數: ${days}`);
    
    // 生成圖表
    const chartInfo = createContractChart(contract, days);
    
    if (!chartInfo) {
      Logger.log(`${contract} 圖表生成失敗`);
      return false;
    }
    
    Logger.log(`${contract} 圖表生成成功，準備發送`);
    
    // 等待圖表完全渲染
    Utilities.sleep(3000);
    SpreadsheetApp.flush();
    
    // 獲取圖表
    const charts = chartInfo.sheet.getCharts();
    if (charts.length === 0) {
      Logger.log(`${contract} 圖表工作表中沒有找到圖表`);
      return false;
    }
    
    let success = false;
    
    // 發送每個圖表
    for (let i = 0; i < charts.length; i++) {
      const chart = charts[i];
      const chartType = i === 0 ? '交易量' : '未平倉量';
      
      Logger.log(`準備發送 ${contract} ${chartType}圖表`);
      
      try {
        // 獲取圖表圖片
        const chartBlob = chart.getAs('image/png');
        
        if (!chartBlob || chartBlob.getBytes().length < 1000) {
          Logger.log(`${contract} ${chartType}圖表圖片生成失敗或太小`);
          continue;
        }
        
        Logger.log(`${contract} ${chartType}圖表圖片生成成功，大小: ${chartBlob.getBytes().length} bytes`);
        
        // 準備說明文字
        const caption = generateChartCaption(contract, chartType, days, chartInfo.dataLength);
        
        // 發送到Telegram
        if (sendToTelegram) {
          const telegramSuccess = sendTelegramPhoto(chartBlob, caption);
          if (telegramSuccess) {
            Logger.log(`${contract} ${chartType}圖表成功發送到Telegram`);
            success = true;
          } else {
            Logger.log(`${contract} ${chartType}圖表發送到Telegram失敗`);
          }
        }
        
        // 發送到LINE（如果需要）
        if (sendToLine) {
          // LINE Notify只支援文字，這裡可以發送圖表連結
          const lineMessage = `📊 ${CONTRACT_NAMES[contract]} ${chartType}圖表已生成\n\n${caption.replace(/<[^>]*>/g, '')}`;
          const lineSuccess = sendLineNotifyMessage(lineMessage);
          if (lineSuccess) {
            Logger.log(`${contract} ${chartType}圖表說明成功發送到LINE`);
            success = true;
          }
        }
        
        // 圖表間延遲
        if (i < charts.length - 1) {
          Utilities.sleep(2000);
        }
        
      } catch (e) {
        Logger.log(`發送 ${contract} ${chartType}圖表時發生錯誤: ${e.message}`);
      }
    }
    
    return success;
    
  } catch (e) {
    Logger.log(`createAndSendChart 發生錯誤: ${e.message}`);
    return false;
  }
}

/**
 * 生成圖表說明文字
 * @param {string} contract - 契約代碼
 * @param {string} chartType - 圖表類型
 * @param {number} days - 天數
 * @param {number} dataLength - 實際資料筆數
 */
function generateChartCaption(contract, chartType, days, dataLength) {
  const contractName = CONTRACT_NAMES[contract] || contract;
  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy年MM月dd日');
  
  return `<b>📊 ${contractName} ${chartType}圖表</b>\n\n` +
         `<b>📅 資料期間：</b>最近 ${dataLength} 天\n` +
         `<b>📈 圖表類型：</b>${chartType}\n` +
         `<b>🔍 顯示內容：</b>\n` +
         (chartType === '交易量' ? 
           `• 🔴 多方交易口數\n• 🟢 空方交易口數\n• 🔵 多空淨額交易口數` :
           `• 🔴 多方未平倉口數\n• 🟢 空方未平倉口數\n• 🔵 多空淨額未平倉口數`) +
         `\n\n<b>⏰ 生成時間：</b>${today}\n` +
         `<i>💡 資料來源：台灣期貨交易所</i>`;
}

/**
 * 發送所有期貨合約的圖表
 */
function sendAllContractCharts() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // 確認是否繼續
    const confirmResponse = ui.alert(
      '發送所有期貨圖表',
      `將為所有期貨合約生成並發送圖表：\n${CONTRACTS.join(', ')}\n\n` +
      `每個合約會生成交易量和未平倉量兩個圖表\n` +
      `總共約需要 ${CONTRACTS.length * 2} 張圖表\n\n` +
      `是否繼續？`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      return;
    }
    
    let totalSuccess = 0;
    let totalFailure = 0;
    const results = [];
    
    // 發送總計開始訊息
    const startMessage = `🚀 <b>開始生成所有期貨合約圖表</b>\n\n` +
                        `📊 合約清單：${CONTRACTS.map(c => CONTRACT_NAMES[c]).join('、')}\n` +
                        `⏰ 開始時間：${Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')}`;
    sendTelegramMessage(startMessage);
    
    // 為每個合約生成並發送圖表
    for (let i = 0; i < CONTRACTS.length; i++) {
      const contract = CONTRACTS[i];
      Logger.log(`處理第 ${i + 1}/${CONTRACTS.length} 個合約: ${contract}`);
      
      try {
        const success = createAndSendChart(contract, 30, true, false);
        
        if (success) {
          totalSuccess++;
          results.push(`✅ ${CONTRACT_NAMES[contract]} - 成功`);
          Logger.log(`${contract} 圖表發送成功`);
        } else {
          totalFailure++;
          results.push(`❌ ${CONTRACT_NAMES[contract]} - 失敗`);
          Logger.log(`${contract} 圖表發送失敗`);
        }
        
      } catch (e) {
        totalFailure++;
        results.push(`❌ ${CONTRACT_NAMES[contract]} - 錯誤: ${e.message}`);
        Logger.log(`處理 ${contract} 時發生錯誤: ${e.message}`);
      }
      
      // 合約間延遲，避免過於頻繁的請求
      if (i < CONTRACTS.length - 1) {
        Logger.log('等待5秒再處理下一個合約...');
        Utilities.sleep(5000);
      }
    }
    
    // 發送完成總結
    const summaryMessage = `🎯 <b>期貨圖表生成完成</b>\n\n` +
                          `📊 <b>統計結果：</b>\n` +
                          `• 成功：${totalSuccess} 個合約\n` +
                          `• 失敗：${totalFailure} 個合約\n\n` +
                          `📋 <b>詳細結果：</b>\n${results.join('\n')}\n\n` +
                          `⏰ <b>完成時間：</b>${Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')}`;
    
    sendTelegramMessage(summaryMessage);
    
    // 在Google Sheets中也顯示結果
    ui.alert(
      '圖表發送完成',
      `圖表生成並發送完成！\n\n` +
      `成功：${totalSuccess} 個合約\n` +
      `失敗：${totalFailure} 個合約\n\n` +
      `詳細結果已發送到Telegram，請查看。`,
      ui.ButtonSet.OK
    );
    
    return {
      totalSuccess: totalSuccess,
      totalFailure: totalFailure,
      results: results
    };
    
  } catch (e) {
    Logger.log(`sendAllContractCharts 發生錯誤: ${e.message}`);
    SpreadsheetApp.getUi().alert(`發送圖表時發生錯誤: ${e.message}`);
    return null;
  }
}

/**
 * 快速發送圖表
 */
function quickSendChart() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // 確認 Telegram 設定
    if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) {
      ui.alert(
        '缺少設定',
        '請先在 Code.gs 檔案中設定 TELEGRAM_BOT_TOKEN 和 TELEGRAM_CHAT_ID',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 選擇要發送的合約
    const contractResponse = ui.prompt(
      '選擇合約',
      '請輸入需要發送圖表的合約代碼：\n' +
      'TX = 臺股期貨\n' +
      'TE = 電子期貨\n' +
      'MTX = 小型臺指期貨\n' +
      'ZMX = 微型臺指期貨\n' +
      'NQF = 美國那斯達克100期貨\n\n' +
      '輸入 ALL 發送所有合約\n' +
      '多個合約請用逗號分隔，如：TX,TE,MTX',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (contractResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const input = contractResponse.getResponseText().trim().toUpperCase();
    let targetContracts = [];
    
    if (input === 'ALL') {
      targetContracts = CONTRACTS;
    } else {
      const inputContracts = input.split(',').map(c => c.trim());
      targetContracts = inputContracts.filter(c => CONTRACTS.includes(c));
      
      if (targetContracts.length === 0) {
        ui.alert('沒有找到有效的合約代碼！');
        return;
      }
    }
    
    // 檢查資料是否存在
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 使用與 createContractChart 相同的智能工作表查找邏輯
    let sheet = null;
    const possibleSheetNames = [
      ALL_CONTRACTS_SHEET_NAME,
      '所有期貨資料',
      'All_Contracts',
      'FuturesData'
    ];
    
    const allSheets = ss.getSheets();
    
    // 嘗試找到資料工作表
    for (const sheetName of possibleSheetNames) {
      sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        Logger.log(`快速發送圖表：找到資料工作表：${sheetName}`);
        break;
      }
    }
    
    // 如果還是找不到，嘗試找包含期貨資料的工作表
    if (!sheet) {
      for (const s of allSheets) {
        const name = s.getName().toLowerCase();
        if (name.includes('期貨') || name.includes('futures') || name.includes('contracts') || name.includes('data')) {
          sheet = s;
          Logger.log(`快速發送圖表：找到可能的資料工作表：${s.getName()}`);
          break;
        }
      }
    }
    
    // 如果仍然找不到，使用第一個有資料的工作表
    if (!sheet && allSheets.length > 0) {
      for (const s of allSheets) {
        if (s.getLastRow() > 5 && s.getLastColumn() > 5) {
          sheet = s;
          Logger.log(`快速發送圖表：使用第一個有資料的工作表：${s.getName()}`);
          break;
        }
      }
    }
    
    if (!sheet || sheet.getLastRow() <= 1) {
      ui.alert(
        '沒有資料',
        '找不到期貨資料工作表或工作表中沒有資料！\n\n請先執行資料爬取。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 檢查目標合約是否有資料
    const data = sheet.getDataRange().getValues();
    const existingContracts = [...new Set(data.slice(1).map(row => row[1]).filter(c => c))];
    const availableContracts = targetContracts.filter(c => existingContracts.includes(c));
    
    if (availableContracts.length === 0) {
      ui.alert(
        '沒有可用資料',
        `選定的合約沒有資料：${targetContracts.join(', ')}\n\n` +
        `工作表中存在的合約：${existingContracts.join(', ')}\n\n` +
        `請先爬取相關資料。`,
        ui.ButtonSet.OK
      );
      return;
    }
    
    if (availableContracts.length < targetContracts.length) {
      const missingContracts = targetContracts.filter(c => !availableContracts.includes(c));
      const confirmResponse = ui.alert(
        '部分合約沒有資料',
        `以下合約沒有資料：${missingContracts.join(', ')}\n\n` +
        `將只發送有資料的合約：${availableContracts.join(', ')}\n\n是否繼續？`,
        ui.ButtonSet.YES_NO
      );
      
      if (confirmResponse !== ui.Button.YES) {
        return;
      }
    }
    
    // 選擇天數
    const daysResponse = ui.prompt(
      '選擇天數',
      '請輸入需要顯示的天數（預設：7天）:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (daysResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const days = parseInt(daysResponse.getResponseText().trim()) || 7;
    
    if (days <= 0 || days > 365) {
      ui.alert('天數必須在1到365之間！');
      return;
    }
    
    // 確認發送
    const confirmMessage = 
      `準備發送圖表：\n\n` +
      `合約：${availableContracts.join(', ')}\n` +
      `天數：${days}天\n` +
      `資料來源：「${sheet.getName()}」工作表\n\n` +
      `是否確定發送到Telegram？`;
    
    const finalConfirm = ui.alert('確認發送', confirmMessage, ui.ButtonSet.YES_NO);
    
    if (finalConfirm !== ui.Button.YES) {
      return;
    }
    
    // 發送圖表
    let successCount = 0;
    let failureCount = 0;
    let details = [];
    
    for (const contract of availableContracts) {
      try {
        Logger.log(`開始為 ${contract} 生成並發送圖表`);
        
        // 先創建圖表
        createContractChart(contract, days);
        
        // 等待圖表創建完成
        Utilities.sleep(2000);
        
        // 找到圖表工作表
        const chartSheetName = `圖表_${contract}`;
        const chartSheet = ss.getSheetByName(chartSheetName);
        
        if (!chartSheet) {
          Logger.log(`找不到 ${contract} 的圖表工作表`);
          failureCount++;
          details.push(`❌ ${contract}: 圖表創建失敗`);
          continue;
        }
        
        // 獲取圖表
        const charts = chartSheet.getCharts();
        if (charts.length === 0) {
          Logger.log(`${contract} 工作表中沒有圖表`);
          failureCount++;
          details.push(`❌ ${contract}: 找不到圖表`);
          continue;
        }
        
        const chart = charts[0];
        const chartBlob = chart.getAs('image/png');
        
        if (!chartBlob || chartBlob.getBytes().length < 1000) {
          Logger.log(`${contract} 圖表圖片生成失敗或太小`);
          failureCount++;
          details.push(`❌ ${contract}: 圖片生成失敗`);
          continue;
        }
        
        // 準備說明文字
        const caption = 
          `📊 ${CONTRACT_NAMES[contract]} (${contract}) 交易分析\n\n` +
          `📅 資料範圍：最近 ${days} 天\n` +
          `📈 包含交易量和未平倉量趨勢\n` +
          `🕐 生成時間：${Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm')}`;
        
        // 發送到 Telegram
        const telegramSuccess = sendTelegramPhoto(chartBlob, caption);
        
        if (telegramSuccess) {
          Logger.log(`${contract} 圖表成功發送到Telegram`);
          successCount++;
          details.push(`✅ ${contract}: 發送成功`);
        } else {
          Logger.log(`${contract} 圖表發送到Telegram失敗`);
          failureCount++;
          details.push(`❌ ${contract}: Telegram發送失敗`);
        }
        
        // 延遲以避免請求過於頻繁
        Utilities.sleep(1000);
        
      } catch (e) {
        Logger.log(`處理 ${contract} 圖表時發生錯誤: ${e.message}`);
        failureCount++;
        details.push(`❌ ${contract}: ${e.message}`);
      }
    }
    
    // 顯示結果
    const resultMessage = 
      `圖表發送完成！\n\n` +
      `成功：${successCount} 個\n` +
      `失敗：${failureCount} 個\n\n` +
      `詳細結果：\n${details.join('\n')}`;
    
    ui.alert('發送結果', resultMessage, ui.ButtonSet.OK);
    
    // 如果有成功發送，也發送總結訊息
    if (successCount > 0) {
      const summaryMessage = 
        `🚀 快速圖表發送完成\n\n` +
        `📊 已發送 ${successCount} 個期貨合約圖表\n` +
        `📅 資料範圍：最近 ${days} 天\n` +
        `🕐 ${Utilities.formatDate(new Date(), 'Asia/Taipei', 'HH:mm')}`;
      
      sendTelegramMessage(summaryMessage);
    }
    
  } catch (e) {
    Logger.log(`快速發送圖表時發生錯誤: ${e.message}`);
    SpreadsheetApp.getUi().alert(`快速發送圖表時發生錯誤: ${e.message}`);
  }
}

/**
 * 創建交易量圖表
 * @param {Sheet} sheet - 圖表工作表
 * @param {number} dataLength - 資料行數
 */
function createVolumeChart(sheet, dataLength) {
  // 計算圖表範圍
  const dataRange = sheet.getRange(1, 1, dataLength + 1, 5); // A1:E[dataLength+1]
  
  // 創建圖表
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    .setPosition(dataLength + 5, 1, 0, 0)
    .setOption('title', '期貨交易量統計')
    .setOption('titleTextStyle', {
      fontSize: 16,
      bold: true
    })
    .setOption('legend', {position: 'top'})
    .setOption('hAxis', {
      title: '日期',
      titleTextStyle: { fontSize: 12 },
      textStyle: { fontSize: 10 },
      slantedText: true,
      slantedTextAngle: 45
    })
    .setOption('vAxis', {
      title: '交易口數',
      titleTextStyle: { fontSize: 12 }
    })
    .setOption('series', {
      0: {color: 'transparent'}, // 隱藏日期列
      1: {color: 'transparent'}, // 隱藏契約名稱列
      2: {color: '#FF6B6B', targetAxisIndex: 0}, // 多方交易口數 - 紅色
      3: {color: '#4ECDC4', targetAxisIndex: 0}, // 空方交易口數 - 青色
      4: {color: '#45B7D1', targetAxisIndex: 0}  // 多空淨額交易口數 - 藍色
    })
    .setOption('width', 800)
    .setOption('height', 500)
    .setOption('backgroundColor', '#FFFFFF')
    .build();
  
  sheet.insertChart(chart);
}

/**
 * 創建未平倉量圖表
 * @param {Sheet} sheet - 圖表工作表
 * @param {number} dataLength - 資料行數
 */
function createOIChart(sheet, dataLength) {
  // 計算圖表範圍 (包含日期、契約以及未平倉口數資料)
  const dataRange = sheet.getRange(1, 1, dataLength + 1, 2)
    .union(sheet.getRange(1, 6, dataLength + 1, 3)); // A1:B[dataLength+1], F1:H[dataLength+1]
  
  // 創建圖表
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dataRange)
    .setPosition(dataLength + 25, 1, 0, 0)
    .setOption('title', '期貨未平倉量統計')
    .setOption('titleTextStyle', {
      fontSize: 16,
      bold: true
    })
    .setOption('legend', {position: 'top'})
    .setOption('hAxis', {
      title: '日期',
      titleTextStyle: { fontSize: 12 },
      textStyle: { fontSize: 10 },
      slantedText: true,
      slantedTextAngle: 45
    })
    .setOption('vAxis', {
      title: '未平倉口數',
      titleTextStyle: { fontSize: 12 }
    })
    .setOption('series', {
      0: {color: 'transparent'}, // 隱藏日期列
      1: {color: 'transparent'}, // 隱藏契約名稱列
      2: {color: '#FF6B6B', targetAxisIndex: 0}, // 多方未平倉口數 - 紅色
      3: {color: '#4ECDC4', targetAxisIndex: 0}, // 空方未平倉口數 - 青色
      4: {color: '#45B7D1', targetAxisIndex: 0}  // 多空淨額未平倉口數 - 藍色
    })
    .setOption('width', 800)
    .setOption('height', 500)
    .setOption('backgroundColor', '#FFFFFF')
    .build();
  
  sheet.insertChart(chart);
}

/**
 * 發送Telegram圖片（如果Code.gs中沒有此函數）
 */
function sendTelegramPhoto(imageBlob, caption = '') {
  try {
    if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) {
      Logger.log('Telegram設定不完整，無法發送圖片');
      return false;
    }
    
    const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendPhoto`;
    
    const payload = {
      'chat_id': TELEGRAM_CHAT_ID,
      'photo': imageBlob,
      'caption': caption,
      'parse_mode': 'HTML'
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      payload: payload,
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode === 200) {
      const data = JSON.parse(responseText);
      if (data.ok) {
        Logger.log('Telegram圖片發送成功');
        return true;
      } else {
        Logger.log(`Telegram API錯誤: ${data.description}`);
        return false;
      }
    } else {
      Logger.log(`Telegram HTTP錯誤: ${responseCode}, ${responseText}`);
      return false;
    }
    
  } catch (error) {
    Logger.log(`發送Telegram圖片時發生錯誤: ${error.toString()}`);
    return false;
  }
}

/**
 * 顯示圖表生成對話框
 */
function showChartDialog() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // 選擇契約
    const contractResponse = ui.prompt(
      '選擇契約',
      '請輸入需要生成圖表的契約代碼：\n' +
      'TX = 臺股期貨\n' +
      'TE = 電子期貨\n' +
      'MTX = 小型臺指期貨\n' +
      'ZMX = 微型臺指期貨\n' +
      'NQF = 美國那斯達克100期貨',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (contractResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const contract = contractResponse.getResponseText().trim().toUpperCase();
    
    // 驗證契約代碼
    if (!CONTRACTS.includes(contract)) {
      ui.alert('契約代碼無效！請輸入有效的契約代碼：TX, TE, MTX, ZMX, NQF');
      return;
    }
    
    // 選擇天數
    const daysResponse = ui.prompt(
      '選擇天數',
      '請輸入需要顯示的天數 (從最近的一天開始往前推算)：',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (daysResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const days = parseInt(daysResponse.getResponseText().trim());
    
    // 驗證天數
    if (isNaN(days) || days <= 0) {
      ui.alert('天數無效！請輸入大於 0 的數字');
      return;
    }
    
    // 生成圖表
    createContractChart(contract, days);
    
  } catch (e) {
    Logger.log(`顯示圖表對話框時發生錯誤: ${e.message}`);
    SpreadsheetApp.getUi().alert(`顯示圖表對話框時發生錯誤: ${e.message}`);
  }
}

// 將圖表功能添加到選單
function addChartMenuItems() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('📊 圖表分析');
  
  menu.addItem('📈 創建期貨交易圖表', 'showChartDialog');
  menu.addSeparator();
  menu.addItem('🚀 快速發送圖表', 'quickSendChart');
  menu.addItem('📤 發送所有合約圖表', 'sendAllContractCharts');
  menu.addSeparator();
  
  // 為每個契約添加直接生成圖表的選項
  CONTRACTS.forEach(contract => {
    menu.addItem(`📊 生成 ${contract} 最近30天圖表`, `createChart${contract}`);
  });
  
  menu.addSeparator();
  
  // 為每個契約添加直接發送圖表的選項
  CONTRACTS.forEach(contract => {
    menu.addItem(`📤 發送 ${contract} 圖表`, `sendChart${contract}`);
  });
  
  menu.addToUi();
}

// 為每個契約生成圖表的快捷函數
function createChartTX() {
  createContractChart('TX', 30);
}

function createChartTE() {
  createContractChart('TE', 30);
}

function createChartMTX() {
  createContractChart('MTX', 30);
}

function createChartZMX() {
  createContractChart('ZMX', 30);
}

function createChartNQF() {
  createContractChart('NQF', 30);
}

// 為每個契約發送圖表的快捷函數
function sendChartTX() {
  createAndSendChart('TX', 30, true, false);
}

function sendChartTE() {
  createAndSendChart('TE', 30, true, false);
}

function sendChartMTX() {
  createAndSendChart('MTX', 30, true, false);
}

function sendChartZMX() {
  createAndSendChart('ZMX', 30, true, false);
}

function sendChartNQF() {
  createAndSendChart('NQF', 30, true, false);
} 