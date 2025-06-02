// 專門用於解析台灣期貨交易所的 HTML 回應

/**
 * 解析台灣期貨交易所的 HTML 回應，提取期貨交易資料
 * @param {string} htmlContent - 從 API 獲取的 HTML 內容
 * @param {string} contract - 需要提取的契約代碼 (TX, TE, MTX, ZMX, NQF)
 * @param {string} dateStr - 日期字串，格式為 'yyyy/MM/dd'
 * @return {Array} 包含提取資料的數組，如果沒有找到資料則返回 null
 */
function parseContractData(htmlContent, contract, dateStr) {
  try {
    // 記錄輸入參數以便調試
    Logger.log(`開始解析 ${contract} 在 ${dateStr} 的資料`);
    Logger.log(`HTML內容長度: ${htmlContent ? htmlContent.length : 0} 字符`);
    
    // 檢查HTML內容是否有效
    if (!htmlContent || htmlContent.length < 100) {
      Logger.log('HTML內容太短或為空，可能未能正確獲取頁面');
      return null;
    }
    
    // 先檢查是否有明確的錯誤消息或提示
    if (hasErrorMessage(htmlContent)) {
      Logger.log('頁面包含錯誤信息');
      return null;
    }
    
    // 檢查是否為「查無資料」頁面
    if (hasNoDataMessage(htmlContent)) {
      Logger.log('頁面顯示無交易資料');
      return null;
    }
    
    // 檢查是否包含契約關鍵字 - 更寬泛的檢查
    const contractKeywords = [contract];
    let contractFound = false;
    
    // 為特定契約添加額外關鍵字
    if (contract === 'TX') contractKeywords.push('臺股期貨', '台指期');
    if (contract === 'TE') contractKeywords.push('電子期貨', '電子期');
    if (contract === 'MTX') contractKeywords.push('小型臺指', '小臺指');
    if (contract === 'ZMX') contractKeywords.push('微型臺指');
    if (contract === 'NQF') contractKeywords.push('美國那斯達克', '那斯達克');
    
    for (const keyword of contractKeywords) {
      if (htmlContent.includes(keyword)) {
        contractFound = true;
        Logger.log(`在頁面中找到契約關鍵字: ${keyword}`);
        break;
      }
    }
    
    if (!contractFound) {
      Logger.log(`頁面中未找到契約 ${contract} 的任何關鍵字`);
      // 不立即返回null，因為有些頁面可能使用不同的表示方式
    }
    
    // 嘗試檢測頁面中的所有表格並找出最可能包含數據的那個
    // 使用多種正則表達式尋找表格
    const tableRegexList = [
      /<table class="table_f"[\s\S]*?<\/table>/gi,  // 標準表格
      /<table class="table_c"[\s\S]*?<\/table>/gi,  // 替代表格樣式
      /<table[^>]*>[\s\S]*?<\/table>/gi  // 任何表格（最後嘗試）
    ];
    
    let allTables = [];
    
    // 收集所有匹配的表格
    for (let regex of tableRegexList) {
      const tables = htmlContent.match(regex) || [];
      allTables = allTables.concat(tables);
    }
    
    // 去除重複表格
    const uniqueTables = [...new Set(allTables)];
    Logger.log(`找到 ${uniqueTables.length} 個唯一表格`);
    
    if (uniqueTables.length === 0) {
      Logger.log('未找到任何表格');
      return null;
    }
    
    // 評分表格 - 根據可能包含所需數據的特徵
    const tableScores = uniqueTables.map((table, index) => {
      let score = 0;
      
      // 檢查是否包含契約標識符
      for (const keyword of contractKeywords) {
        if (table.includes(keyword)) score += 10;
      }
      
      // 檢查是否包含交易和未平倉資訊
      if (table.includes('交易口數')) score += 5;
      if (table.includes('未平倉')) score += 5;
      if (table.includes('多方') && table.includes('空方')) score += 5;
      if (table.includes('序號') && table.includes('商品名稱')) score += 3;
      if (table.includes('身份別')) score += 3;  // 新增檢查身份別
      if (table.includes('交易口數與契約金額') && table.includes('未平倉餘額')) score += 10;  // 新表格的最強識別特徵
      
      // 檢查表格大小 - 更大的表格可能包含更多資訊
      score += Math.min(5, table.length / 5000); // 最多加5分
      
      // 檢查表格的格式 - 期望的表格通常有更多的行和列
      const rows = table.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];
      score += Math.min(5, rows.length / 3); // 更多行獲得更高分數，最多5分
      
      return { index, score, table };
    });
    
    // 按分數排序表格，選擇最高分的
    tableScores.sort((a, b) => b.score - a.score);
    
    // 記錄排名靠前的表格
    tableScores.slice(0, 3).forEach(t => {
      Logger.log(`表格 #${t.index} 得分: ${t.score}`);
    });
    
    if (tableScores.length === 0 || tableScores[0].score < 5) {
      Logger.log('沒有找到可能包含目標數據的表格');
      return null;
    }
    
    // 使用得分最高的表格
    const dataTable = tableScores[0].table;
    
    // 從表格中尋找符合條件的行
    const rows = dataTable.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];
    Logger.log(`最佳表格中有 ${rows.length} 行`);
    
    if (rows.length < 2) { // 至少需要表頭行和一行數據
      Logger.log('表格行數不足');
      return null;
    }
    
    // 處理新的表格結構 - 檢查是否為新版複雜表頭
    let isNewTableFormat = false;
    
    // 多種方式檢測新表格格式
    const hasComplexHeader = dataTable.includes('交易口數與契約金額') && dataTable.includes('未平倉餘額');
    const hasIdentityColumn = dataTable.includes('身份別');
    const has15Columns = rows.length > 1 && rows[1].match(/<td[^>]*>[\s\S]*?<\/td>/gi) && rows[1].match(/<td[^>]*>[\s\S]*?<\/td>/gi).length >= 14;
    
    if (hasComplexHeader || hasIdentityColumn || has15Columns) {
      Logger.log('檢測到新的表格格式: ' + 
                (hasComplexHeader ? '包含複合表頭 ' : '') + 
                (hasIdentityColumn ? '包含身份別欄位 ' : '') + 
                (has15Columns ? '欄位數量充足' : ''));
      isNewTableFormat = true;
    }
    
    // 根據表格格式選擇不同的解析邏輯
    if (isNewTableFormat) {
      return parseNewTableFormat(dataTable, rows, contract, dateStr, contractKeywords);
    } else {
      return parseStandardTableFormat(dataTable, rows, contract, dateStr, contractKeywords);
    }
    
  } catch (e) {
    Logger.log(`解析資料時發生錯誤: ${e.message}`);
    Logger.log(`錯誤堆疊: ${e.stack}`);
    return null;
  }
}

/**
 * 解析新的表格格式 (包含交易口數與契約金額和未平倉餘額的複雜結構)
 */
function parseNewTableFormat(dataTable, rows, contract, dateStr, contractKeywords) {
  try {
    Logger.log('使用新表格格式解析邏輯');
    
    // 處理表頭 - 新的表格有複雜的多層表頭
    // 第一層表頭通常是大類別
    // 第二層表頭是多方/空方/多空淨額
    // 第三層表頭是口數/契約金額
    
    let headerRows = [];
    let dataRows = [];
    let headerRowsCount = 0;
    
    // 識別表頭行和數據行
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].includes('<th') || rows[i].includes('交易口數') || rows[i].includes('未平倉')) {
        headerRows.push(rows[i]);
        headerRowsCount++;
      } else {
        dataRows.push(rows[i]);
      }
    }
    
    Logger.log(`識別到 ${headerRowsCount} 個表頭行和 ${dataRows.length} 個數據行`);
    
    if (headerRows.length === 0 || dataRows.length === 0) {
      Logger.log('無法識別表頭行或數據行');
      return null;
    }
    
    // 尋找契約行 - 遍歷所有數據行
    let targetRow = null;
    let bestMatchScore = 0;
    
    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      let currentScore = 0;
      
      // 檢查此行是否包含任何契約關鍵字
      for (const keyword of contractKeywords) {
        if (row.includes(keyword)) {
          currentScore += 5;
          Logger.log(`在第 ${i} 行找到契約 ${contract} 關鍵字: ${keyword}`);
        }
      }
      
      // 判斷是否包含期貨或自營商、投信等關鍵字
      if (row.includes('期貨')) currentScore += 2;
      if (row.includes('自營商')) currentScore += 2;
      if (row.includes('投信')) currentScore += 2;
      if (row.includes('外資')) currentScore += 2;
      
      // 判斷是否是目標契約的行數據
      if (currentScore > bestMatchScore) {
        bestMatchScore = currentScore;
        targetRow = row;
        Logger.log(`更新最佳匹配行為第 ${i} 行，得分: ${currentScore}`);
      }
    }
    
    if (!targetRow) {
      Logger.log(`未找到包含契約 ${contract} 的行`);
      return null;
    }
    
    Logger.log(`找到最可能的契約行，得分: ${bestMatchScore}`);
    
    // 解析行中的所有單元格
    const cells = targetRow.match(/<td[^>]*>([\s\S]*?)<\/td>/gi) || [];
    
    if (cells.length < 7) { // 至少需要足夠的單元格
      Logger.log(`行中單元格數量不足: ${cells.length}`);
      return null;
    }
    
    // 清理單元格內容
    const cellContents = cells.map(cell => {
      return cell.replace(/<[^>]*>/g, '').trim().replace(/\s+/g, ' ');
    });
    
    Logger.log(`找到 ${cellContents.length} 個單元格`);
    Logger.log(`單元格內容: ${cellContents.join(' | ')}`);
    
    // 新表格格式下的單元格位置 (根據用戶提供的精確索引)
    // 0: 序號
    // 1: 商品名稱
    // 2: 身份別
    // 3: 多方交易口數
    // 4: 多方契約金額
    // 5: 空方交易口數
    // 6: 空方契約金額
    // 7: 多空淨額交易口數
    // 8: 多空淨額契約金額
    // 9: 多方未平倉口數
    // 10: 多方未平倉契約金額
    // 11: 空方未平倉口數
    // 12: 空方未平倉契約金額
    // 13: 多空淨額未平倉口數
    // 14: 多空淨額未平倉契約金額
    
    // 檢查是否有足夠的欄位
    if (cellContents.length < 14) {
      Logger.log(`單元格數量不足，無法獲取所有數據: ${cellContents.length} < 14`);
      
      // 嘗試按現有數量獲取最多的數據
      const dataArray = [
        dateStr,                              // 日期
        contract,                             // 契約代碼
        cellContents.length > 2 ? cellContents[2] || '' : '',  // 身份別
        cellContents.length > 3 ? parseNumber(cellContents[3] || '0') : 0,  // 多方交易口數
        cellContents.length > 4 ? parseNumber(cellContents[4] || '0') : 0,  // 多方契約金額
        cellContents.length > 5 ? parseNumber(cellContents[5] || '0') : 0,  // 空方交易口數
        cellContents.length > 6 ? parseNumber(cellContents[6] || '0') : 0,  // 空方契約金額
        cellContents.length > 7 ? parseNumber(cellContents[7] || '0') : 0,  // 多空淨額交易口數
        cellContents.length > 8 ? parseNumber(cellContents[8] || '0') : 0,  // 多空淨額契約金額
        cellContents.length > 9 ? parseNumber(cellContents[9] || '0') : 0,  // 多方未平倉口數
        cellContents.length > 10 ? parseNumber(cellContents[10] || '0') : 0,  // 多方未平倉契約金額
        cellContents.length > 11 ? parseNumber(cellContents[11] || '0') : 0,  // 空方未平倉口數
        cellContents.length > 12 ? parseNumber(cellContents[12] || '0') : 0,  // 空方未平倉契約金額
        cellContents.length > 13 ? parseNumber(cellContents[13] || '0') : 0,  // 多空淨額未平倉口數
        cellContents.length > 14 ? parseNumber(cellContents[14] || '0') : 0   // 多空淨額未平倉契約金額
      ];
      
      Logger.log(`部分數據解析結果: ${JSON.stringify(dataArray)}`);
      return dataArray;
    }
    
    // 提取完整資料 - 包括口數、契約金額和身份別
    const dataArray = [
      dateStr,                                 // 日期
      contract,                                // 契約代碼
      cellContents[2] || '',                   // 身份別 (索引2)
      parseNumber(cellContents[3]),           // 多方交易口數 (索引3)
      parseNumber(cellContents[4]),           // 多方契約金額 (索引4)
      parseNumber(cellContents[5]),           // 空方交易口數 (索引5)
      parseNumber(cellContents[6]),           // 空方契約金額 (索引6)
      parseNumber(cellContents[7]),           // 多空淨額交易口數 (索引7)
      parseNumber(cellContents[8]),           // 多空淨額契約金額 (索引8)
      parseNumber(cellContents[9]),           // 多方未平倉口數 (索引9)
      parseNumber(cellContents[10]),          // 多方未平倉契約金額 (索引10)
      parseNumber(cellContents[11]),          // 空方未平倉口數 (索引11)
      parseNumber(cellContents[12]),          // 空方未平倉契約金額 (索引12)
      parseNumber(cellContents[13]),          // 多空淨額未平倉口數 (索引13)
      parseNumber(cellContents[14]),          // 多空淨額未平倉契約金額 (索引14)
    ];
    
    Logger.log(`新表格格式解析結果: ${JSON.stringify(dataArray)}`);
    return dataArray;
    
  } catch (e) {
    Logger.log(`解析新表格格式時發生錯誤: ${e.message}`);
    Logger.log(`錯誤堆疊: ${e.stack}`);
    return null;
  }
}

/**
 * 解析標準表格格式 (原有的解析邏輯)
 */
function parseStandardTableFormat(dataTable, rows, contract, dateStr, contractKeywords) {
  try {
    Logger.log('使用標準表格格式解析邏輯');
    
    // 嘗試從表格中提取標題行以識別列的位置
    const headerRow = rows[0];
    const headerCells = headerRow.match(/<th[^>]*>([\s\S]*?)<\/th>/gi) || [];
    
    // 預設的數據索引
    let dataIndices = {
      dateIndex: -1,       // 日期列
      contractIndex: -1,   // 契約名稱列
      buyVolIndex: -1,     // 多方交易口數
      sellVolIndex: -1,    // 空方交易口數
      netVolIndex: -1,     // 多空淨額交易口數
      buyOIIndex: -1,      // 多方未平倉口數
      sellOIIndex: -1,     // 空方未平倉口數
      netOIIndex: -1       // 多空淨額未平倉口數
    };
    
    // 分析表頭找出列的含義
    if (headerCells.length > 0) {
      const headerTexts = headerCells.map(cell => {
        return cell.replace(/<[^>]*>/g, '').trim();
      });
      
      Logger.log(`表頭內容: ${headerTexts.join(' | ')}`);
      
      // 嘗試識別每列的含義
      headerTexts.forEach((text, index) => {
        const lowerText = text.toLowerCase();
        if (lowerText.includes('日期')) {
          dataIndices.dateIndex = index;
        } else if (lowerText.includes('契約') || lowerText.includes('商品')) {
          dataIndices.contractIndex = index;
        } else if (lowerText.includes('多方') && lowerText.includes('交易')) {
          dataIndices.buyVolIndex = index;
        } else if (lowerText.includes('空方') && lowerText.includes('交易')) {
          dataIndices.sellVolIndex = index;
        } else if ((lowerText.includes('多空') || lowerText.includes('淨額')) && lowerText.includes('交易')) {
          dataIndices.netVolIndex = index;
        } else if (lowerText.includes('多方') && lowerText.includes('未平倉')) {
          dataIndices.buyOIIndex = index;
        } else if (lowerText.includes('空方') && lowerText.includes('未平倉')) {
          dataIndices.sellOIIndex = index;
        } else if ((lowerText.includes('多空') || lowerText.includes('淨額')) && lowerText.includes('未平倉')) {
          dataIndices.netOIIndex = index;
        }
      });
    }
    
    // 尋找契約行 - 遍歷所有數據行
    let targetRow = null;
    let targetRowIndex = -1;
    
    for (let i = 1; i < rows.length; i++) { // 跳過表頭行
      const row = rows[i];
      let contractMatch = false;
      
      // 檢查此行是否包含任何契約關鍵字
      for (const keyword of contractKeywords) {
        if (row.includes(keyword)) {
          contractMatch = true;
          Logger.log(`在第 ${i} 行找到契約 ${contract} 關鍵字: ${keyword}`);
          break;
        }
      }
      
      if (contractMatch) {
        targetRow = row;
        targetRowIndex = i;
        break;
      }
    }
    
    if (!targetRow) {
      Logger.log(`未找到包含契約 ${contract} 的行`);
      return null;
    }
    
    // 解析行中的所有單元格
    const cells = targetRow.match(/<td[^>]*>([\s\S]*?)<\/td>/gi) || [];
    
    if (cells.length < 4) { // 至少需要足夠的單元格
      Logger.log(`行中單元格數量不足: ${cells.length}`);
      return null;
    }
    
    // 清理單元格內容
    const cellContents = cells.map(cell => {
      return cell.replace(/<[^>]*>/g, '').trim().replace(/\s+/g, ' ');
    });
    
    Logger.log(`找到 ${cellContents.length} 個單元格`);
    Logger.log(`單元格內容: ${cellContents.join(' | ')}`);
    
    // 如果表頭分析不成功，嘗試根據常見模式推斷各列的含義
    if (dataIndices.buyVolIndex === -1 || dataIndices.sellVolIndex === -1) {
      Logger.log('從表頭無法識別所有列，嘗試推斷...');
      
      // 一些典型的模式
      if (cellContents.length === 8) {
        // 典型模式：[序號, 契約, 多方交易, 空方交易, 多空淨額交易, 多方未平倉, 空方未平倉, 多空淨額未平倉]
        dataIndices = {
          contractIndex: 1,
          buyVolIndex: 2,
          sellVolIndex: 3,
          netVolIndex: 4,
          buyOIIndex: 5,
          sellOIIndex: 6,
          netOIIndex: 7
        };
      } else if (cellContents.length === 7) {
        // 另一種模式：[契約, 多方交易, 空方交易, 多空淨額交易, 多方未平倉, 空方未平倉, 多空淨額未平倉]
        dataIndices = {
          contractIndex: 0,
          buyVolIndex: 1,
          sellVolIndex: 2,
          netVolIndex: 3,
          buyOIIndex: 4,
          sellOIIndex: 5,
          netOIIndex: 6
        };
      } else {
        // 最後的嘗試：推斷列的含義
        for (let i = 0; i < cellContents.length; i++) {
          const content = cellContents[i].toLowerCase();
          
          // 檢查是否包含契約名稱
          for (const keyword of contractKeywords) {
            if (content.includes(keyword.toLowerCase())) {
              dataIndices.contractIndex = i;
              break;
            }
          }
        }
        
        // 如果找到契約列，推斷其他列
        if (dataIndices.contractIndex !== -1) {
          // 假設契約列之後的列順序為: 多方交易, 空方交易, 淨額交易, 多方未平倉, 空方未平倉, 淨額未平倉
          const startIndex = dataIndices.contractIndex + 1;
          if (startIndex + 5 < cellContents.length) {
            dataIndices.buyVolIndex = startIndex;
            dataIndices.sellVolIndex = startIndex + 1;
            dataIndices.netVolIndex = startIndex + 2;
            dataIndices.buyOIIndex = startIndex + 3;
            dataIndices.sellOIIndex = startIndex + 4;
            dataIndices.netOIIndex = startIndex + 5;
          }
        }
      }
    }
    
    // 最終檢查索引是否都在有效範圍內
    const maxIndex = Math.max(
      dataIndices.buyVolIndex,
      dataIndices.sellVolIndex,
      dataIndices.netVolIndex,
      dataIndices.buyOIIndex,
      dataIndices.sellOIIndex,
      dataIndices.netOIIndex
    );
    
    if (maxIndex >= cellContents.length || maxIndex === -1) {
      Logger.log(`索引範圍無效: 最大索引 ${maxIndex}, 單元格數量 ${cellContents.length}`);
      
      // 最後嘗試：假設數據是按順序排列的簡單模式
      if (cellContents.length >= 6) {
        const baseData = [
          dateStr,
          contract,
          '', // 身份別 (標準格式通常無此數據，保留空字串)
          parseNumber(cellContents[cellContents.length - 6] || '0'), // 多方交易口數
          0, // 多方契約金額 (舊格式無此數據)
          parseNumber(cellContents[cellContents.length - 5] || '0'), // 空方交易口數
          0, // 空方契約金額 (舊格式無此數據)
          parseNumber(cellContents[cellContents.length - 4] || '0'), // 多空淨額交易口數
          0, // 多空淨額契約金額 (舊格式無此數據)
          parseNumber(cellContents[cellContents.length - 3] || '0'), // 多方未平倉口數
          0, // 多方未平倉契約金額 (舊格式無此數據)
          parseNumber(cellContents[cellContents.length - 2] || '0'), // 空方未平倉口數
          0, // 空方未平倉契約金額 (舊格式無此數據)
          parseNumber(cellContents[cellContents.length - 1] || '0'),  // 多空淨額未平倉口數
          0 // 多空淨額未平倉契約金額 (舊格式無此數據)
        ];
        return baseData;
      }
      
      return null;
    }
    
    // 提取資料 - 按照新順序
    const dataArray = [
      dateStr,                                        // 日期
      contract,                                       // 契約代碼
      '',                                             // 身份別 (標準格式通常無此數據，保留空字串)
      parseNumber(cellContents[dataIndices.buyVolIndex]),     // 多方交易口數
      0, // 多方契約金額 (舊格式無此數據)
      parseNumber(cellContents[dataIndices.sellVolIndex]),    // 空方交易口數
      0, // 空方契約金額 (舊格式無此數據)
      parseNumber(cellContents[dataIndices.netVolIndex]),     // 多空淨額交易口數
      0, // 多空淨額契約金額 (舊格式無此數據)
      parseNumber(cellContents[dataIndices.buyOIIndex]),      // 多方未平倉口數
      0, // 多方未平倉契約金額 (舊格式無此數據)
      parseNumber(cellContents[dataIndices.sellOIIndex]),     // 空方未平倉口數
      0, // 空方未平倉契約金額 (舊格式無此數據)
      parseNumber(cellContents[dataIndices.netOIIndex]),       // 多空淨額未平倉口數
      0 // 多空淨額未平倉契約金額 (舊格式無此數據)
    ];
    
    Logger.log(`標準表格格式解析結果: ${JSON.stringify(dataArray)}`);
    return dataArray;
    
  } catch (e) {
    Logger.log(`解析標準表格格式時發生錯誤: ${e.message}`);
    Logger.log(`錯誤堆疊: ${e.stack}`);
    return null;
  }
}

/**
 * 解析表格中的數字，處理千位分隔符並轉換為整數
 * @param {string} str - 包含數字的字串
 * @return {number} 解析後的整數值，如果無法解析則返回 0
 */
function parseNumber(str) {
  if (!str) return 0;
  
  try {
    // 移除所有非數字字符（保留減號表示負數）
    let cleanedStr = str.toString().replace(/[^\d\-\.]/g, '');
    
    // 特殊處理：如果包含小數點但我們只需要整數
    if (cleanedStr.includes('.')) {
      cleanedStr = cleanedStr.split('.')[0];
    }
    
    // 嘗試解析為整數
    const num = parseInt(cleanedStr, 10);
    
    // 檢查是否為有效數字
    return isNaN(num) ? 0 : num;
  } catch (e) {
    Logger.log(`解析數字 "${str}" 時發生錯誤: ${e.message}`);
    return 0;
  }
}

/**
 * 檢查網頁是否顯示「無交易資料」訊息
 * @param {string} htmlContent - 從 API 獲取的 HTML 內容
 * @return {boolean} 如果顯示無交易資料訊息則返回 true，否則返回 false
 */
function hasNoDataMessage(htmlContent) {
  // 檢查常見的「無交易資料」訊息
  const noDataPatterns = [
    '查無資料',
    '無交易資料',
    '尚無資料',
    '無此資料',
    'No data',
    '不存在',
    '不提供',
    '維護中'
  ];
  
  for (const pattern of noDataPatterns) {
    if (htmlContent.includes(pattern)) {
      Logger.log(`找到無資料訊息: "${pattern}"`);
      return true;
    }
  }
  
  return false;
}

/**
 * 檢查網頁是否顯示錯誤訊息
 * @param {string} htmlContent - 從 API 獲取的 HTML 內容
 * @return {boolean} 如果顯示錯誤訊息則返回 true，否則返回 false
 */
function hasErrorMessage(htmlContent) {
  // 更精確地檢查錯誤訊息，避免誤判
  // 檢查常見的錯誤訊息模式，需要在上下文中確認是錯誤訊息而非內容的一部分
  const errorContextPatterns = [
    '系統發生錯誤',
    '查詢發生錯誤',
    '日期錯誤',
    '請輸入正確的日期',
    'Error occurred',
    '資料錯誤',
    '錯誤代碼',
    '伺服器錯誤',
    'Server Error',
    '無法處理您的請求'
  ];
  
  for (const pattern of errorContextPatterns) {
    if (htmlContent.includes(pattern)) {
      Logger.log(`找到明確的錯誤訊息: "${pattern}"`);
      return true;
    }
  }
  
  // 檢查是否有錯誤提示框或錯誤標記
  if (htmlContent.includes('<div class="error">') || 
      htmlContent.includes('<span class="error">') ||
      htmlContent.includes('alert-danger') ||
      htmlContent.includes('error-message')) {
    Logger.log('找到錯誤提示框或錯誤標記');
    return true;
  }
  
  // 檢查是否返回了空數據表或僅有表頭的表格
  const tables = htmlContent.match(/<table[^>]*>[\s\S]*?<\/table>/gi) || [];
  if (tables.length > 0) {
    // 檢查第一個表格是否有數據行
    const firstTable = tables[0];
    const dataRows = firstTable.match(/<tr[^>]*>(?!.*<th[^>]*>)[\s\S]*?<\/tr>/gi) || [];
    
    // 如果表格只有表頭行，沒有數據行
    if (dataRows.length === 0) {
      const tableHeader = firstTable.match(/<th[^>]*>[\s\S]*?<\/th>/gi) || [];
      if (tableHeader.length > 0) {
        Logger.log('表格只有表頭，沒有數據行');
        return false; // 這可能不是錯誤，而是沒有數據
      }
    }
  }
  
  return false; // 沒有明確的錯誤訊息
}

/**
 * 測試函數 - 可以在編輯器中運行以測試解析邏輯
 */
function testParser() {
  // 取得今日日期
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, 'Asia/Taipei', 'yyyy/MM/dd');
  
  // 測試NQF
  testContractParsing('NQF', formattedDate);
}

/**
 * 測試特定契約的解析
 * @param {string} contract - 契約代碼
 * @param {string} dateStr - 日期字串
 */
function testContractParsing(contract, dateStr) {
  Logger.log(`=== 開始測試 ${contract} 在 ${dateStr} 的解析 ===`);
  
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
    Logger.log('成功獲取回應');
    
    // 檢查是否有無資料訊息
    if (hasNoDataMessage(response)) {
      Logger.log('回應中包含無資料訊息');
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
  
  Logger.log(`=== 測試完成 ===`);
}

/**
 * 從台灣期貨交易所獲取資料 - 為測試函數提供
 */
function fetchDataFromTaifex(queryData) {
  // 構建查詢參數字串
  const queryString = Object.keys(queryData)
    .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(queryData[key])}`)
    .join('&');
  
  // 構建完整的 URL
  const url = `https://www.taifex.com.tw/cht/3/futContractsDate?${queryString}`;
  
  // 發送 GET 請求
  const options = {
    'method': 'get',
    'followRedirects': true,
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      return response.getContentText();
    } else {
      Logger.log(`請求失敗，狀態碼: ${responseCode}`);
      return null;
    }
  } catch (e) {
    Logger.log(`發送請求時發生錯誤: ${e.message}`);
    return null;
  }
}

/**
 * 擴展isBusinessDay函數
 */
// 此函數已移至 Code.gs 中，在這裡刪除以避免重複定義 