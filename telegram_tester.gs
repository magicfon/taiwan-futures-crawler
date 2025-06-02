/**
 * Telegram Bot 測試專用檔案
 * 版本：1.0.0
 * 建立日期：2024年12月
 * 
 * 此檔案專門用於測試和除錯Telegram Bot功能
 * 使用前請先在Google Apps Script的設定中添加以下腳本屬性：
 * - TELEGRAM_BOT_TOKEN：您的Telegram Bot Token
 * - TELEGRAM_CHAT_ID：您的Chat ID
 */

// ==================== 基礎測試區域 ====================

/**
 * 極簡基礎測試 - 不依賴任何外部服務
 */
function basicTest() {
  console.log('🔍 極簡基礎測試開始...');
  console.log('時間:', new Date());
  console.log('✅ 函數執行正常');
  return true;
}

/**
 * 權限測試 - 僅測試基本權限
 */
function permissionTest() {
  console.log('🔐 權限測試開始...');
  
  try {
    // 測試1: 基本console.log
    console.log('✅ 測試1: console.log 正常');
    
    // 測試2: Date物件
    const now = new Date();
    console.log('✅ 測試2: Date物件正常 -', now);
    
    // 測試3: PropertiesService (不進行網路請求)
    try {
      const props = PropertiesService.getScriptProperties();
      console.log('✅ 測試3: PropertiesService 存取正常');
      
      // 嘗試讀取屬性
      const testRead = props.getProperty('TELEGRAM_BOT_TOKEN');
      console.log('✅ 測試4: 屬性讀取正常 - Token存在:', !!testRead);
    } catch (e) {
      console.error('❌ 測試3/4: PropertiesService 失敗 -', e.toString());
    }
    
    // 測試5: JSON處理
    try {
      const testJson = JSON.stringify({test: 'value'});
      console.log('✅ 測試5: JSON處理正常');
    } catch (e) {
      console.error('❌ 測試5: JSON處理失敗 -', e.toString());
    }
    
    console.log('🎉 基礎權限測試完成');
    return true;
    
  } catch (error) {
    console.error('❌ 權限測試失敗:', error.toString());
    return false;
  }
}

/**
 * 最基本的網路測試
 */
function networkTest() {
  console.log('🌐 網路測試開始...');
  
  try {
    console.log('準備發送網路請求...');
    
    const response = UrlFetchApp.fetch('https://httpbin.org/get', {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    const statusCode = response.getResponseCode();
    console.log('網路請求狀態碼:', statusCode);
    
    if (statusCode === 200) {
      console.log('✅ 網路連線正常');
      return true;
    } else {
      console.error('❌ 網路請求失敗，狀態碼:', statusCode);
      return false;
    }
    
  } catch (error) {
    console.error('❌ 網路測試異常:', error.toString());
    console.error('可能原因: 權限未授權或網路限制');
    return false;
  }
}

// ==================== 配置檢查區域 ====================

/**
 * 檢查Telegram配置是否正確設定
 */
function checkTelegramConfig() {
  console.log('🔍 開始檢查Telegram配置...');
  
  const botToken = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
  
  console.log('Bot Token存在:', botToken ? '✅ 是' : '❌ 否');
  console.log('Chat ID存在:', chatId ? '✅ 是' : '❌ 否');
  
  if (botToken) {
    console.log('Bot Token格式:', isValidBotToken(botToken) ? '✅ 正確' : '❌ 錯誤');
    console.log('Bot Token長度:', botToken.length);
    console.log('Bot Token前10字符:', botToken.substring(0, 10) + '...');
  }
  
  if (chatId) {
    console.log('Chat ID格式:', isValidChatId(chatId) ? '✅ 正確' : '❌ 錯誤');
    console.log('Chat ID:', chatId);
  }
  
  return {
    hasToken: !!botToken,
    hasChat: !!chatId,
    tokenValid: botToken ? isValidBotToken(botToken) : false,
    chatValid: chatId ? isValidChatId(chatId) : false
  };
}

/**
 * 驗證Bot Token格式
 */
function isValidBotToken(token) {
  // Telegram Bot Token格式: 數字:字母數字組合
  const pattern = /^\d+:[A-Za-z0-9_-]+$/;
  return pattern.test(token);
}

/**
 * 驗證Chat ID格式
 */
function isValidChatId(chatId) {
  // Chat ID可以是負數（群組）或正數（個人）
  const pattern = /^-?\d+$/;
  return pattern.test(chatId);
}

// ==================== 基本連線測試區域 ====================

/**
 * 基本Telegram API連線測試
 */
function testTelegramConnection() {
  console.log('🔗 開始基本連線測試...');
  
  const config = checkTelegramConfig();
  if (!config.hasToken || !config.hasChat) {
    console.error('❌ 配置不完整，請先設定Bot Token和Chat ID');
    return false;
  }
  
  try {
    const botToken = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
    const url = `https://api.telegram.org/bot${botToken}/getMe`;
    
    console.log('📡 發送請求到Telegram API...');
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('HTTP狀態碼:', responseCode);
    console.log('回應內容:', responseText);
    
    if (responseCode === 200) {
      const data = JSON.parse(responseText);
      if (data.ok) {
        console.log('✅ Bot連線成功！');
        console.log('Bot資訊:', JSON.stringify(data.result, null, 2));
        return true;
      } else {
        console.error('❌ API回應錯誤:', data.description);
        return false;
      }
    } else {
      console.error('❌ HTTP錯誤:', responseCode);
      return false;
    }
    
  } catch (error) {
    console.error('❌ 連線測試失敗:', error.toString());
    return false;
  }
}

/**
 * 測試基本訊息發送
 */
function testBasicMessage() {
  console.log('📨 開始基本訊息測試...');
  
  const testMessage = `🔔 Telegram連線測試
時間：${new Date().toLocaleString('zh-TW')}
狀態：測試中...
✅ 如果您看到這則訊息，表示Telegram Bot設定成功！`;

  return sendTelegramMessage(testMessage);
}

// ==================== 訊息發送功能區域 ====================

/**
 * 發送Telegram訊息（主要函數）
 */
function sendTelegramMessage(message, parseMode = 'HTML') {
  try {
    const botToken = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
    const chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
    
    if (!botToken || !chatId) {
      console.error('❌ Telegram配置不完整');
      return false;
    }
    
    const url = `https://api.telegram.org/bot${botToken}/sendMessage`;
    const payload = {
      chat_id: chatId,
      text: message,
      parse_mode: parseMode
    };
    
    console.log('📤 發送訊息到Telegram...');
    console.log('訊息內容:', message.substring(0, 100) + (message.length > 100 ? '...' : ''));
    
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('HTTP狀態碼:', responseCode);
    
    if (responseCode === 200) {
      const data = JSON.parse(responseText);
      if (data.ok) {
        console.log('✅ 訊息發送成功！');
        console.log('訊息ID:', data.result.message_id);
        return true;
      } else {
        console.error('❌ API錯誤:', data.description);
        console.error('錯誤代碼:', data.error_code);
        return false;
      }
    } else {
      console.error('❌ HTTP錯誤:', responseCode);
      console.error('回應內容:', responseText);
      return false;
    }
    
  } catch (error) {
    console.error('❌ 發送訊息時發生錯誤:', error.toString());
    console.error('錯誤堆疊:', error.stack);
    return false;
  }
}

// ==================== 格式化測試區域 ====================

/**
 * 測試HTML格式化
 */
function testHTMLFormatting() {
  console.log('🎨 測試HTML格式化...');
  
  const htmlMessage = `<b>🔔 台灣期貨資料爬取成功</b>

<b>📊 資料統計：</b>
• 總筆數：<code>150</code>筆
• 成功率：<code>98.67%</code>
• 處理時間：<code>45.2</code>秒

<b>📈 主要商品：</b>
• 台指期 (TXF)：<code>16,850</code> 點
• 小台指 (MXF)：<code>16,845</code> 點
• 電子期 (EXF)：<code>825.5</code> 點

<b>⏰ 執行時間：</b>
<code>2024-12-19 14:30:25</code>

<i>✅ 資料已成功寫入Google試算表</i>`;

  return sendTelegramMessage(htmlMessage, 'HTML');
}

/**
 * 測試不同訊息格式
 */
function testDifferentFormats() {
  console.log('📝 測試不同訊息格式...');
  
  const tests = [
    {
      name: '純文字格式',
      message: '這是純文字測試訊息\n時間：' + new Date().toLocaleString('zh-TW'),
      parseMode: null
    },
    {
      name: 'Markdown格式',
      message: '*粗體文字* 和 _斜體文字_\n`程式碼` 和 ```程式碼區塊```',
      parseMode: 'Markdown'
    },
    {
      name: 'HTML格式',
      message: '<b>粗體</b> 和 <i>斜體</i>\n<code>程式碼</code> 和 <pre>程式碼區塊</pre>',
      parseMode: 'HTML'
    }
  ];
  
  let results = [];
  
  for (let test of tests) {
    console.log(`測試 ${test.name}...`);
    const success = sendTelegramMessage(test.message, test.parseMode);
    results.push({
      name: test.name,
      success: success
    });
    
    // 間隔1秒避免API限制
    Utilities.sleep(1000);
  }
  
  console.log('📊 測試結果：');
  results.forEach(result => {
    console.log(`${result.name}: ${result.success ? '✅ 成功' : '❌ 失敗'}`);
  });
  
  return results;
}

// ==================== 模擬實際使用區域 ====================

/**
 * 模擬爬取成功通知
 */
function testSuccessNotification() {
  console.log('✅ 測試成功通知...');
  
  const mockData = {
    totalRecords: 145,
    successRate: 97.24,
    processingTime: 38.7,
    timestamp: new Date(),
    topContracts: [
      { name: '台指期 (TXF)', price: '16,850', change: '+85' },
      { name: '小台指 (MXF)', price: '16,845', change: '+83' },
      { name: '電子期 (EXF)', price: '825.5', change: '+12.3' }
    ]
  };
  
  const message = generateSuccessMessage(mockData);
  return sendTelegramMessage(message, 'HTML');
}

/**
 * 模擬爬取失敗通知
 */
function testFailureNotification() {
  console.log('❌ 測試失敗通知...');
  
  const mockError = {
    errorMessage: '網路連線逾時',
    attemptCount: 2,
    nextRetryTime: new Date(Date.now() + 10 * 60 * 1000),
    timestamp: new Date()
  };
  
  const message = generateFailureMessage(mockError);
  return sendTelegramMessage(message, 'HTML');
}

/**
 * 模擬重試通知
 */
function testRetryNotification() {
  console.log('🔄 測試重試通知...');
  
  const retryInfo = {
    attemptNumber: 2,
    maxAttempts: 3,
    lastError: '資料解析錯誤',
    nextRetryTime: new Date(Date.now() + 10 * 60 * 1000),
    timestamp: new Date()
  };
  
  const message = generateRetryMessage(retryInfo);
  return sendTelegramMessage(message, 'HTML');
}

// ==================== 訊息生成函數區域 ====================

/**
 * 生成成功通知訊息
 */
function generateSuccessMessage(data) {
  const timeStr = data.timestamp.toLocaleString('zh-TW');
  
  let contractsText = '';
  if (data.topContracts && data.topContracts.length > 0) {
    contractsText = data.topContracts.map(contract => 
      `• ${contract.name}：<code>${contract.price}</code> (<code>${contract.change}</code>)`
    ).join('\n');
  }
  
  return `<b>🎉 台灣期貨資料爬取成功</b>

<b>📊 執行統計：</b>
• 總筆數：<code>${data.totalRecords}</code> 筆
• 成功率：<code>${data.successRate}%</code>
• 處理時間：<code>${data.processingTime}</code> 秒

${contractsText ? `<b>📈 主要商品：</b>\n${contractsText}\n\n` : ''}

<b>⏰ 執行時間：</b>
<code>${timeStr}</code>

<i>✅ 資料已成功寫入Google試算表</i>`;
}

/**
 * 生成失敗通知訊息
 */
function generateFailureMessage(error) {
  const timeStr = error.timestamp.toLocaleString('zh-TW');
  const retryTimeStr = error.nextRetryTime.toLocaleString('zh-TW');
  
  return `<b>⚠️ 台灣期貨資料爬取失敗</b>

<b>❌ 錯誤詳情：</b>
<code>${error.errorMessage}</code>

<b>🔄 重試資訊：</b>
• 已嘗試：<code>${error.attemptCount}</code> 次
• 下次重試：<code>${retryTimeStr}</code>

<b>⏰ 發生時間：</b>
<code>${timeStr}</code>

<i>🤖 系統將自動重試...</i>`;
}

/**
 * 生成重試通知訊息
 */
function generateRetryMessage(retry) {
  const timeStr = retry.timestamp.toLocaleString('zh-TW');
  const retryTimeStr = retry.nextRetryTime.toLocaleString('zh-TW');
  
  return `<b>🔄 台灣期貨爬取重試中</b>

<b>📊 重試狀態：</b>
• 第 <code>${retry.attemptNumber}</code> 次重試 (共 <code>${retry.maxAttempts}</code> 次)
• 上次錯誤：<code>${retry.lastError}</code>

<b>⏰ 時間資訊：</b>
• 重試時間：<code>${timeStr}</code>
• 下次重試：<code>${retryTimeStr}</code>

<i>🤖 正在自動重試爬取...</i>`;
}

// ==================== 除錯工具區域 ====================

/**
 * 詳細的網路連線測試
 */
function detailedNetworkTest() {
  console.log('🌐 執行詳細網路測試...');
  
  const tests = [
    {
      name: 'Google連線測試',
      url: 'https://www.google.com',
      method: 'GET'
    },
    {
      name: 'Telegram API基本連線',
      url: 'https://api.telegram.org',
      method: 'GET'
    },
    {
      name: 'Bot API測試',
      url: `https://api.telegram.org/bot${PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN')}/getMe`,
      method: 'GET'
    }
  ];
  
  let results = [];
  
  for (let test of tests) {
    console.log(`測試 ${test.name}...`);
    try {
      const response = UrlFetchApp.fetch(test.url, {
        method: test.method,
        muteHttpExceptions: true
      });
      
      const result = {
        name: test.name,
        success: response.getResponseCode() < 400,
        statusCode: response.getResponseCode(),
        responseTime: new Date().getTime()
      };
      
      results.push(result);
      console.log(`${test.name}: ${result.success ? '✅' : '❌'} (${result.statusCode})`);
      
    } catch (error) {
      results.push({
        name: test.name,
        success: false,
        error: error.toString()
      });
      console.error(`${test.name}: ❌ ${error.toString()}`);
    }
  }
  
  return results;
}

/**
 * 權限檢查
 */
function checkPermissions() {
  console.log('🔐 檢查權限設定...');
  
  const permissions = {
    scriptProperties: false,
    urlFetch: false,
    utilities: false
  };
  
  try {
    // 測試腳本屬性存取
    PropertiesService.getScriptProperties().getProperty('TEST');
    permissions.scriptProperties = true;
    console.log('✅ 腳本屬性存取：正常');
  } catch (error) {
    console.error('❌ 腳本屬性存取：失敗', error.toString());
  }
  
  try {
    // 測試網路請求
    UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
    permissions.urlFetch = true;
    console.log('✅ 網路請求權限：正常');
  } catch (error) {
    console.error('❌ 網路請求權限：失敗', error.toString());
  }
  
  try {
    // 測試工具函數
    Utilities.sleep(100);
    permissions.utilities = true;
    console.log('✅ 工具函數權限：正常');
  } catch (error) {
    console.error('❌ 工具函數權限：失敗', error.toString());
  }
  
  return permissions;
}

/**
 * 緊急簡化測試
 */
function emergencySimpleTest() {
  console.log('🚨 執行緊急簡化測試...');
  
  try {
    const botToken = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
    const chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
    
    if (!botToken) {
      console.error('❌ 找不到TELEGRAM_BOT_TOKEN');
      return false;
    }
    
    if (!chatId) {
      console.error('❌ 找不到TELEGRAM_CHAT_ID');
      return false;
    }
    
    console.log('✅ 配置檢查通過');
    
    const url = `https://api.telegram.org/bot${botToken}/sendMessage`;
    const message = '🔔 緊急測試訊息 ' + new Date().getTime();
    
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        chat_id: chatId,
        text: message
      })
    });
    
    console.log('回應狀態:', response.getResponseCode());
    console.log('回應內容:', response.getContentText());
    
    return response.getResponseCode() === 200;
    
  } catch (error) {
    console.error('❌ 緊急測試失敗:', error.toString());
    return false;
  }
}

// ==================== 完整測試套件區域 ====================

/**
 * 執行完整測試套件
 */
function runFullTestSuite() {
  console.log('🧪 開始執行完整Telegram測試套件...');
  
  const results = {
    config: false,
    connection: false,
    basicMessage: false,
    htmlFormatting: false,
    successNotification: false,
    failureNotification: false,
    retryNotification: false,
    networkTest: [],
    permissions: {}
  };
  
  // 1. 配置檢查
  console.log('\n--- 第1步：配置檢查 ---');
  const config = checkTelegramConfig();
  results.config = config.hasToken && config.hasChat && config.tokenValid && config.chatValid;
  
  if (!results.config) {
    console.error('❌ 配置檢查失敗，終止測試');
    return results;
  }
  
  // 2. 連線測試
  console.log('\n--- 第2步：連線測試 ---');
  results.connection = testTelegramConnection();
  
  if (!results.connection) {
    console.error('❌ 連線測試失敗，終止測試');
    return results;
  }
  
  // 3. 基本訊息測試
  console.log('\n--- 第3步：基本訊息測試 ---');
  results.basicMessage = testBasicMessage();
  Utilities.sleep(2000);
  
  // 4. HTML格式測試
  console.log('\n--- 第4步：HTML格式測試 ---');
  results.htmlFormatting = testHTMLFormatting();
  Utilities.sleep(2000);
  
  // 5. 成功通知測試
  console.log('\n--- 第5步：成功通知測試 ---');
  results.successNotification = testSuccessNotification();
  Utilities.sleep(2000);
  
  // 6. 失敗通知測試
  console.log('\n--- 第6步：失敗通知測試 ---');
  results.failureNotification = testFailureNotification();
  Utilities.sleep(2000);
  
  // 7. 重試通知測試
  console.log('\n--- 第7步：重試通知測試 ---');
  results.retryNotification = testRetryNotification();
  Utilities.sleep(2000);
  
  // 8. 網路測試
  console.log('\n--- 第8步：網路測試 ---');
  results.networkTest = detailedNetworkTest();
  
  // 9. 權限檢查
  console.log('\n--- 第9步：權限檢查 ---');
  results.permissions = checkPermissions();
  
  // 輸出測試結果
  console.log('\n=== 🎯 測試結果總覽 ===');
  console.log(`配置檢查: ${results.config ? '✅' : '❌'}`);
  console.log(`連線測試: ${results.connection ? '✅' : '❌'}`);
  console.log(`基本訊息: ${results.basicMessage ? '✅' : '❌'}`);
  console.log(`HTML格式: ${results.htmlFormatting ? '✅' : '❌'}`);
  console.log(`成功通知: ${results.successNotification ? '✅' : '❌'}`);
  console.log(`失敗通知: ${results.failureNotification ? '✅' : '❌'}`);
  console.log(`重試通知: ${results.retryNotification ? '✅' : '❌'}`);
  
  const allPassed = Object.values(results).every(result => 
    typeof result === 'boolean' ? result : true
  );
  
  console.log(`\n🎉 整體測試結果: ${allPassed ? '✅ 全部通過' : '⚠️ 部分失敗'}`);
  
  return results;
}

// ==================== 使用說明區域 ====================

/**
 * 顯示使用說明
 */
function showUsageInstructions() {
  console.log(`
🤖 Telegram Bot 測試工具使用說明

📋 設定步驟：
1. 建立Telegram Bot：
   • 與 @BotFather 對話
   • 輸入 /newbot
   • 設定Bot名稱和用戶名
   • 複製得到的Token

2. 獲取Chat ID：
   • 與 @userinfobot 對話，獲取您的Chat ID
   • 或者發送訊息給您的Bot，然後造訪：
     https://api.telegram.org/bot<YOUR_TOKEN>/getUpdates

3. 設定腳本屬性：
   • 在Google Apps Script中，點擊「專案設定」
   • 在「腳本屬性」區域添加：
     - TELEGRAM_BOT_TOKEN: 您的Bot Token
     - TELEGRAM_CHAT_ID: 您的Chat ID

4. 執行測試：
   • 儲存設定後，執行 quickDiagnosis()
   • 如果成功，您應該會收到測試訊息

💡 常見問題：
   • Token格式必須包含冒號（:）
   • Chat ID必須是純數字
   • 首次執行需要授權網路存取權限
  `);
}

/**
 * 快速診斷函數 - 針對「不明錯誤」問題
 */
function quickDiagnosis() {
  console.log('🚨 快速診斷開始...');
  console.log('時間:', new Date().toLocaleString('zh-TW'));
  
  try {
    // 步驟1：基本環境檢查
    console.log('\n=== 步驟1：基本環境檢查 ===');
    console.log('Google Apps Script版本: V8');
    console.log('執行環境: 正常');
    
    // 步驟2：權限快速檢查
    console.log('\n=== 步驟2：權限快速檢查 ===');
    try {
      const testProp = PropertiesService.getScriptProperties().getProperty('TEST_QUICK');
      console.log('✅ 腳本屬性存取: 正常');
    } catch (e) {
      console.error('❌ 腳本屬性存取: 失敗 -', e.toString());
    }
    
    // 步驟3：配置檢查
    console.log('\n=== 步驟3：配置檢查 ===');
    const botToken = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
    const chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
    
    if (!botToken) {
      console.error('❌ 關鍵問題：未找到 TELEGRAM_BOT_TOKEN');
      console.error('   請到「專案設定」→「腳本屬性」新增此設定');
      return false;
    }
    
    if (!chatId) {
      console.error('❌ 關鍵問題：未找到 TELEGRAM_CHAT_ID');
      console.error('   請到「專案設定」→「腳本屬性」新增此設定');
      return false;
    }
    
    console.log('✅ Bot Token: 已設定 (長度:', botToken.length, ')');
    console.log('✅ Chat ID: 已設定 (', chatId, ')');
    
    // 步驟4：網路連線檢查
    console.log('\n=== 步驟4：網路連線檢查 ===');
    try {
      const response = UrlFetchApp.fetch('https://www.google.com', {
        method: 'GET',
        muteHttpExceptions: true
      });
      console.log('✅ 基本網路連線: 正常 (狀態碼:', response.getResponseCode(), ')');
    } catch (e) {
      console.error('❌ 基本網路連線: 失敗 -', e.toString());
      return false;
    }
    
    // 步驟5：Telegram API檢查
    console.log('\n=== 步驟5：Telegram API檢查 ===');
    try {
      const telegramUrl = `https://api.telegram.org/bot${botToken}/getMe`;
      const telegramResponse = UrlFetchApp.fetch(telegramUrl, {
        method: 'GET',
        muteHttpExceptions: true
      });
      
      const statusCode = telegramResponse.getResponseCode();
      console.log('Telegram API狀態碼:', statusCode);
      
      if (statusCode === 200) {
        const data = JSON.parse(telegramResponse.getContentText());
        if (data.ok) {
          console.log('✅ Telegram Bot: 連線正常');
          console.log('Bot名稱:', data.result.first_name);
          console.log('Bot用戶名:', data.result.username);
        } else {
          console.error('❌ Telegram Bot: API錯誤 -', data.description);
          return false;
        }
      } else {
        console.error('❌ Telegram Bot: HTTP錯誤 -', statusCode);
        return false;
      }
    } catch (e) {
      console.error('❌ Telegram API連線: 失敗 -', e.toString());
      return false;
    }
    
    // 步驟6：簡單訊息測試
    console.log('\n=== 步驟6：簡單訊息測試 ===');
    try {
      const testMessage = `🔍 快速診斷測試\n時間: ${new Date().toLocaleString('zh-TW')}\n狀態: 所有檢查通過 ✅`;
      
      const sendUrl = `https://api.telegram.org/bot${botToken}/sendMessage`;
      const sendResponse = UrlFetchApp.fetch(sendUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        payload: JSON.stringify({
          chat_id: chatId,
          text: testMessage
        }),
        muteHttpExceptions: true
      });
      
      const sendStatusCode = sendResponse.getResponseCode();
      if (sendStatusCode === 200) {
        console.log('✅ 測試訊息發送: 成功');
        console.log('🎉 診斷結果: 所有功能正常，Telegram Bot已可使用！');
        return true;
      } else {
        console.error('❌ 測試訊息發送: 失敗 (狀態碼:', sendStatusCode, ')');
        console.error('回應內容:', sendResponse.getContentText());
        return false;
      }
    } catch (e) {
      console.error('❌ 測試訊息發送: 異常 -', e.toString());
      return false;
    }
    
  } catch (error) {
    console.error('❌ 診斷過程發生錯誤:', error.toString());
    console.error('錯誤堆疊:', error.stack);
    return false;
  }
}

/**
 * 設定幫助工具
 */
function setupHelper() {
  console.log(`
🛠️ Telegram Bot 設定幫助

📋 如果您還沒有設定，請按照以下步驟：

1️⃣ 建立Telegram Bot：
   • 開啟Telegram，搜尋 @BotFather
   • 傳送 /newbot
   • 依指示設定Bot名稱和用戶名
   • 複製獲得的Token

2️⃣ 獲取Chat ID：
   • 搜尋 @userinfobot 並開始對話
   • 機器人會告訴您的Chat ID
   • 或者先傳訊息給您的Bot，然後造訪：
     https://api.telegram.org/bot<YOUR_TOKEN>/getUpdates

3️⃣ 在Google Apps Script中設定：
   • 點擊左側「專案設定」（齒輪圖示）
   • 向下捲動到「腳本屬性」區域
   • 點擊「新增腳本屬性」
   • 新增兩個屬性：
     - 屬性: TELEGRAM_BOT_TOKEN
       值: 您的Bot Token（如：123456789:ABCdefGHIjklMNOpqrsTUVwxyz）
     - 屬性: TELEGRAM_CHAT_ID  
       值: 您的Chat ID（如：987654321）

4️⃣ 執行測試：
   • 儲存設定後，執行 quickDiagnosis()
   • 如果成功，您應該會收到測試訊息

💡 常見問題：
   • Token格式必須包含冒號（:）
   • Chat ID必須是純數字
   • 首次執行需要授權網路存取權限
  `);
}

// ==================== 圖表生成與發送區域 ====================

/**
 * 從試算表讀取近30個交易日的多空淨額交易口數資料
 */
function getLast30DaysNetVolumeData() {
  try {
    // 修改為使用「所有期貨資料」工作表
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('所有期貨資料') || ss.getSheets()[0];
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      console.log('試算表中沒有資料');
      return [];
    }
    
    // 獲取所有資料
    const allData = sheet.getRange(2, 1, lastRow - 1, 15).getValues(); // 從第2行開始，假設第1行是標題
    
    // 資料結構：
    // 0: 日期, 1: 契約代碼, 2: 身份別, 3: 多方交易口數, 4: 多方契約金額, 
    // 5: 空方交易口數, 6: 空方契約金額, 7: 多空淨額交易口數, ...
    
    // 按日期和契約分組資料
    const groupedData = {};
    
    allData.forEach(row => {
      const dateStr = formatDate(row[0]);
      const contract = row[1];
      const identity = row[2];
      const netVolume = parseFloat(row[7]) || 0;
      
      // 只處理主要契約（TX, MXF, EXF等）
      if (!['TX', 'MXF', 'EXF', 'TXF', 'NQF'].includes(contract)) {
        return;
      }
      
      if (!groupedData[dateStr]) {
        groupedData[dateStr] = {};
      }
      
      if (!groupedData[dateStr][contract]) {
        groupedData[dateStr][contract] = {};
      }
      
      groupedData[dateStr][contract][identity] = netVolume;
    });
    
    // 計算每日各契約的總淨額
    const dailyData = [];
    
    Object.keys(groupedData).sort().slice(-30).forEach(dateStr => {
      const dayData = { date: dateStr };
      
      ['TX', 'MXF', 'EXF', 'TXF', 'NQF'].forEach(contract => {
        let totalNet = 0;
        if (groupedData[dateStr][contract]) {
          Object.values(groupedData[dateStr][contract]).forEach(netVol => {
            totalNet += netVol;
          });
        }
        dayData[contract] = totalNet;
      });
      
      dailyData.push(dayData);
    });
    
    console.log(`成功獲取${dailyData.length}天的資料`);
    return dailyData;
    
  } catch (error) {
    console.error('讀取資料時發生錯誤:', error.toString());
    return [];
  }
}

/**
 * 格式化日期為字串
 */
function formatDate(date) {
  if (typeof date === 'string') return date;
  if (date instanceof Date) {
    return date.toLocaleDateString('zh-TW').replace(/\//g, '-');
  }
  return date.toString();
}

/**
 * 為單一契約建立多空淨額交易口數柱狀圖（使用驗證成功的配置）
 */
function createSingleContractChart(data, contract, title = '') {
  try {
    console.log(`📊 使用驗證成功的配置建立${contract}圖表...`);
    
    if (!data || data.length === 0) {
      console.error('沒有資料可以建立圖表');
      return null;
    }
    
    // 建立臨時試算表
    const tempSS = SpreadsheetApp.create(`圖表_${contract}_${new Date().getTime()}`);
    const sheet = tempSS.getActiveSheet();
    
    // 使用驗證成功的格式
    const headers = ['日期', contract];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 填入資料，確保格式與成功版本一致
    const chartData = data.map(day => [
      day.date,
      day[contract] || 0
    ]);
    
    sheet.getRange(2, 1, chartData.length, headers.length).setValues(chartData);
    
    console.log(`${contract}資料已寫入，筆數: ${chartData.length}`);
    
    // 計算統計
    const values = chartData.map(row => row[1]);
    const maxVal = Math.max(...values);
    const minVal = Math.min(...values);
    const avgVal = values.reduce((sum, val) => sum + val, 0) / values.length;
    
    const contractColors = {
      'TX': '#FF6B6B',
      'MXF': '#4ECDC4',
      'EXF': '#45B7D1',
      'TXF': '#96CEB4',
      'NQF': '#FECA57'
    };
    
    const contractNames = {
      'TX': '台指期',
      'MXF': '小台指',
      'EXF': '電子期',
      'TXF': '台指期貨',
      'NQF': '那斯達克期'
    };
    
    const chartTitle = title || `${contractNames[contract] || contract} 近30個交易日多空淨額`;
    
    // 使用驗證成功的圖表配置
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(1, 1, chartData.length + 1, headers.length))
      .setPosition(1, 4, 0, 0)
      .setOption('title', chartTitle)
      .setOption('titleTextStyle', {
        fontSize: 16,
        bold: true
      })
      .setOption('legend', {
        position: 'none'
      })
      .setOption('hAxis', {
        title: '日期',
        titleTextStyle: { fontSize: 12 },
        textStyle: { fontSize: 10 },
        slantedText: true,
        slantedTextAngle: 45
      })
      .setOption('vAxis', {
        title: '多空淨額交易口數',
        titleTextStyle: { fontSize: 12 }
      })
      .setOption('series', {
        0: { color: contractColors[contract] || '#4285F4' }
      })
      .setOption('width', 800)
      .setOption('height', 500)
      .setOption('backgroundColor', '#FFFFFF')
      .build();
    
    sheet.insertChart(chart);
    
    console.log(`${contract}圖表建立完成，等待渲染...`);
    SpreadsheetApp.flush();
    Utilities.sleep(3000);
    
    const charts = sheet.getCharts();
    if (charts.length === 0) {
      console.error('❌ 圖表建立失敗');
      return null;
    }
    
    console.log(`✅ ${contract}圖表驗證成功`);
    
    return {
      spreadsheet: tempSS,
      sheet: sheet,
      chart: charts[0],
      contract: contract,
      stats: {
        max: maxVal,
        min: minVal,
        average: Math.round(avgVal),
        range: maxVal - minVal
      }
    };
    
  } catch (error) {
    console.error(`❌ 建立${contract}圖表失敗:`, error.toString());
    return null;
  }
}

/**
 * 生成所有契約的單獨圖表並發送到Telegram
 */
function sendAllContractCharts() {
  console.log('🧪 開始生成所有契約的單獨圖表...');
  
  try {
    // 讀取資料（優先使用真實資料，如無則使用模擬資料）
    let data = getLast30DaysNetVolumeData();
    let isRealData = true;
    
    if (!data || data.length === 0) {
      console.log('無法獲取真實資料，使用模擬資料');
      data = generateMockNetVolumeData();
      isRealData = false;
    }
    
    const contracts = ['TX', 'MXF', 'EXF', 'TXF', 'NQF'];
    const contractNames = {
      'TX': '台指期',
      'MXF': '小台指',
      'EXF': '電子期',
      'TXF': '台指期貨',
      'NQF': '那斯達克期'
    };
    
    let successCount = 0;
    let results = [];
    
    // 為每個契約生成圖表
    for (const contract of contracts) {
      console.log(`\n--- 處理 ${contract} (${contractNames[contract]}) ---`);
      
      // 建立圖表
      const chartInfo = createSingleContractChart(data, contract);
      
      if (!chartInfo) {
        console.error(`❌ ${contract} 圖表建立失敗`);
        continue;
      }
      
      // 準備說明訊息
      const caption = generateContractChartCaption(contract, chartInfo.stats, isRealData, data);
      
      // 發送圖表
      const success = sendChartToTelegram(chartInfo, caption);
      
      if (success) {
        console.log(`✅ ${contract} 圖表發送成功`);
        successCount++;
      } else {
        console.error(`❌ ${contract} 圖表發送失敗`);
      }
      
      results.push({
        contract: contract,
        success: success
      });
      
      // 間隔2秒避免API限制
      if (contracts.indexOf(contract) < contracts.length - 1) {
        console.log('等待2秒...');
        Utilities.sleep(2000);
      }
    }
    
    // 發送總結訊息
    const summaryMessage = `<b>📊 期貨圖表生成完成</b>

<b>📈 生成結果：</b>
${results.map(r => `• ${contractNames[r.contract]}: ${r.success ? '✅' : '❌'}`).join('\n')}

<b>📊 統計：</b>
• 成功: <code>${successCount}</code> / <code>${contracts.length}</code>
• 資料來源: <code>${isRealData ? '真實期貨資料' : '模擬測試資料'}</code>

<b>⏰ 完成時間：</b>
<code>${new Date().toLocaleString('zh-TW')}</code>

<i>🎯 已為每個契約生成獨立圖表</i>`;
    
    sendTelegramMessage(summaryMessage, 'HTML');
    
    console.log(`\n🎉 完成！成功生成 ${successCount}/${contracts.length} 個圖表`);
    return results;
    
  } catch (error) {
    console.error('生成圖表過程中發生錯誤:', error.toString());
    return [];
  }
}

/**
 * 生成單一契約圖表的說明文字
 */
function generateContractChartCaption(contract, stats, isRealData, data) {
  const contractNames = {
    'TX': '台指期',
    'MXF': '小台指',
    'EXF': '電子期', 
    'TXF': '台指期貨',
    'NQF': '那斯達克期'
  };
  
  const dateRange = data && data.length > 0 ? 
    `${data[0]?.date} - ${data[data.length - 1]?.date}` : 
    '近30個交易日';
  
  // 根據數值判斷趨勢
  let trend = '';
  if (stats.average > 1000) {
    trend = '📈 偏多格局';
  } else if (stats.average < -1000) {
    trend = '📉 偏空格局';
  } else {
    trend = '⚖️ 相對均衡';
  }
  
  return `<b>📊 ${contractNames[contract] || contract} 多空淨額圖表</b>

<b>📅 資料期間：</b>${dateRange}

<b>📈 統計摘要：</b>
• 最高淨額：<code>${stats.max.toLocaleString()}</code>
• 最低淨額：<code>${stats.min.toLocaleString()}</code>
• 平均淨額：<code>${stats.average.toLocaleString()}</code>
• 波動範圍：<code>${stats.range.toLocaleString()}</code>

<b>🎯 市場觀察：</b>
${trend}

<b>⏰ 生成時間：</b>
<code>${new Date().toLocaleString('zh-TW')}</code>

<i>${isRealData ? '✅ 真實期貨交易資料' : '🧪 模擬測試資料'}</i>`;
}

/**
 * 測試單一契約圖表生成
 */
function testSingleContractChart(contract = 'TX') {
  console.log(`🧪 開始測試${contract}圖表生成功能...`);
  
  // 生成模擬資料
  const mockData = generateMockNetVolumeData();
  
  console.log(`建立${contract}的模擬資料:`, JSON.stringify(mockData.slice(0, 3), null, 2));
  
  // 建立圖表
  const chartInfo = createSingleContractChart(mockData, contract);
  
  if (!chartInfo) {
    console.error(`❌ ${contract}圖表建立失敗`);
    return false;
  }
  
  console.log(`✅ ${contract}圖表建立成功`);
  
  // 準備說明訊息
  const caption = generateContractChartCaption(contract, chartInfo.stats, false, mockData);
  
  // 發送圖表
  const success = sendChartToTelegram(chartInfo, caption);
  
  if (success) {
    console.log(`🎉 ${contract}測試成功！圖表已發送到Telegram`);
    return true;
  } else {
    console.error(`❌ ${contract}圖表發送失敗`);
    return false;
  }
}

/**
 * 測試所有契約的圖表生成（使用模擬資料）
 */
function testAllContractCharts() {
  console.log('🧪 開始測試所有契約圖表生成功能...');
  
  const mockData = generateMockNetVolumeData();
  const contracts = ['TX', 'MXF', 'EXF', 'TXF', 'NQF'];
  
  let successCount = 0;
  
  for (const contract of contracts) {
    console.log(`\n--- 測試 ${contract} ---`);
    
    const chartInfo = createSingleContractChart(mockData, contract);
    
    if (chartInfo) {
      const caption = generateContractChartCaption(contract, chartInfo.stats, false, mockData);
      const success = sendChartToTelegram(chartInfo, caption);
      
      if (success) {
        successCount++;
        console.log(`✅ ${contract} 測試成功`);
      } else {
        console.error(`❌ ${contract} 測試失敗`);
      }
    } else {
      console.error(`❌ ${contract} 圖表建立失敗`);
    }
    
    // 間隔2秒
    if (contracts.indexOf(contract) < contracts.length - 1) {
      Utilities.sleep(2000);
    }
  }
  
  console.log(`\n🎉 測試完成！成功: ${successCount}/${contracts.length}`);
  return successCount === contracts.length;
}

/**
 * 建立多空淨額交易口數柱狀圖（合併版本，保留舊功能）
 */
function createNetVolumeChart(data, title = '近30個交易日多空淨額交易口數') {
  try {
    console.log('開始建立合併圖表...');
    
    if (!data || data.length === 0) {
      console.error('沒有資料可以建立圖表');
      return null;
    }
    
    // 建立臨時試算表來存放圖表
    const tempSS = SpreadsheetApp.create('臨時圖表_' + new Date().getTime());
    const sheet = tempSS.getActiveSheet();
    
    // 設定表頭
    const headers = ['日期', 'TX', 'MXF', 'EXF', 'TXF', 'NQF'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 填入資料
    const chartData = data.map(day => [
      day.date,
      day.TX || 0,
      day.MXF || 0,
      day.EXF || 0,
      day.TXF || 0,
      day.NQF || 0
    ]);
    
    sheet.getRange(2, 1, chartData.length, headers.length).setValues(chartData);
    
    // 建立圖表
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(1, 1, chartData.length + 1, headers.length))
      .setPosition(1, 8, 0, 0)
      .setOption('title', title)
      .setOption('titleTextStyle', {
        fontSize: 16,
        bold: true
      })
      .setOption('legend', {
        position: 'top',
        alignment: 'center'
      })
      .setOption('hAxis', {
        title: '日期',
        titleTextStyle: { fontSize: 12 },
        textStyle: { fontSize: 10 },
        slantedText: true,
        slantedTextAngle: 45
      })
      .setOption('vAxis', {
        title: '多空淨額交易口數',
        titleTextStyle: { fontSize: 12 }
      })
      .setOption('series', {
        0: { color: '#FF6B6B', name: 'TX' },      // 紅色
        1: { color: '#4ECDC4', name: 'MXF' },     // 青色
        2: { color: '#45B7D1', name: 'EXF' },     // 藍色  
        3: { color: '#96CEB4', name: 'TXF' },     // 綠色
        4: { color: '#FECA57', name: 'NQF' }      // 黃色
      })
      .setOption('width', 800)
      .setOption('height', 500)
      .setOption('backgroundColor', '#FFFFFF')
      .build();
    
    sheet.insertChart(chart);
    
    console.log('合併圖表建立完成');
    
    // 等待一下讓圖表完全載入
    Utilities.sleep(2000);
    
    return {
      spreadsheet: tempSS,
      sheet: sheet,
      chart: chart
    };
    
  } catch (error) {
    console.error('建立合併圖表時發生錯誤:', error.toString());
    return null;
  }
}

/**
 * 將圖表匯出為圖片並發送到Telegram（改進版）
 */
function sendChartToTelegramImproved(chartInfo, message = '') {
  try {
    console.log('🖼️ 開始改進版圖表匯出流程...');
    
    if (!chartInfo || !chartInfo.chart) {
      console.error('❌ 圖表資訊無效');
      return false;
    }
    
    console.log('✅ 圖表資訊驗證通過，開始取得圖片...');
    
    // 再次等待確保圖表完全渲染
    console.log('⏳ 等待圖表完全渲染...');
    SpreadsheetApp.flush();
    Utilities.sleep(5000);  // 增加到5秒
    
    // 嘗試多次獲取圖表blob
    let chartBlob = null;
    let attempts = 0;
    const maxAttempts = 5;
    
    while (attempts < maxAttempts && !chartBlob) {
      attempts++;
      console.log(`🔄 第${attempts}/${maxAttempts}次嘗試獲取圖表圖片...`);
      
      try {
        // 重新獲取圖表（以防第一次獲取的有問題）
        const charts = chartInfo.sheet.getCharts();
        if (charts.length === 0) {
          console.error('❌ 找不到圖表');
          break;
        }
        
        const currentChart = charts[0];
        chartBlob = currentChart.getAs('image/png');
        
        if (chartBlob && chartBlob.getBytes().length > 5000) {  // 提高最小檔案大小要求
          console.log(`✅ 圖表圖片獲取成功！大小: ${chartBlob.getBytes().length} bytes`);
          break;
        } else {
          const size = chartBlob ? chartBlob.getBytes().length : 0;
          console.log(`⚠️ 第${attempts}次獲取的圖片太小(${size} bytes)或為空，重試...`);
          chartBlob = null;
          
          if (attempts < maxAttempts) {
            Utilities.sleep(3000);  // 每次重試間隔3秒
          }
        }
      } catch (e) {
        console.error(`❌ 第${attempts}次獲取圖片時發生錯誤:`, e.toString());
        if (attempts < maxAttempts) {
          Utilities.sleep(3000);
        }
      }
    }
    
    if (!chartBlob) {
      console.error('❌ 無法獲取有效的圖表圖片，已嘗試', maxAttempts, '次');
      
      // 提供試算表URL供手動檢查
      const spreadsheetUrl = chartInfo.spreadsheet.getUrl();
      console.error('📋 請手動檢查試算表:', spreadsheetUrl);
      
      // 發送錯誤通知
      const errorMessage = `❌ 圖表生成失敗\n\n契約: ${chartInfo.contract}\n錯誤: 無法獲取圖片\n\n請檢查試算表: ${spreadsheetUrl}`;
      sendTelegramMessage(errorMessage);
      
      return false;
    }
    
    // 發送圖片到Telegram
    console.log('📤 準備發送圖片到Telegram...');
    const success = sendTelegramPhoto(chartBlob, message);
    
    if (success) {
      console.log('🎉 圖表發送成功！');
    } else {
      console.error('❌ Telegram發送失敗');
    }
    
    // 清理臨時試算表
    try {
      console.log('🗑️ 清理臨時試算表...');
      Utilities.sleep(1000);  // 等待一下再清理
      DriveApp.getFileById(chartInfo.spreadsheet.getId()).setTrashed(true);
      console.log('✅ 臨時試算表已清理');
    } catch (cleanupError) {
      console.log('⚠️ 清理臨時試算表時發生錯誤（不影響主要功能）:', cleanupError.toString());
    }
    
    return success;
    
  } catch (error) {
    console.error('❌ 發送圖表到Telegram時發生錯誤:', error.toString());
    console.error('🔍 錯誤堆疊:', error.stack);
    return false;
  }
}

/**
 * 除錯用：手動檢查圖表建立過程
 */
function debugChartCreation(contract = 'TX') {
  console.log(`🔍 開始除錯${contract}圖表建立過程...`);
  
  try {
    // 生成測試資料
    const mockData = generateMockNetVolumeData();
    console.log('✅ 測試資料生成成功，筆數:', mockData.length);
    
    // 顯示部分測試資料
    console.log('📊 測試資料範例:', JSON.stringify(mockData.slice(0, 5), null, 2));
    
    // 建立圖表但不發送
    const chartInfo = createSingleContractChart(mockData, contract);
    
    if (!chartInfo) {
      console.error('❌ 圖表建立失敗');
      return false;
    }
    
    console.log('✅ 圖表資訊:', {
      spreadsheetId: chartInfo.spreadsheet.getId(),
      sheetName: chartInfo.sheet.getName(),
      chartExists: !!chartInfo.chart,
      stats: chartInfo.stats
    });
    
    // 檢查試算表內容
    const sheet = chartInfo.sheet;
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    console.log('📋 試算表資料檢查:');
    console.log('- 總行數:', values.length);
    console.log('- 總列數:', values[0] ? values[0].length : 0);
    console.log('- 表頭:', values[0]);
    console.log('- 第一筆資料:', values[1]);
    console.log('- 最後一筆資料:', values[values.length - 1]);
    
    // 檢查圖表
    const charts = sheet.getCharts();
    console.log('📈 圖表檢查:');
    console.log('- 圖表數量:', charts.length);
    
    if (charts.length > 0) {
      const chart = charts[0];
      console.log('- 圖表類型:', chart.getChartType());
      console.log('- 圖表選項存在:', !!chart.getOptions());
      
      // 檢查圖表範圍
      try {
        const ranges = chart.getRanges();
        console.log('- 圖表資料範圍數量:', ranges.length);
        if (ranges.length > 0) {
          console.log('- 第一個範圍:', ranges[0].getA1Notation());
        }
      } catch (rangeError) {
        console.log('- 無法獲取圖表範圍:', rangeError.toString());
      }
    }
    
    // 嘗試獲取圖片
    console.log('🖼️ 嘗試獲取圖片...');
    
    let imageSuccess = false;
    let imageSize = 0;
    
    try {
      // 等待足夠時間讓圖表渲染
      console.log('⏳ 等待圖表渲染 (5秒)...');
      Utilities.sleep(5000);
      
      const blob = chartInfo.chart.getAs('image/png');
      if (blob) {
        imageSize = blob.getBytes().length;
        console.log('✅ 圖片獲取成功！大小:', imageSize, 'bytes');
        
        if (imageSize < 1000) {
          console.log('⚠️ 圖片檔案很小，可能是空白的');
        } else if (imageSize < 5000) {
          console.log('⚠️ 圖片檔案偏小，內容可能不完整');
        } else {
          console.log('✅ 圖片大小正常，應該包含完整內容');
          imageSuccess = true;
        }
      } else {
        console.log('❌ 圖片blob為空');
      }
    } catch (imgError) {
      console.error('❌ 獲取圖片時發生錯誤:', imgError.toString());
    }
    
    // 提供試算表URL讓用戶手動檢查
    const spreadsheetUrl = chartInfo.spreadsheet.getUrl();
    console.log('🔗 試算表URL (手動檢查用):', spreadsheetUrl);
    
    // 總結
    console.log('\n📋 除錯總結:');
    console.log('✅ 資料生成: 成功');
    console.log('✅ 試算表建立: 成功');
    console.log('✅ 圖表建立:', charts.length > 0 ? '成功' : '失敗');
    console.log('✅ 圖片獲取:', imageSuccess ? '成功' : '失敗');
    console.log('📊 圖片大小:', imageSize, 'bytes');
    
    // 發送測試通知
    const testMessage = `🔍 ${contract} 圖表除錯結果

📊 圖表建立: ${charts.length > 0 ? '✅ 成功' : '❌ 失敗'}
🖼️ 圖片獲取: ${imageSuccess ? '✅ 成功' : '❌ 失敗'}
📏 圖片大小: ${imageSize} bytes

🔗 試算表連結:
${spreadsheetUrl}

💡 請手動檢查試算表中的圖表是否正常顯示`;

    sendTelegramMessage(testMessage);
    
    console.log('🔍 除錯完成！結果已發送到Telegram');
    
    return {
      success: true,
      chartInfo: chartInfo,
      spreadsheetUrl: spreadsheetUrl,
      imageSuccess: imageSuccess,
      imageSize: imageSize
    };
    
  } catch (error) {
    console.error('❌ 除錯過程發生錯誤:', error.toString());
    console.error('🔍 錯誤堆疊:', error.stack);
    return false;
  }
}

/**
 * 測試改進後的圖表功能
 */
function testImprovedChart(contract = 'TX') {
  console.log(`🚀 測試改進後的${contract}圖表功能...`);
  
  const mockData = generateMockNetVolumeData();
  console.log('📊 資料準備完成，開始建立圖表...');
  
  const chartInfo = createSingleContractChart(mockData, contract);
  
  if (!chartInfo) {
    console.error(`❌ ${contract}圖表建立失敗`);
    return false;
  }
  
  console.log(`✅ ${contract}圖表建立成功，準備發送...`);
  
  const caption = generateContractChartCaption(contract, chartInfo.stats, false, mockData);
  const success = sendChartToTelegramImproved(chartInfo, caption);  // 使用改進版本
  
  if (success) {
    console.log(`🎉 ${contract}改進版測試成功！`);
    return true;
  } else {
    console.error(`❌ ${contract}圖表發送失敗`);
    return false;
  }
}

/**
 * 簡化測試：只建立圖表但不發送，檢查內容
 */
function quickChartTest(contract = 'TX') {
  console.log(`⚡ 快速測試${contract}圖表建立...`);
  
  try {
    const mockData = generateMockNetVolumeData();
    const chartInfo = createSingleContractChart(mockData, contract);
    
    if (chartInfo) {
      const url = chartInfo.spreadsheet.getUrl();
      console.log('✅ 圖表建立成功！');
      console.log('🔗 檢查連結:', url);
      
      // 發送連結到Telegram
      const message = `⚡ ${contract} 快速測試結果

✅ 圖表建立成功！

🔗 請點擊連結檢查圖表:
${url}

💡 如果圖表顯示正常，則說明問題可能在圖片匯出環節`;

      sendTelegramMessage(message);
      return true;
    } else {
      console.error('❌ 圖表建立失敗');
      return false;
    }
  } catch (error) {
    console.error('❌ 測試失敗:', error.toString());
    return false;
  }
}

/**
 * 測試圖表生成和發送功能（合併版本）
 */
function testChartGeneration() {
  console.log('🧪 開始測試圖表生成功能...');
  
  // 建立模擬資料
  const mockData = generateMockNetVolumeData();
  
  console.log('建立的模擬資料:', JSON.stringify(mockData.slice(0, 3), null, 2));
  
  // 建立圖表
  const chartInfo = createNetVolumeChart(mockData, '測試：近30個交易日多空淨額交易口數');
  
  if (!chartInfo) {
    console.error('❌ 圖表建立失敗');
    return false;
  }
  
  console.log('✅ 圖表建立成功');
  
  // 準備說明訊息
  const caption = `<b>📊 期貨多空淨額交易口數圖表</b>

<b>📅 資料期間：</b>近30個交易日
<b>📈 圖表說明：</b>
• 紅色：TX (台指期)
• 青色：MXF (小台指)  
• 藍色：EXF (電子期)
• 綠色：TXF (台指期貨)
• 黃色：NQF (那斯達克期)

<b>⏰ 生成時間：</b>
<code>${new Date().toLocaleString('zh-TW')}</code>

<i>✅ 這是測試圖表，使用模擬資料生成</i>`;

  // 發送圖表
  const success = sendChartToTelegram(chartInfo, caption);
  
  if (success) {
    console.log('🎉 測試成功！圖表已發送到Telegram');
    return true;
  } else {
    console.error('❌ 圖表發送失敗');
    return false;
  }
}

/**
 * 使用真實資料測試圖表功能
 */
function testRealDataChart() {
  console.log('🧪 開始測試真實資料圖表功能...');
  
  // 讀取真實資料
  const realData = getLast30DaysNetVolumeData();
  
  if (!realData || realData.length === 0) {
    console.error('❌ 無法獲取真實資料，改用模擬資料測試');
    return testChartGeneration();
  }
  
  console.log(`✅ 成功讀取${realData.length}天的真實資料`);
  
  // 建立圖表
  const chartInfo = createNetVolumeChart(realData, '真實資料：近30個交易日多空淨額交易口數');
  
  if (!chartInfo) {
    console.error('❌ 圖表建立失敗');
    return false;
  }
  
  // 計算統計資料
  const stats = calculateNetVolumeStats(realData);
  
  // 準備說明訊息
  const caption = `<b>📊 期貨多空淨額交易口數圖表</b>

<b>📅 資料期間：</b>${stats.dateRange}
<b>📈 統計摘要：</b>
• TX 平均淨額：<code>${stats.averages.TX}</code>
• MXF 平均淨額：<code>${stats.averages.MXF}</code>
• EXF 平均淨額：<code>${stats.averages.EXF}</code>
• TXF 平均淨額：<code>${stats.averages.TXF}</code>  
• NQF 平均淨額：<code>${stats.averages.NQF}</code>

<b>⏰ 生成時間：</b>
<code>${new Date().toLocaleString('zh-TW')}</code>

<i>✅ 使用真實期貨交易資料</i>`;

  // 發送圖表
  const success = sendChartToTelegram(chartInfo, caption);
  
  if (success) {
    console.log('🎉 真實資料測試成功！圖表已發送到Telegram');
    return true;
  } else {
    console.error('❌ 圖表發送失敗');
    return false;
  }
}

/**
 * 計算淨額統計資料
 */
function calculateNetVolumeStats(data) {
  const contracts = ['TX', 'MXF', 'EXF', 'TXF', 'NQF'];
  const stats = {
    dateRange: `${data[0]?.date} - ${data[data.length - 1]?.date}`,
    averages: {}
  };
  
  contracts.forEach(contract => {
    const values = data.map(day => day[contract] || 0);
    const average = values.reduce((sum, val) => sum + val, 0) / values.length;
    stats.averages[contract] = Math.round(average).toLocaleString();
  });
  
  return stats;
}

/**
 * 發送圖片到Telegram
 */
function sendTelegramPhoto(imageBlob, caption = '') {
  try {
    const botToken = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
    const chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
    
    if (!botToken || !chatId) {
      console.error('❌ Telegram配置不完整');
      return false;
    }
    
    const url = `https://api.telegram.org/bot${botToken}/sendPhoto`;
    
    // 準備multipart/form-data
    const payload = {
      'chat_id': chatId,
      'photo': imageBlob,
      'caption': caption,
      'parse_mode': 'HTML'
    };
    
    console.log('📤 發送圖片到Telegram...');
    
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      payload: payload,
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('HTTP狀態碼:', responseCode);
    
    if (responseCode === 200) {
      const data = JSON.parse(responseText);
      if (data.ok) {
        console.log('✅ 圖片發送成功！');
        return true;
      } else {
        console.error('❌ API錯誤:', data.description);
        console.error('錯誤代碼:', data.error_code);
        return false;
      }
    } else {
      console.error('❌ HTTP錯誤:', responseCode);
      console.error('回應內容:', responseText);
      return false;
    }
    
  } catch (error) {
    console.error('❌ 發送圖片時發生錯誤:', error.toString());
    return false;
  }
}

/**
 * 將圖表匯出為圖片並發送到Telegram（原版本，保留兼容性）
 */
function sendChartToTelegram(chartInfo, message = '') {
  try {
    console.log('開始匯出圖表為圖片...');
    
    if (!chartInfo || !chartInfo.chart) {
      console.error('圖表資訊無效');
      return false;
    }
    
    // 獲取圖表的blob
    const chartBlob = chartInfo.chart.getAs('image/png');
    
    if (!chartBlob) {
      console.error('無法獲取圖表圖片');
      return false;
    }
    
    console.log('圖表圖片獲取成功，大小:', chartBlob.getBytes().length, 'bytes');
    
    // 發送圖片到Telegram
    const success = sendTelegramPhoto(chartBlob, message);
    
    // 清理臨時試算表
    try {
      DriveApp.getFileById(chartInfo.spreadsheet.getId()).setTrashed(true);
      console.log('臨時試算表已清理');
    } catch (cleanupError) {
      console.log('清理臨時試算表時發生錯誤（這不影響主要功能）:', cleanupError.toString());
    }
    
    return success;
    
  } catch (error) {
    console.error('發送圖表到Telegram時發生錯誤:', error.toString());
    return false;
  }
}

/**
 * 生成模擬的30天資料用於測試
 */
function generateMockNetVolumeData() {
  const data = [];
  const today = new Date();
  
  for (let i = 29; i >= 0; i--) {
    const date = new Date(today);
    date.setDate(date.getDate() - i);
    
    // 跳過週末
    if (date.getDay() === 0 || date.getDay() === 6) {
      continue;
    }
    
    const dateStr = date.toLocaleDateString('zh-TW').replace(/\//g, '-');
    
    data.push({
      date: dateStr,
      TX: Math.floor(Math.random() * 20000 - 10000),      // -10000 到 +10000
      MXF: Math.floor(Math.random() * 15000 - 7500),      // -7500 到 +7500
      EXF: Math.floor(Math.random() * 8000 - 4000),       // -4000 到 +4000
      TXF: Math.floor(Math.random() * 12000 - 6000),      // -6000 到 +6000
      NQF: Math.floor(Math.random() * 5000 - 2500)        // -2500 到 +2500
    });
  }
  
  return data;
}

/**
 * 檢查資料格式和建立簡化圖表測試
 */
function testDataAndSimpleChart(contract = 'TX') {
  console.log(`🔍 檢查${contract}的資料格式和建立簡化圖表...`);
  
  try {
    // 生成並檢查測試資料
    const mockData = generateMockNetVolumeData();
    console.log('原始資料範例:', JSON.stringify(mockData.slice(0, 3), null, 2));
    
    // 提取單一契約資料
    const contractData = mockData.map(day => ({
      date: day.date,
      value: day[contract] || 0
    }));
    
    console.log(`${contract}提取後資料:`, JSON.stringify(contractData.slice(0, 3), null, 2));
    
    // 建立最簡單的試算表和圖表
    const tempSS = SpreadsheetApp.create(`簡化測試_${contract}_${new Date().getTime()}`);
    const sheet = tempSS.getActiveSheet();
    
    // 設定表頭
    sheet.getRange('A1').setValue('日期');
    sheet.getRange('B1').setValue(contract);
    
    // 逐行填入資料
    contractData.forEach((row, index) => {
      const rowNum = index + 2;
      sheet.getRange(`A${rowNum}`).setValue(row.date);
      sheet.getRange(`B${rowNum}`).setValue(row.value);
    });
    
    console.log(`已填入${contractData.length}筆資料到試算表`);
    
    // 檢查實際寫入的資料
    const writtenData = sheet.getRange(1, 1, contractData.length + 1, 2).getValues();
    console.log('實際寫入的資料前3筆:', writtenData.slice(0, 3));
    
    // 建立最簡單的圖表配置
    const dataRange = sheet.getRange(1, 1, contractData.length + 1, 2);
    console.log('圖表資料範圍:', dataRange.getA1Notation());
    
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataRange)
      .setPosition(1, 4, 0, 0)
      .setOption('title', `${contract} 簡化測試圖表`)
      .setOption('width', 600)
      .setOption('height', 400)
      .build();
    
    sheet.insertChart(chart);
    
    console.log('圖表建立完成，等待渲染...');
    SpreadsheetApp.flush();
    Utilities.sleep(3000);
    
    // 檢查結果
    const charts = sheet.getCharts();
    const spreadsheetUrl = tempSS.getUrl();
    
    console.log('圖表數量:', charts.length);
    console.log('試算表URL:', spreadsheetUrl);
    
    // 發送結果
    const resultMessage = `🔍 ${contract} 簡化測試結果

📊 資料筆數: ${contractData.length}
📈 圖表建立: ${charts.length > 0 ? '✅ 成功' : '❌ 失敗'}

🔗 檢查連結:
${spreadsheetUrl}

📋 資料範例:
${JSON.stringify(contractData.slice(0, 3), null, 2)}

💡 請檢查試算表中的圖表是否正常顯示`;

    sendTelegramMessage(resultMessage);
    
    return {
      success: charts.length > 0,
      spreadsheetUrl: spreadsheetUrl,
      dataCount: contractData.length
    };
    
  } catch (error) {
    console.error('❌ 簡化測試失敗:', error.toString());
    sendTelegramMessage(`❌ ${contract} 簡化測試失敗: ${error.toString()}`);
    return false;
  }
}

/**
 * 比較成功和失敗的圖表配置
 */
function compareChartConfigs(contract = 'TX') {
  console.log(`🔄 比較${contract}的成功與失敗配置...`);
  
  try {
    const mockData = generateMockNetVolumeData();
    
    // 建立試算表
    const tempSS = SpreadsheetApp.create(`比較測試_${contract}_${new Date().getTime()}`);
    const sheet = tempSS.getActiveSheet();
    
    // 1. 先建立成功的合併圖表格式
    console.log('建立成功的合併格式資料...');
    
    // 合併格式：多列資料
    const headers = ['日期', 'TX', 'MXF', 'EXF', 'TXF', 'NQF'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    const chartData = mockData.map(day => [
      day.date,
      day.TX || 0,
      day.MXF || 0,
      day.EXF || 0,
      day.TXF || 0,
      day.NQF || 0
    ]);
    
    sheet.getRange(2, 1, chartData.length, headers.length).setValues(chartData);
    
    // 建立成功的合併圖表（只顯示指定契約）
    const successChart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(1, 1, chartData.length + 1, 2))  // 只取日期和第一個契約
      .setPosition(1, 8, 0, 0)
      .setOption('title', `${contract} 使用合併格式`)
      .setOption('legend', { position: 'none' })
      .setOption('width', 600)
      .setOption('height', 400)
      .build();
    
    sheet.insertChart(successChart);
    
    // 2. 在同一個試算表的不同位置建立單一契約格式
    console.log('建立單一契約格式資料...');
    
    const startRow = chartData.length + 5;  // 在下方建立
    
    // 單一格式：兩列資料
    sheet.getRange(startRow, 1).setValue('日期');
    sheet.getRange(startRow, 2).setValue(contract);
    
    const singleData = mockData.map(day => [day.date, day[contract] || 0]);
    sheet.getRange(startRow + 1, 1, singleData.length, 2).setValues(singleData);
    
    // 建立單一契約圖表
    const singleChart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(startRow, 1, singleData.length + 1, 2))
      .setPosition(startRow, 4, 0, 0)
      .setOption('title', `${contract} 使用單一格式`)
      .setOption('legend', { position: 'none' })
      .setOption('width', 600)
      .setOption('height', 400)
      .build();
    
    sheet.insertChart(singleChart);
    
    SpreadsheetApp.flush();
    Utilities.sleep(3000);
    
    const charts = sheet.getCharts();
    const spreadsheetUrl = tempSS.getUrl();
    
    console.log('總圖表數:', charts.length);
    
    const message = `🔄 ${contract} 配置比較測試

📊 建立結果:
• 合併格式圖表: ${charts.length >= 1 ? '✅' : '❌'}
• 單一格式圖表: ${charts.length >= 2 ? '✅' : '❌'}
• 總圖表數: ${charts.length}

🔗 檢查連結:
${spreadsheetUrl}

💡 請比較兩個圖表的顯示效果，看哪種格式正常`;

    sendTelegramMessage(message);
    
    return {
      success: charts.length >= 2,
      spreadsheetUrl: spreadsheetUrl,
      chartCount: charts.length
    };
    
  } catch (error) {
    console.error('❌ 比較測試失敗:', error.toString());
    return false;
  }
}

/**
 * 使用成功格式建立單一契約圖表
 */
function createSingleChartWithSuccessFormat(data, contract, title = '') {
  try {
    console.log(`📊 使用成功格式建立${contract}圖表...`);
    
    if (!data || data.length === 0) {
      console.error('沒有資料可以建立圖表');
      return null;
    }
    
    // 建立臨時試算表
    const tempSS = SpreadsheetApp.create(`成功格式_${contract}_${new Date().getTime()}`);
    const sheet = tempSS.getActiveSheet();
    
    // 使用成功的合併圖表格式，但只填入兩列
    const headers = ['日期', contract];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 填入資料，確保格式與成功版本一致
    const chartData = data.map(day => [
      day.date,
      day[contract] || 0
    ]);
    
    sheet.getRange(2, 1, chartData.length, headers.length).setValues(chartData);
    
    // 計算統計
    const values = chartData.map(row => row[1]);
    const maxVal = Math.max(...values);
    const minVal = Math.min(...values);
    const avgVal = values.reduce((sum, val) => sum + val, 0) / values.length;
    
    const contractColors = {
      'TX': '#FF6B6B',
      'MXF': '#4ECDC4',
      'EXF': '#45B7D1',
      'TXF': '#96CEB4',
      'NQF': '#FECA57'
    };
    
    const contractNames = {
      'TX': '台指期',
      'MXF': '小台指',
      'EXF': '電子期',
      'TXF': '台指期貨',
      'NQF': '那斯達克期'
    };
    
    const chartTitle = title || `${contractNames[contract] || contract} 近30個交易日多空淨額`;
    
    // 完全按照成功的合併圖表配置
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(1, 1, chartData.length + 1, headers.length))
      .setPosition(1, 4, 0, 0)
      .setOption('title', chartTitle)
      .setOption('titleTextStyle', {
        fontSize: 16,
        bold: true
      })
      .setOption('legend', {
        position: 'none'
      })
      .setOption('hAxis', {
        title: '日期',
        titleTextStyle: { fontSize: 12 },
        textStyle: { fontSize: 10 },
        slantedText: true,
        slantedTextAngle: 45
      })
      .setOption('vAxis', {
        title: '多空淨額交易口數',
        titleTextStyle: { fontSize: 12 }
      })
      .setOption('series', {
        0: { color: contractColors[contract] || '#4285F4' }
      })
      .setOption('width', 800)
      .setOption('height', 500)
      .setOption('backgroundColor', '#FFFFFF')
      .build();
    
    sheet.insertChart(chart);
    
    SpreadsheetApp.flush();
    Utilities.sleep(3000);
    
    const charts = sheet.getCharts();
    if (charts.length === 0) {
      console.error('❌ 圖表建立失敗');
      return null;
    }
    
    console.log(`✅ ${contract}圖表建立成功`);
    
    return {
      spreadsheet: tempSS,
      sheet: sheet,
      chart: charts[0],
      contract: contract,
      stats: {
        max: maxVal,
        min: minVal,
        average: Math.round(avgVal),
        range: maxVal - minVal
      }
    };
    
  } catch (error) {
    console.error(`❌ 建立${contract}圖表失敗:`, error.toString());
    return null;
  }
}

/**
 * 測試使用成功格式的單一契約圖表
 */
function testSuccessFormatChart(contract = 'TX') {
  console.log(`🎯 測試${contract}成功格式圖表...`);
  
  const mockData = generateMockNetVolumeData();
  const chartInfo = createSingleChartWithSuccessFormat(mockData, contract);
  
  if (!chartInfo) {
    console.error(`❌ ${contract}圖表建立失敗`);
    return false;
  }
  
  console.log(`✅ ${contract}圖表建立成功，準備發送...`);
  
  const caption = generateContractChartCaption(contract, chartInfo.stats, false, mockData);
  const success = sendChartToTelegramImproved(chartInfo, caption);
  
  if (success) {
    console.log(`🎉 ${contract}成功格式測試完成！`);
    return true;
  } else {
    console.error(`❌ ${contract}圖表發送失敗`);
    return false;
  }
}

/**
 * 使用真實期貨資料測試單一契約圖表
 */
function testRealDataSingleChart(contract = 'TX') {
  console.log(`📈 使用真實資料測試${contract}圖表...`);
  
  try {
    // 讀取真實的期貨資料
    const realData = getLast30DaysNetVolumeData();
    
    if (!realData || realData.length === 0) {
      console.error('❌ 無法獲取真實資料，請檢查試算表');
      
      // 發送錯誤通知
      const errorMessage = `❌ 無法讀取真實期貨資料

可能原因：
• 試算表名稱不正確
• 資料格式不符
• 權限問題

請檢查：
1. 試算表是否命名為「所有期貨資料」
2. 資料欄位是否正確（日期、契約代碼、身份別、多空淨額等）
3. 是否有近期的資料

💡 建議先執行 debugRealDataReading() 診斷問題`;

      sendTelegramMessage(errorMessage);
      return false;
    }
    
    console.log(`✅ 成功讀取${realData.length}天的真實資料`);
    console.log('真實資料範例:', JSON.stringify(realData.slice(0, 2), null, 2));
    
    // 檢查指定契約是否有資料
    const contractHasData = realData.some(day => day[contract] && day[contract] !== 0);
    if (!contractHasData) {
      console.warn(`⚠️ ${contract}契約在真實資料中沒有數值或全為0`);
    }
    
    // 建立圖表
    const chartInfo = createSingleContractChart(realData, contract);
    
    if (!chartInfo) {
      console.error(`❌ ${contract}真實資料圖表建立失敗`);
      return false;
    }
    
    console.log(`✅ ${contract}真實資料圖表建立成功`);
    
    // 準備說明訊息
    const caption = generateContractChartCaption(contract, chartInfo.stats, true, realData);
    
    // 發送圖表
    const success = sendChartToTelegramImproved(chartInfo, caption);
    
    if (success) {
      console.log(`🎉 ${contract}真實資料圖表發送成功！`);
      return true;
    } else {
      console.error(`❌ ${contract}真實資料圖表發送失敗`);
      return false;
    }
    
  } catch (error) {
    console.error('❌ 真實資料測試失敗:', error.toString());
    
    const errorMessage = `❌ ${contract} 真實資料測試失敗

錯誤訊息: ${error.toString()}

💡 建議：
1. 檢查試算表是否存在且有資料
2. 執行 debugRealDataReading() 診斷問題
3. 確認資料格式是否正確`;

    sendTelegramMessage(errorMessage);
    return false;
  }
}

/**
 * 使用真實資料生成所有契約圖表
 */
function sendAllRealDataCharts() {
  console.log('🚀 使用真實資料生成所有契約圖表...');
  
  try {
    // 讀取真實資料
    const realData = getLast30DaysNetVolumeData();
    
    if (!realData || realData.length === 0) {
      console.error('❌ 無法獲取真實資料');
      sendTelegramMessage('❌ 無法讀取真實期貨資料，請檢查試算表設定');
      return [];
    }
    
    console.log(`✅ 讀取到${realData.length}天的真實資料`);
    
    const contracts = ['TX', 'MXF', 'EXF', 'TXF', 'NQF'];
    const contractNames = {
      'TX': '台指期',
      'MXF': '小台指', 
      'EXF': '電子期',
      'TXF': '台指期貨',
      'NQF': '那斯達克期'
    };
    
    let successCount = 0;
    let results = [];
    
    // 為每個契約生成圖表
    for (const contract of contracts) {
      console.log(`\n--- 處理 ${contract} (${contractNames[contract]}) 真實資料 ---`);
      
      // 檢查該契約是否有資料
      const contractHasData = realData.some(day => day[contract] && day[contract] !== 0);
      if (!contractHasData) {
        console.warn(`⚠️ ${contract}在真實資料中沒有數值，跳過`);
        results.push({
          contract: contract,
          success: false,
          reason: '無資料'
        });
        continue;
      }
      
      // 建立圖表
      const chartInfo = createSingleContractChart(realData, contract);
      
      if (!chartInfo) {
        console.error(`❌ ${contract} 圖表建立失敗`);
        results.push({
          contract: contract,
          success: false,
          reason: '圖表建立失敗'
        });
        continue;
      }
      
      // 準備說明訊息
      const caption = generateContractChartCaption(contract, chartInfo.stats, true, realData);
      
      // 發送圖表
      const success = sendChartToTelegramImproved(chartInfo, caption);
      
      if (success) {
        console.log(`✅ ${contract} 真實資料圖表發送成功`);
        successCount++;
        results.push({
          contract: contract,
          success: true
        });
      } else {
        console.error(`❌ ${contract} 真實資料圖表發送失敗`);
        results.push({
          contract: contract,
          success: false,
          reason: '發送失敗'
        });
      }
      
      // 間隔2秒避免API限制
      if (contracts.indexOf(contract) < contracts.length - 1) {
        console.log('等待2秒...');
        Utilities.sleep(2000);
      }
    }
    
    // 發送總結訊息
    const dateRange = realData.length > 0 ? 
      `${realData[0]?.date} - ${realData[realData.length - 1]?.date}` : 
      '無資料';
    
    const summaryMessage = `📊 真實期貨資料圖表生成完成

📅 資料期間: ${dateRange}
📈 資料天數: ${realData.length} 天

📊 生成結果:
${results.map(r => {
  const status = r.success ? '✅' : `❌ (${r.reason || '失敗'})`;
  return `• ${contractNames[r.contract]}: ${status}`;
}).join('\n')}

📈 統計:
• 成功: ${successCount} / ${contracts.length}
• 資料來源: 真實期貨交易資料

⏰ 完成時間: ${new Date().toLocaleString('zh-TW')}

🎯 已完成所有契約的真實資料圖表生成`;
    
    sendTelegramMessage(summaryMessage, 'HTML');
    
    console.log(`\n🎉 完成！真實資料圖表成功生成 ${successCount}/${contracts.length} 個`);
    return results;
    
  } catch (error) {
    console.error('❌ 真實資料圖表生成失敗:', error.toString());
    
    const errorMessage = `❌ 真實資料圖表生成失敗

錯誤: ${error.toString()}

💡 建議檢查：
1. 試算表是否存在且命名為「所有期貨資料」
2. 資料格式是否正確
3. 是否有近期的交易資料`;

    sendTelegramMessage(errorMessage);
    return [];
  }
}

/**
 * 除錯真實資料讀取
 */
function debugRealDataReading() {
  console.log('🔍 開始除錯真實資料讀取...');
  
  try {
    // 檢查試算表
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log('✅ 成功獲取活動試算表:', ss.getName());
    
    // 檢查工作表
    const sheets = ss.getSheets();
    console.log('📋 可用工作表:', sheets.map(s => s.getName()));
    
    // 嘗試找到期貨資料工作表
    let sheet = ss.getSheetByName('所有期貨資料');
    if (!sheet) {
      sheet = ss.getSheets()[0];
      console.log('⚠️ 找不到「所有期貨資料」工作表，使用第一個工作表:', sheet.getName());
    } else {
      console.log('✅ 找到所有期貨資料工作表');
    }
    
    // 檢查資料
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    console.log('📊 工作表資訊:');
    console.log('- 最後一行:', lastRow);
    console.log('- 最後一列:', lastCol);
    
    if (lastRow <= 1) {
      console.error('❌ 工作表沒有資料（只有標題行或空白）');
      sendTelegramMessage('❌ 所有期貨資料工作表是空的，請先匯入資料');
      return false;
    }
    
    // 檢查表頭
    const headers = sheet.getRange(1, 1, 1, Math.min(lastCol, 15)).getValues()[0];
    console.log('📋 表頭:', headers);
    
    // 檢查前幾筆資料
    const sampleData = sheet.getRange(2, 1, Math.min(5, lastRow - 1), Math.min(lastCol, 15)).getValues();
    console.log('📊 資料範例:', sampleData);
    
    // 檢查日期欄位
    const dateColumn = sampleData.map(row => row[0]);
    console.log('📅 日期欄位範例:', dateColumn);
    
    // 檢查契約代碼
    const contractColumn = sampleData.map(row => row[1]);
    console.log('📋 契約代碼範例:', contractColumn);
    
    // 檢查多空淨額欄位（假設在第8欄）
    if (lastCol >= 8) {
      const netVolumeColumn = sampleData.map(row => row[7]);
      console.log('📈 多空淨額範例:', netVolumeColumn);
    }
    
    // 發送除錯結果
    const debugMessage = `🔍 真實資料除錯結果

📊 試算表: ${ss.getName()}
📋 工作表: ${sheet.getName()}
📏 資料規模: ${lastRow - 1} 筆資料, ${lastCol} 欄位

📋 表頭:
${headers.slice(0, 8).map((h, i) => `${i + 1}. ${h}`).join('\n')}

📊 資料範例:
${sampleData.slice(0, 3).map((row, i) => 
  `第${i + 1}筆: ${row.slice(0, 5).join(' | ')}`
).join('\n')}

💡 請確認：
1. 第1欄是日期
2. 第2欄是契約代碼（TX, MXF等）
3. 第8欄是多空淨額交易口數`;

    sendTelegramMessage(debugMessage);
    
    return true;
    
  } catch (error) {
    console.error('❌ 除錯失敗:', error.toString());
    sendTelegramMessage(`❌ 資料除錯失敗: ${error.toString()}`);
    return false;
  }
}

/**
 * 快速測試真實資料是否可讀取
 */
function quickTestRealData() {
  console.log('⚡ 快速測試真實資料讀取...');
  
  try {
    const realData = getLast30DaysNetVolumeData();
    
    if (realData && realData.length > 0) {
      const message = `⚡ 真實資料測試結果

✅ 成功讀取 ${realData.length} 天的資料

📊 資料範例:
${JSON.stringify(realData.slice(0, 2), null, 2)}

📈 各契約資料概覽:
• TX: ${realData.filter(d => d.TX && d.TX !== 0).length} 天有資料
• MXF: ${realData.filter(d => d.MXF && d.MXF !== 0).length} 天有資料
• EXF: ${realData.filter(d => d.EXF && d.EXF !== 0).length} 天有資料
• TXF: ${realData.filter(d => d.TXF && d.TXF !== 0).length} 天有資料
• NQF: ${realData.filter(d => d.NQF && d.NQF !== 0).length} 天有資料

🎯 可以開始使用真實資料生成圖表了！`;

      sendTelegramMessage(message);
      return true;
    } else {
      sendTelegramMessage('❌ 無法讀取真實資料，請執行 debugRealDataReading() 診斷問題');
      return false;
    }
  } catch (error) {
    sendTelegramMessage(`❌ 真實資料測試失敗: ${error.toString()}`);
    return false;
  }
}