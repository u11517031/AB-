/**
 * 數位科技概論｜版本四：幾 A 幾 B 個人存檔後端 (Code.gs)
 * 注意：本版本不用執行 setup，採 USER_ACCESSING 存取者身分模式
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('幾 A 幾 B 數字猜測 (Google 個人存檔版)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 取得目前登入使用者的 Email 資訊
 */
function getCurrentUserInfo() {
  var email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return {
    email: email || '已登入使用者',
    hasSheet: hasUserSheet_()
  };
}

/**
 * 檢查目前使用者是否已經建立過存檔試算表
 */
function hasUserSheet_() {
  var userProperties = PropertiesService.getUserProperties();
  var sheetId = userProperties.getProperty('GUESS_V4_BOOK_ID');
  if (!sheetId) return false;
  try {
    SpreadsheetApp.openById(sheetId);
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * 建立使用者自己的遊戲存檔試算表（由前端按鈕驅動）
 */
function createUserSheet() {
  var userProperties = PropertiesService.getUserProperties();
  var existingId = userProperties.getProperty('GUESS_V4_BOOK_ID');
  
  if (existingId) {
    try {
      var existingSs = SpreadsheetApp.openById(existingId);
      return { success: true, url: existingSs.getUrl(), message: '已沿用現有存檔' };
    } catch (e) {}
  }

  // 建立個人雲端硬碟試算表
  var ss = SpreadsheetApp.create('猜數字版本四｜我的遊戲存檔');
  userProperties.setProperty('GUESS_V4_BOOK_ID', ss.getId());

  var sheetName = '遊戲存檔';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var defaultSheet = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
    if (defaultSheet && ss.getSheets().length > 1) {
      try { ss.deleteSheet(defaultSheet); } catch (e) {}
    }
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '使用者 Email',
      '局ID',
      '局數',
      '狀態',
      '開始時間',
      '更新時間',
      '猜測次數',
      '猜測歷程',
      '完整存檔JSON'
    ]);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#f3f4f6');
  }

  try {
    ss.setSpreadsheetTimeZone('Asia/Taipei');
  } catch(e) {}

  return { success: true, url: ss.getUrl(), message: '個人存檔建立成功！' };
}

/**
 * 取得使用者個人的試算表工作表
 */
function getUserSheet_() {
  var userProperties = PropertiesService.getUserProperties();
  var sheetId = userProperties.getProperty('GUESS_V4_BOOK_ID');
  if (!sheetId) {
    throw new Error('尚未建立個人遊戲存檔，請先點擊「建立我的遊戲存檔」。');
  }
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('遊戲存檔');
  if (!sheet) {
    throw new Error('找不到「遊戲存檔」工作表，請重新整理頁面試試。');
  }
  return { sheet: sheet, url: ss.getUrl() };
}

/**
 * 產生不重複 4 位數字
 */
function generate4DigitSecret_() {
  var digits = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'];
  var result = '';
  for (var i = 0; i < 4; i++) {
    var idx = Math.floor(Math.random() * digits.length);
    result += digits[idx];
    digits.splice(idx, 1);
  }
  return result;
}

/**
 * 計算幾 A 幾 B
 */
function calculateAB_(secret, guess) {
  var a = 0, b = 0;
  for (var i = 0; i < 4; i++) {
    if (guess[i] === secret[i]) {
      a++;
    } else if (secret.indexOf(guess[i]) !== -1) {
      b++;
    }
  }
  return { a: a, b: b };
}

/**
 * API 入口：取得個人試算表內的所有遊戲
 */
function getMyGames() {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var sheetInfo = getUserSheet_();
    var sheet = sheetInfo.sheet;
    var data = sheet.getDataRange().getValues();
    var games = [];

    for (var i = 1; i < data.length; i++) {
      var jsonStr = data[i][8];
      if (jsonStr) {
        try {
          var gameData = JSON.parse(jsonStr);
          games.push(gameData);
        } catch (e) {}
      }
    }

    return { success: true, games: games, sheetUrl: sheetInfo.url };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * API 入口：開新局
 */
function startNewGame() {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var sheetInfo = getUserSheet_();
    var sheet = sheetInfo.sheet;
    var data = sheet.getDataRange().getValues();
    
    var userGameCount = data.length - 1;
    if (userGameCount < 0) userGameCount = 0;

    if (userGameCount >= 50) {
      return { success: false, message: '您的個人存檔已達上限 50 局。' };
    }

    var email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'User';
    var gameNumber = userGameCount + 1;
    var gameId = 'game_' + Date.now();
    var secret = generate4DigitSecret_();
    var nowIso = new Date().toISOString();

    var gameObject = {
      gameId: gameId,
      gameNumber: gameNumber,
      userId: email,
      secret: secret,
      status: '未完成',
      startTime: nowIso,
      updatedTime: nowIso,
      history: []
    };

    sheet.appendRow([
      email,
      gameId,
      gameNumber,
      '未完成',
      nowIso,
      nowIso,
      0,
      '',
      JSON.stringify(gameObject)
    ]);

    return { success: true, game: gameObject, sheetUrl: sheetInfo.url };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * API 入口：送出猜測
 */
function submitGuess(gameId, guessInput) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    guessInput = (guessInput || '').trim();

    if (!/^\d{4}$/.test(guessInput)) {
      return { success: false, message: '請輸入恰好 4 位數字（可由 0 開頭）' };
    }
    if (new Set(guessInput.split('')).size !== 4) {
      return { success: false, message: '4 位數字不能重複！' };
    }

    var sheetInfo = getUserSheet_();
    var sheet = sheetInfo.sheet;
    var data = sheet.getDataRange().getValues();
    var targetRowIndex = -1;
    var gameObject = null;

    for (var i = 1; i < data.length; i++) {
      if ((data[i][1] || '').toString().trim() === gameId) {
        targetRowIndex = i + 1;
        try {
          gameObject = JSON.parse(data[i][8]);
        } catch (e) {}
        break;
      }
    }

    if (targetRowIndex === -1 || !gameObject) {
      return { success: false, message: '找不到對應的遊戲局數紀錄' };
    }

    if (gameObject.status === '已猜對') {
      return { success: false, message: '此局已經完成，無法繼續猜測' };
    }

    if (gameObject.history.length >= 100) {
      return { success: false, message: '已達單局最高 100 次猜測限制' };
    }

    var result = calculateAB_(gameObject.secret, guessInput);
    var resultStr = result.a + 'A' + result.b + 'B';
    var isWin = (result.a === 4);
    var nowIso = new Date().toISOString();

    var attemptLog = {
      attemptNumber: gameObject.history.length + 1,
      guess: guessInput,
      result: resultStr,
      timestamp: nowIso
    };

    gameObject.history.push(attemptLog);
    gameObject.updatedTime = nowIso;
    if (isWin) {
      gameObject.status = '已猜對';
    }

    var summaryHistory = gameObject.history.map(function(h) {
      return h.attemptNumber + '. ' + h.guess + ' -> ' + h.result;
    }).join(' | ');

    sheet.getRange(targetRowIndex, 4).setValue(gameObject.status);
    sheet.getRange(targetRowIndex, 6).setValue(nowIso);
    sheet.getRange(targetRowIndex, 7).setValue(gameObject.history.length);
    sheet.getRange(targetRowIndex, 8).setValue(summaryHistory);
    sheet.getRange(targetRowIndex, 9).setValue(JSON.stringify(gameObject));

    return {
      success: true,
      game: gameObject,
      lastResult: attemptLog
    };

  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
