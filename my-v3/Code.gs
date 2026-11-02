/**
 * 數位科技概論｜版本三：幾 A 幾 B 雲端存檔後端 (Code.gs)
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('幾 A 幾 B 數字猜測 (雲端代號版)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 初始化 setup：自動建立或連結 Google 試算表
 */
function setup() {
  var userProperties = PropertiesService.getScriptProperties();
  var spreadsheetId = userProperties.getProperty('GUESS_BOOK_ID');
  var ss;

  if (spreadsheetId) {
    try {
      ss = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      ss = null;
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create('數位科技概論｜幾A幾B版本三存檔');
    userProperties.setProperty('GUESS_BOOK_ID', ss.getId());
  }

  var sheetName = '遊戲存檔';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var defaultSheet = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
    if (defaultSheet && ss.getSheets().length > 1) {
      try { ss.deleteSheet(defaultSheet); } catch (e) {}
    }
  }

  // 設定表頭
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '使用者代號',
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

  // 嘗試設定時區為台北
  try {
    ss.setSpreadsheetTimeZone('Asia/Taipei');
  } catch(e) {}

  Logger.log('請保留此試算表網址：' + ss.getUrl());
  return 'setup 完成！試算表網址：' + ss.getUrl();
}

/**
 * 取得試算表 Sheet 物件
 */
function getSheet_() {
  var spreadsheetId = PropertiesService.getScriptProperties().getProperty('GUESS_BOOK_ID');
  if (!spreadsheetId) {
    throw new Error('系統尚未初始化，請先由管理員執行 setup() 函式。');
  }
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName('遊戲存檔');
  if (!sheet) {
    throw new Error('找不到「遊戲存檔」工作表，請重新執行 setup()。');
  }
  return sheet;
}

/**
 * 產生不重複的 4 位數字答案 (包含首位 0)
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
  var a = 0;
  var b = 0;
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
 * API 入口：取得某代號的所有遊戲歷史
 */
function getGamesByUserId(userId) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    userId = (userId || '').trim().toLowerCase();
    if (!userId) return { success: false, message: '使用者代號不能為空' };

    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    var games = [];

    for (var i = 1; i < data.length; i++) {
      var rowUser = (data[i][0] || '').toString().trim().toLowerCase();
      if (rowUser === userId) {
        var jsonStr = data[i][8];
        if (jsonStr) {
          try {
            var gameData = JSON.parse(jsonStr);
            games.push(gameData);
          } catch (e) {}
        }
      }
    }

    return { success: true, games: games };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * API 入口：開新局
 */
function startNewGame(userId) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    userId = (userId || '').trim().toLowerCase();
    if (!userId) return { success: false, message: '代號不能為空' };

    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    var userGameCount = 0;

    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString().trim().toLowerCase() === userId) {
        userGameCount++;
      }
    }

    if (userGameCount >= 50) {
      return { success: false, message: '已達此代號上限 50 局。' };
    }

    var gameNumber = userGameCount + 1;
    var gameId = 'game_' + userId + '_' + Date.now();
    var secret = generate4DigitSecret_();
    var nowIso = new Date().toISOString();

    var gameObject = {
      gameId: gameId,
      gameNumber: gameNumber,
      userId: userId,
      secret: secret,
      status: '未完成',
      startTime: nowIso,
      updatedTime: nowIso,
      history: []
    };

    sheet.appendRow([
      userId,
      gameId,
      gameNumber,
      '未完成',
      nowIso,
      nowIso,
      0,
      '',
      JSON.stringify(gameObject)
    ]);

    return { success: true, game: gameObject };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * API 入口：送出猜測並更新紀錄
 */
function submitGuess(userId, gameId, guessInput) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    userId = (userId || '').trim().toLowerCase();
    guessInput = (guessInput || '').trim();

    // 格式防呆
    if (!/^\d{4}$/.test(guessInput)) {
      return { success: false, message: '請輸入恰好四位數字（可由 0 開頭）' };
    }
    var uniqueDigits = new Set(guessInput.split(''));
    if (uniqueDigits.size !== 4) {
      return { success: false, message: '四位數字不能重複！' };
    }

    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    var targetRowIndex = -1;
    var gameObject = null;

    for (var i = 1; i < data.length; i++) {
      var rUser = (data[i][0] || '').toString().trim().toLowerCase();
      var rGameId = (data[i][1] || '').toString().trim();
      if (rUser === userId && rGameId === gameId) {
        targetRowIndex = i + 1; // 1-based index
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

    // 計算 A/B
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

    // 格式化簡易紀錄字串
    var summaryHistory = gameObject.history.map(function(h) {
      return h.attemptNumber + '. ' + h.guess + ' -> ' + h.result;
    }).join(' | ');

    // 更新試算表 row
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
