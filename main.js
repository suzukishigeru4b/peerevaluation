/**
 * 相互評価システム (Server-side Script)
 * Final Version (JSON版)
 * © 2025 Shigeru Suzuki
 */

// ==================================================
// 1. 設定・定数
// ==================================================
const SYSTEM_CONFIG = {
  SHEET_NAME_SETTINGS: '設定',
  SHEET_NAME_MEMBERS: 'メンバー',
  SHEET_NAME_DATA: '相互評価データ',
  COLUMN_OFFSET: 3,
  ROW_OFFSET: 2,

  RATING_OPTIONS: ['4:そう思う', '3:やや思う', '2:やや思わない', '1:思わない'],
  RATING_SCORES: [4, 3, 2, 1],

  // セキュリティ設定
  MAX_INPUT_LENGTH: 1000,
  MAX_TARGET_NAME_LENGTH: 100
};

// ==================================================
// 2. Webアプリ エントリーポイント
// ==================================================
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  const config = _fetchSystemConfig();
  return template.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle(config.pageTitle || '相互評価システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
}

// ==================================================
// 2.5 セキュリティユーティリティ
// ==================================================

/**
 * 入力値のサニタイズ処理
 * @param {string} str - サニタイズする文字列
 * @param {number} maxLength - 最大文字数
 * @returns {string} - サニタイズされた文字列
 */
function _sanitizeInput(str, maxLength = SYSTEM_CONFIG.MAX_INPUT_LENGTH) {
  if (str == null) return '';
  return String(str)
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]+>/g, '')
    .replace(/javascript:/gi, '')
    .replace(/on\w+\s*=/gi, '')
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .trim()
    .substring(0, maxLength);
}

/**
 * 文字列型かどうかの検証
 * @param {*} value - 検証する値
 * @returns {boolean}
 */
function _isValidString(value) {
  return typeof value === 'string' && value.length > 0;
}

// ==================================================
// 3. クライアント用API
// ==================================================

function fetchUserAndGroupData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const myEmail = Session.getActiveUser().getEmail();

  const sheetMembers = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAME_MEMBERS);
  const sheetEvaluate = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAME_DATA);
  const sheetSettings = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAME_SETTINGS);

  const memberData = sheetMembers.getDataRange().getValues();
  memberData.shift();

  const myData = memberData.find(row => row[0] === myEmail);
  if (!myData) {
    return JSON.stringify({ error: '名簿にあなたのメールアドレスが登録されていません。' });
  }

  // Setベースの重複除去（中間配列を削減）
  const groupSet = new Set(memberData.map(row => row[1]).filter(g => g !== ""));
  const groupList = [...groupSet].map(groupName => [groupName, '']);


  // 評価データの取得 (一度のAPI呼び出しで完結)
  const evalLastRow = sheetEvaluate.getLastRow();
  if (evalLastRow > 2) {
    const evalData = sheetEvaluate.getRange(3, 1, evalLastRow - 2, 3).getValues();

    const myEvaluationsMap = new Map();
    evalData.forEach(row => {
      if (row[1] === myEmail) {
        myEvaluationsMap.set(row[2], row[0]);
      }
    });

    groupList.forEach(groupRow => {
      const targetName = groupRow[0];
      if (myEvaluationsMap.has(targetName)) {
        groupRow[1] = myEvaluationsMap.get(targetName);
      }
    });
  }

  // 設定を直接取得 (ヘルパー関数の呼び出しを省略して最適化)
  const settingsValues = sheetSettings.getRange(2, 2, 3, 1).getValues();
  const systemConfig = {
    pageTitle: settingsValues[0][0],
    isInputEnabled: settingsValues[1][0],
    isOutputEnabled: settingsValues[2][0]
  };

  return JSON.stringify({
    systemConfig: systemConfig,
    userEmail: myEmail,
    userName: myData[2],
    userGroup: myData[1],
    groupList: groupList
  });
}

function fetchEvaluationForm(targetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const myEmail = Session.getActiveUser().getEmail();
  const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAME_DATA);

  const headers = sheet.getRange(1, 1, 2, sheet.getLastColumn()).getValues();
  const questionList = headers[0].slice(SYSTEM_CONFIG.COLUMN_OFFSET);
  const questionTypes = headers[1].slice(SYSTEM_CONFIG.COLUMN_OFFSET);

  let answerList = [];
  if (sheet.getLastRow() > 2) {
    const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn()).getValues();
    const existingRow = data.find(row => row[1] === myEmail && row[2] === targetName);
    if (existingRow) {
      answerList = existingRow.slice(SYSTEM_CONFIG.COLUMN_OFFSET);
    }
  }

  if (answerList.length === 0) {
    answerList = new Array(questionList.length).fill('');
  }

  return JSON.stringify({
    targetName: targetName,
    questionList: questionList,
    questionTypes: questionTypes,
    ratingOptions: SYSTEM_CONFIG.RATING_OPTIONS,
    answerList: answerList
  });
}

function saveEvaluationData(jsonString) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const data = JSON.parse(jsonString);
      const targetName = data.targetName;

      // 入力検証
      if (!_isValidString(targetName)) {
        throw new Error("評価対象が指定されていません。");
      }

      if (targetName.length > SYSTEM_CONFIG.MAX_TARGET_NAME_LENGTH) {
        throw new Error("評価対象名が長すぎます。");
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetMembers = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAME_MEMBERS);
      const memberData = sheetMembers.getDataRange().getValues();
      const exists = memberData.some(row => row[1] === targetName);
      if (!exists) {
        throw new Error("不正な評価対象です。");
      }

      // 回答のサニタイズと検証
      if (!Array.isArray(data.answerList)) {
        throw new Error("不正なデータ形式です。");
      }

      const sanitizedAnswers = data.answerList.map(ans => {
        const sanitized = _sanitizeInput(ans);
        return sanitized.replace(/\r?\n/g, '');
      });

      const myEmail = Session.getActiveUser().getEmail();
      const now = new Date();

      const sheetData = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAME_DATA);
      const rowData = [now, myEmail, targetName, ...sanitizedAnswers];
      const lastRow = sheetData.getLastRow();

      if (lastRow <= 2) {
        sheetData.appendRow(rowData);
      } else {
        const range = sheetData.getRange(3, 1, lastRow - 2, 3);
        const metaData = range.getValues();
        const updateIndex = metaData.findIndex(row => row[1] === myEmail && row[2] === targetName);

        if (updateIndex >= 0) {
          sheetData.getRange(updateIndex + 3, 1, 1, rowData.length).setValues([rowData]);
        } else {
          sheetData.appendRow(rowData);
        }
      }
      return JSON.stringify({ status: 'success' });

    } catch (e) {
      return JSON.stringify({ status: 'error', message: e.message });
    } finally {
      lock.releaseLock();
    }
  } else {
    return JSON.stringify({ status: 'busy', message: 'サーバー混雑中。再試行してください。' });
  }
}

function fetchFeedbackResults(targetGroupName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAME_DATA);

  const headers = sheet.getRange(1, 1, 2, sheet.getLastColumn()).getValues();
  const questionList = headers[0].slice(SYSTEM_CONFIG.COLUMN_OFFSET);
  const questionTypes = headers[1].slice(SYSTEM_CONFIG.COLUMN_OFFSET);

  const lastRow = sheet.getLastRow();
  if (lastRow <= SYSTEM_CONFIG.ROW_OFFSET) {
    return JSON.stringify({ summaryData: [], commentList: [] });
  }

  const allData = sheet.getRange(
    SYSTEM_CONFIG.ROW_OFFSET + 1, 1, lastRow - SYSTEM_CONFIG.ROW_OFFSET, sheet.getLastColumn()
  ).getValues();
  const targetData = allData.filter(row => row[2] === targetGroupName);

  const summaryData = [];
  summaryData.push(['項目', ...SYSTEM_CONFIG.RATING_OPTIONS]);

  const commentList = [];

  questionTypes.forEach((type, idx) => {
    const qText = questionList[idx];

    if (type === 'R') {
      const counts = new Array(SYSTEM_CONFIG.RATING_OPTIONS.length).fill(0);
      let sum = 0;
      let validCount = 0;

      targetData.forEach(row => {
        const ans = row[idx + SYSTEM_CONFIG.COLUMN_OFFSET];
        const selectIndex = SYSTEM_CONFIG.RATING_OPTIONS.indexOf(ans);
        if (selectIndex >= 0) {
          counts[selectIndex]++;
          sum += SYSTEM_CONFIG.RATING_SCORES[selectIndex];
          validCount++;
        }
      });

      let ave = validCount > 0 ? (sum / validCount).toFixed(2) : "0.00";
      const label = `${qText}\n(評価者:${validCount}人, 平均:${ave})`;
      summaryData.push([label, ...counts]);

    } else if (type === 'T') {
      const comments = [];
      comments.push(qText);

      targetData.forEach(row => {
        const ans = row[idx + SYSTEM_CONFIG.COLUMN_OFFSET];
        if (ans && ans.toString().trim() !== "") {
          comments.push(ans.toString().replace(/\r?\n/g, '').trim());
        }
      });
      commentList.push(comments);
    }
  });

  return JSON.stringify({
    summaryData: summaryData,
    commentList: commentList
  });
}

// ==================================================
// 4. 管理者用メニュー
// ==================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu("結果集計")
    .addItem("評価結果の集計", "aggregateResults")
    .addItem("評価結果クリア", "resetAggregation")
    .addToUi();
}

function resetAggregation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAME_MEMBERS);
  if (sheet.getLastColumn() > SYSTEM_CONFIG.COLUMN_OFFSET) {
    sheet.getRange(1, SYSTEM_CONFIG.COLUMN_OFFSET + 1, sheet.getLastRow(), sheet.getLastColumn() - SYSTEM_CONFIG.COLUMN_OFFSET).clear();
  }
}

function aggregateResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetData = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAME_DATA);
  const sheetMembers = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAME_MEMBERS);

  const headers = sheetData.getRange(1, SYSTEM_CONFIG.COLUMN_OFFSET + 1, 2, sheetData.getLastColumn() - SYSTEM_CONFIG.COLUMN_OFFSET).getValues();
  const questionList = headers[0];
  const questionTypes = headers[1];

  if (sheetData.getLastRow() <= 2) return;
  const answers = sheetData.getRange(3, 1, sheetData.getLastRow() - 2, sheetData.getLastColumn()).getValues();

  const targets = [...new Set(answers.map(row => row[2]))];
  const resultsMap = {};

  targets.forEach(target => {
    const targetAnswers = answers.filter(row => row[2] === target);
    const count = targetAnswers.length;
    const rowResult = [count];

    questionTypes.forEach((type, qIdx) => {
      if (type === 'R') {
        let sum = 0;
        let valid = 0;
        targetAnswers.forEach(ansRow => {
          const val = ansRow[qIdx + SYSTEM_CONFIG.COLUMN_OFFSET];
          const sIdx = SYSTEM_CONFIG.RATING_OPTIONS.indexOf(val);
          if (sIdx >= 0) {
            sum += SYSTEM_CONFIG.RATING_SCORES[sIdx];
            valid++;
          }
        });
        rowResult.push(valid > 0 ? (sum / valid) : '');
      } else {
        rowResult.push('');
      }
    });
    resultsMap[target] = rowResult;
  });

  const memberData = sheetMembers.getDataRange().getValues();
  const newHeader = ['評価者数', ...questionList];
  sheetMembers.getRange(1, SYSTEM_CONFIG.COLUMN_OFFSET + 1, 1, newHeader.length).setValues([newHeader]);

  const outputData = [];
  for (let i = 1; i < memberData.length; i++) {
    const groupName = memberData[i][1];
    if (resultsMap[groupName]) {
      outputData.push(resultsMap[groupName]);
    } else {
      outputData.push(new Array(newHeader.length).fill(''));
    }
  }

  if (outputData.length > 0) {
    sheetMembers.getRange(2, SYSTEM_CONFIG.COLUMN_OFFSET + 1, outputData.length, outputData[0].length).setValues(outputData);
  }
}

function _fetchSystemConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAME_SETTINGS);
  const values = sheet.getRange(2, 2, 3, 1).getValues();
  return {
    pageTitle: values[0][0],
    isInputEnabled: values[1][0],
    isOutputEnabled: values[2][0]
  };
}