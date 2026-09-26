/**
 * ============================================================
 * 校園生活 50 問
 * 正式多班級版
 *
 * 架構：
 * 1 個題庫
 * 18 個班級
 * localStorage 本機即時保存
 * Google Sheet 雲端備份
 * 班級切換
 * 同步狀態
 * 二次確認重置目前班級
 *
 * ------------------------------------------------------------
 * 工作表：
 *
 * 【題庫】
 * A 編號
 * B 類別
 * C 題目
 * D 選項A
 * E 選項B
 * F 選項C
 * G 選項D
 * H 正確答案
 * I 答案解析
 * J 出處
 * K 答案圖片URL
 * L 題目圖片URL
 * M 啟用
 *
 * 【班級】
 * A 班級代碼
 * B 班級名稱
 * C 啟用
 *
 * 【遊戲進度】
 * A 班級代碼
 * B 班級名稱
 * C 遊戲局號
 * D 答題紀錄JSON
 * E 正確
 * F 錯誤
 * G 已完成
 * H 最後更新
 * ============================================================
 */

const QUESTION_SHEET = '題庫';
const CLASS_SHEET = '班級';
const PROGRESS_SHEET = '遊戲進度';

const CLASS_COUNT = 18;


/**
 * ============================================================
 * Web App 入口
 * ============================================================
 */
function doGet() {

  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('校園生活50問')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}


/**
 * ============================================================
 * 正式版初始化
 *
 * 重要：
 * 這個函式「不會清除既有題庫」。
 *
 * 只做：
 * 1. 題庫不存在 → 建立空白題庫
 * 2. 班級不存在 → 建立18班
 * 3. 遊戲進度不存在 → 建立
 *
 * 可安全再次執行。
 * ============================================================
 */
function setupSystem() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  setupQuestionSheet_(ss);

  migrateQuestionSheetToV12_(ss);

  setupClassSheet_(ss);

  setupProgressSheet_(ss);

  SpreadsheetApp.flush();

  SpreadsheetApp
    .getUi()
    .alert(
      '系統初始化完成。\n\n' +
      '已確認：\n' +
      '1. 題庫\n' +
      '2. 班級\n' +
      '3. 遊戲進度\n\n' +
      '既有題庫不會被清除。'
    );
}


/**
 * ============================================================
 * 題庫
 *
 * 已存在 → 不動
 * 不存在 → 建立空白格式
 * ============================================================
 */
function setupQuestionSheet_(ss) {

  let sheet =
    ss.getSheetByName(
      QUESTION_SHEET
    );

  if (sheet) {

    return;

  }

  sheet =
    ss.insertSheet(
      QUESTION_SHEET
    );

  const headers = [

    '編號',
    '類別',
    '題目',

    '選項A',
    '選項B',
    '選項C',
    '選項D',

    '正確答案',
    '答案解析',
    '出處',
    '答案圖片URL',
    '題目圖片URL',
    '啟用'

  ];

  sheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues([headers]);

  formatHeader_(
    sheet,
    headers.length
  );

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 400);

  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 200);

  sheet.setColumnWidth(8, 90);
  sheet.setColumnWidth(9, 450);
  sheet.setColumnWidth(10, 250);
  sheet.setColumnWidth(11, 350);
  sheet.setColumnWidth(12, 350);
  sheet.setColumnWidth(13, 80);
}




/**
 * ============================================================
 * V1.2.1：修正 L / M 欄資料驗證
 *
 * L 題目圖片URL：
 * - 必須是一般文字儲存格
 * - 清除舊「啟用」欄遺留的核取方塊驗證
 *
 * M 啟用：
 * - 明確套用核取方塊驗證
 * - 不改寫既有 TRUE / FALSE 值
 * ============================================================
 */
function fixQuestionSheetValidationV121_(sheet) {

  const maxRows =
    sheet.getMaxRows();

  if (maxRows < 2) {
    return;
  }

  const dataRowCount =
    maxRows - 1;

  // L：題目圖片URL，不允許殘留 checkbox / TRUE-FALSE 驗證。
  sheet
    .getRange(
      2,
      12,
      dataRowCount,
      1
    )
    .clearDataValidations();

  // M：啟用，重新明確套用 checkbox 驗證。
  // insertCheckboxes() 只建立驗證規則，不主動改寫既有 TRUE/FALSE。
  sheet
    .getRange(
      2,
      13,
      dataRowCount,
      1
    )
    .insertCheckboxes();

}


/**
 * ============================================================
 * 題庫 V1.2：12欄 → 13欄安全遷移
 *
 * 原架構：
 * K 圖片URL（答案／宣導圖片）
 * L 啟用
 *
 * 新架構：
 * K 答案圖片URL
 * L 題目圖片URL
 * M 啟用
 *
 * 原子修改原則：
 * 1. 不清除任何既有題目。
 * 2. 不搬動 K 欄既有圖片資料。
 * 3. 在原 L 欄前插入一欄，原 L「啟用」自動右移到 M。
 * 4. 若已是 13 欄新架構，不重複插欄。
 * ============================================================
 */
function migrateQuestionSheetToV12_(ss) {

  const sheet =
    ss.getSheetByName(
      QUESTION_SHEET
    );

  if (!sheet) {
    return;
  }

  const lastColumn =
    sheet.getLastColumn();

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        Math.max(
          lastColumn,
          13
        )
      )
      .getValues()[0];

  const kHeader =
    String(
      headers[10] || ''
    ).trim();

  const lHeader =
    String(
      headers[11] || ''
    ).trim();

  const mHeader =
    String(
      headers[12] || ''
    ).trim();

  // 已完成 V1.2 遷移：只校正欄名，不插入新欄。
  if (
    lHeader === '題目圖片URL' &&
    mHeader === '啟用'
  ) {

    sheet
      .getRange(
        1,
        11,
        1,
        3
      )
      .setValues([[
        '答案圖片URL',
        '題目圖片URL',
        '啟用'
      ]]);

    fixQuestionSheetValidationV121_(
      sheet
    );

    return;
  }

  // 舊版 12 欄：K 為圖片URL，L 為啟用。
  // 只允許這個已知架構進行遷移，避免誤改未知資料表。
  if (
    (
      kHeader === '圖片URL' ||
      kHeader === '答案圖片URL'
    ) &&
    lHeader === '啟用' &&
    !mHeader
  ) {

    // 在 L 前插入新欄。
    // 原 L「啟用」及其所有資料會由 Google Sheet 自動右移到 M。
    sheet.insertColumnBefore(12);

    sheet
      .getRange(
        1,
        11,
        1,
        3
      )
      .setValues([[
        '答案圖片URL',
        '題目圖片URL',
        '啟用'
      ]]);

    sheet.setColumnWidth(
      11,
      350
    );

    sheet.setColumnWidth(
      12,
      350
    );

    sheet.setColumnWidth(
      13,
      80
    );

    fixQuestionSheetValidationV121_(
      sheet
    );

    return;
  }

  throw new Error(
    '題庫欄位架構不是已知的 V1.1 或 V1.2 格式，為避免誤改資料，已停止自動遷移。'
  );
}


/**
 * ============================================================
 * 班級
 *
 * 不存在才建立。
 *
 * 預設：
 * 01 班級01
 * 02 班級02
 * ...
 * 18 班級18
 *
 * 老師之後直接在 Sheet 修改即可。
 * ============================================================
 */
function setupClassSheet_(ss) {

  let sheet =
    ss.getSheetByName(
      CLASS_SHEET
    );

  if (sheet) {

    return;

  }

  sheet =
    ss.insertSheet(
      CLASS_SHEET
    );

  const headers = [

    '班級代碼',
    '班級名稱',
    '啟用'

  ];

  sheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues([headers]);

  formatHeader_(
    sheet,
    headers.length
  );

  const rows = [];

  for (
    let i = 1;
    i <= CLASS_COUNT;
    i++
  ) {

    const code =
      String(i)
        .padStart(
          2,
          '0'
        );

    rows.push([

      code,

      '班級' + code,

      true

    ]);

  }

  sheet
    .getRange(
      2,
      1,
      rows.length,
      3
    )
    .setValues(rows);

  sheet
    .getRange(
      2,
      3,
      rows.length,
      1
    )
    .insertCheckboxes();

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 80);
}


/**
 * ============================================================
 * 遊戲進度
 * ============================================================
 */
function setupProgressSheet_(ss) {

  let sheet =
    ss.getSheetByName(
      PROGRESS_SHEET
    );

  if (sheet) {

    return;

  }

  sheet =
    ss.insertSheet(
      PROGRESS_SHEET
    );

  const headers = [

    '班級代碼',
    '班級名稱',
    '遊戲局號',
    '答題紀錄JSON',
    '正確',
    '錯誤',
    '已完成',
    '最後更新'

  ];

  sheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues([headers]);

  formatHeader_(
    sheet,
    headers.length
  );

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 500);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 90);
  sheet.setColumnWidth(8, 180);
}


/**
 * ============================================================
 * 標題格式
 * ============================================================
 */
function formatHeader_(
  sheet,
  columnCount
) {

  sheet
    .getRange(
      1,
      1,
      1,
      columnCount
    )
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#1976d2')
    .setFontColor('#ffffff');
}


/**
 * ============================================================
 * 取得班級
 * ============================================================
 */
function getClasses() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CLASS_SHEET
      );

  if (!sheet) {

    throw new Error(
      '找不到「班級」工作表，請先執行 setupSystem()。'
    );

  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    return [];

  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        3
      )
      .getValues();

  return values

    .filter(
      function(row) {

        return (
          row[2] === true &&
          row[0] !== ''
        );

      }
    )

    .map(
      function(row) {

        return {

          code:
            String(row[0]),

          name:
            String(
              row[1] ||
              row[0]
            )

        };

      }
    );
}


/**
 * ============================================================
 * 取得題庫
 *
 * 不把正確答案與解析送到前端。
 * ============================================================
 */
function getQuestions() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        QUESTION_SHEET
      );

  if (!sheet) {

    throw new Error(
      '找不到「題庫」工作表。'
    );

  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    return [];

  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        13
      )
      .getValues();

  const result =
    values

      .filter(
        function(row) {

          return (
            row[12] === true &&
            row[0] !== '' &&
            row[2] !== ''
          );

        }
      )

      .map(
        function(row) {

          return {

            id:
              row[0],

            category:
              row[1],

            question:
              row[2],

            choices: {

              A:
                row[3],

              B:
                row[4],

              C:
                row[5],

              D:
                row[6]

            },

            questionImageUrl:
              normalizeImageUrl_(
                row[11]
              )

          };

        }
      );

  result.sort(
    function(a, b) {

      return (
        Number(a.id) -
        Number(b.id)
      );

    }
  );

  return result;
}


/**
 * ============================================================
 * 核對答案
 * ============================================================
 */
function checkAnswer(
  questionId,
  selectedAnswer
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        QUESTION_SHEET
      );

  if (!sheet) {

    throw new Error(
      '找不到「題庫」工作表。'
    );

  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    throw new Error(
      '題庫目前沒有題目。'
    );

  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        13
      )
      .getValues();

  const row =
    values.find(
      function(item) {

        return (
          String(item[0]) ===
          String(questionId)
        );

      }
    );

  if (!row) {

    throw new Error(
      '找不到第 ' +
      questionId +
      ' 題。'
    );

  }

  const correctAnswer =
    String(row[7])
      .trim()
      .toUpperCase();

  const userAnswer =
    String(selectedAnswer)
      .trim()
      .toUpperCase();

  return {

    id:
      row[0],

    correct:
      userAnswer ===
      correctAnswer,

    selectedAnswer:
      userAnswer,

    correctAnswer:
      correctAnswer,

    explanation:
      row[8],

    source:
      row[9],

    answerImageUrl:
      normalizeImageUrl_(
        row[10]
      )

  };
}


/**
 * ============================================================
 * 讀取某班雲端進度
 * ============================================================
 */
function getClassProgress(
  classCode
) {

  const code =
    String(classCode || '')
      .trim();

  if (!code) {

    throw new Error(
      '缺少班級代碼。'
    );

  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const classInfo =
    getClassInfo_(
      ss,
      code
    );

  if (!classInfo) {

    throw new Error(
      '找不到班級：' +
      code
    );

  }

  const sheet =
    ss.getSheetByName(
      PROGRESS_SHEET
    );

  if (!sheet) {

    throw new Error(
      '找不到「遊戲進度」工作表。'
    );

  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    return createEmptyProgress_(
      classInfo
    );

  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        8
      )
      .getValues();

  const row =
    values.find(
      function(item) {

        return (
          String(item[0]) ===
          code
        );

      }
    );

  if (!row) {

    return createEmptyProgress_(
      classInfo
    );

  }

  let answers = {};

  try {

    answers =
      row[3]
        ?
        JSON.parse(
          String(row[3])
        )
        :
        {};

  } catch (error) {

    answers = {};

  }

  return {

    classCode:
      code,

    className:
      classInfo.name,

    gameNo:
      Number(row[2]) || 1,

    answers:
      answers,

    correct:
      Number(row[4]) || 0,

    wrong:
      Number(row[5]) || 0,

    finished:
      Number(row[6]) || 0,

    updatedAt:
      dateToIso_(
        row[7]
      )

  };
}


/**
 * ============================================================
 * 儲存某班雲端進度
 *
 * 重要：
 * 使用 ScriptLock，避免短時間內兩次寫入互相覆蓋。
 *
 * 前端送來的 answers 會與雲端現有 answers 合併。
 *
 * 已存在的題號：
 * 不由另一份資料自動改寫。
 *
 * 這符合目前遊戲規則：
 * 每題一局只回答一次。
 * ============================================================
 */
function saveClassProgress(
  payload
) {

  if (
    !payload ||
    !payload.classCode
  ) {

    throw new Error(
      '缺少班級資料。'
    );

  }

  const lock =
    LockService.getScriptLock();

  lock.waitLock(10000);

  try {

    const ss =
      SpreadsheetApp
        .getActiveSpreadsheet();

    const code =
      String(
        payload.classCode
      ).trim();

    const classInfo =
      getClassInfo_(
        ss,
        code
      );

    if (!classInfo) {

      throw new Error(
        '找不到班級：' +
        code
      );

    }

    const sheet =
      ss.getSheetByName(
        PROGRESS_SHEET
      );

    if (!sheet) {

      throw new Error(
        '找不到「遊戲進度」工作表。'
      );

    }

    const incomingAnswers =
      sanitizeAnswers_(
        payload.answers
      );

    const lastRow =
      sheet.getLastRow();

    let targetRow =
      -1;

    let existingAnswers =
      {};

    let existingGameNo =
      1;

    if (
      lastRow >= 2
    ) {

      const values =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            8
          )
          .getValues();

      for (
        let i = 0;
        i < values.length;
        i++
      ) {

        if (
          String(values[i][0]) ===
          code
        ) {

          targetRow =
            i + 2;

          existingGameNo =
            Number(
              values[i][2]
            ) || 1;

          try {

            existingAnswers =
              values[i][3]
                ?
                JSON.parse(
                  String(
                    values[i][3]
                  )
                )
                :
                {};

          } catch (error) {

            existingAnswers =
              {};

          }

          break;

        }

      }

    }

    existingAnswers =
      sanitizeAnswers_(
        existingAnswers
      );

    /**
     * 原子合併：
     *
     * 先保留雲端已存在的題目。
     * 只補入雲端沒有的題目。
     */
    const mergedAnswers =
      Object.assign(
        {},
        existingAnswers
      );

    Object
      .keys(
        incomingAnswers
      )
      .forEach(
        function(questionId) {

          if (
            !Object.prototype
              .hasOwnProperty
              .call(
                mergedAnswers,
                questionId
              )
          ) {

            mergedAnswers[
              questionId
            ] =
              incomingAnswers[
                questionId
              ];

          } else if (
            payload.allowRetryUpgrade === true &&
            mergedAnswers[questionId] === 'wrong' &&
            incomingAnswers[questionId] === 'correct'
          ) {

            /**
             * V1.4.8 重答模式：
             * 只允許 wrong -> correct 單向升級。
             * correct 永遠不會被後續答案降回 wrong。
             */
            mergedAnswers[
              questionId
            ] =
              'correct';

          }

        }
      );

    const stats =
      calculateStats_(
        mergedAnswers
      );

    const now =
      new Date();

    const requestedGameNo =
      Number(
        payload.gameNo
      ) || 1;

    const gameNo =
      Math.max(
        existingGameNo,
        requestedGameNo
      );

    const rowData = [

      code,

      classInfo.name,

      gameNo,

      JSON.stringify(
        mergedAnswers
      ),

      stats.correct,

      stats.wrong,

      stats.finished,

      now

    ];

    if (
      targetRow === -1
    ) {

      targetRow =
        sheet.getLastRow() + 1;

    }

    sheet
      .getRange(
        targetRow,
        1,
        1,
        8
      )
      .setValues([
        rowData
      ]);

    SpreadsheetApp.flush();

    return {

      success:
        true,

      classCode:
        code,

      className:
        classInfo.name,

      gameNo:
        gameNo,

      answers:
        mergedAnswers,

      correct:
        stats.correct,

      wrong:
        stats.wrong,

      finished:
        stats.finished,

      updatedAt:
        now.toISOString()

    };

  } finally {

    lock.releaseLock();

  }
}


/**
 * ============================================================
 * 重置目前班級
 *
 * 只清除指定班級。
 * 其他17班完全不動。
 *
 * 前端負責二次確認。
 * ============================================================
 */
function resetClassProgress(
  classCode
) {

  const code =
    String(classCode || '')
      .trim();

  if (!code) {

    throw new Error(
      '缺少班級代碼。'
    );

  }

  const lock =
    LockService.getScriptLock();

  lock.waitLock(10000);

  try {

    const ss =
      SpreadsheetApp
        .getActiveSpreadsheet();

    const classInfo =
      getClassInfo_(
        ss,
        code
      );

    if (!classInfo) {

      throw new Error(
        '找不到班級：' +
        code
      );

    }

    const sheet =
      ss.getSheetByName(
        PROGRESS_SHEET
      );

    if (!sheet) {

      throw new Error(
        '找不到「遊戲進度」工作表。'
      );

    }

    const lastRow =
      sheet.getLastRow();

    let targetRow =
      -1;

    let oldGameNo =
      0;

    if (
      lastRow >= 2
    ) {

      const values =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            8
          )
          .getValues();

      for (
        let i = 0;
        i < values.length;
        i++
      ) {

        if (
          String(values[i][0]) ===
          code
        ) {

          targetRow =
            i + 2;

          oldGameNo =
            Number(
              values[i][2]
            ) || 0;

          break;

        }

      }

    }

    const newGameNo =
      oldGameNo + 1;

    const now =
      new Date();

    const rowData = [

      code,

      classInfo.name,

      newGameNo,

      '{}',

      0,

      0,

      0,

      now

    ];

    if (
      targetRow === -1
    ) {

      targetRow =
        sheet.getLastRow() + 1;

    }

    sheet
      .getRange(
        targetRow,
        1,
        1,
        8
      )
      .setValues([
        rowData
      ]);

    SpreadsheetApp.flush();

    return {

      success:
        true,

      classCode:
        code,

      className:
        classInfo.name,

      gameNo:
        newGameNo,

      answers:
        {},

      correct:
        0,

      wrong:
        0,

      finished:
        0,

      updatedAt:
        now.toISOString()

    };

  } finally {

    lock.releaseLock();

  }
}


/**
 * ============================================================
 * 找班級
 * ============================================================
 */
function getClassInfo_(
  ss,
  code
) {

  const sheet =
    ss.getSheetByName(
      CLASS_SHEET
    );

  if (!sheet) {

    return null;

  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    return null;

  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        3
      )
      .getValues();

  const row =
    values.find(
      function(item) {

        return (
          String(item[0]) ===
            String(code) &&
          item[2] === true
        );

      }
    );

  if (!row) {

    return null;

  }

  return {

    code:
      String(row[0]),

    name:
      String(
        row[1] ||
        row[0]
      )

  };
}


/**
 * ============================================================
 * 空白進度
 * ============================================================
 */
function createEmptyProgress_(
  classInfo
) {

  return {

    classCode:
      classInfo.code,

    className:
      classInfo.name,

    gameNo:
      1,

    answers:
      {},

    correct:
      0,

    wrong:
      0,

    finished:
      0,

    updatedAt:
      null

  };
}


/**
 * ============================================================
 * 清理答案
 *
 * 只接受：
 * correct
 * wrong
 * ============================================================
 */
function sanitizeAnswers_(
  answers
) {

  const result = {};

  if (
    !answers ||
    typeof answers !==
      'object'
  ) {

    return result;

  }

  Object
    .keys(
      answers
    )
    .forEach(
      function(key) {

        const value =
          answers[key];

        if (
          value === 'correct' ||
          value === 'wrong'
        ) {

          result[
            String(key)
          ] =
            value;

        }

      }
    );

  return result;
}


/**
 * ============================================================
 * 統計
 * ============================================================
 */
function calculateStats_(
  answers
) {

  let correct =
    0;

  let wrong =
    0;

  Object
    .values(
      answers
    )
    .forEach(
      function(value) {

        if (
          value ===
          'correct'
        ) {

          correct++;

        }

        if (
          value ===
          'wrong'
        ) {

          wrong++;

        }

      }
    );

  return {

    correct:
      correct,

    wrong:
      wrong,

    finished:
      correct + wrong

  };
}


/**
 * ============================================================
 * 日期轉 ISO
 * ============================================================
 */
function dateToIso_(
  value
) {

  if (!value) {

    return null;

  }

  if (
    value instanceof Date
  ) {

    return value
      .toISOString();

  }

  const date =
    new Date(value);

  if (
    isNaN(
      date.getTime()
    )
  ) {

    return null;

  }

  return date
    .toISOString();
}


/**
 * ============================================================
 * 圖片URL
 * ============================================================
 */
function normalizeImageUrl_(
  url
) {

  if (!url) {

    return '';

  }

  const text =
    String(url)
      .trim();

  if (!text) {

    return '';

  }

  let match =
    text.match(
      /\/file\/d\/([a-zA-Z0-9_-]+)/
    );

  if (
    match &&
    match[1]
  ) {

    return (
      'https://drive.google.com/thumbnail?id=' +
      match[1] +
      '&sz=w1600'
    );

  }

  match =
    text.match(
      /[?&]id=([a-zA-Z0-9_-]+)/
    );

  if (
    match &&
    match[1]
  ) {

    return (
      'https://drive.google.com/thumbnail?id=' +
      match[1] +
      '&sz=w1600'
    );

  }

  return text;
}

/**
 * ============================================================
 * V1.2.1 手動修正入口
 *
 * 使用方式：
 * Apps Script 函式選單選擇：
 * repairQuestionSheetValidationV121
 *
 * 執行一次即可。
 *
 * 只做：
 * 1. L 欄清除資料驗證
 * 2. M 欄重新套用核取方塊
 *
 * 不修改任何題目內容、圖片URL、答案或答題進度。
 * ============================================================
 */
function repairQuestionSheetValidationV121() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      QUESTION_SHEET
    );

  if (!sheet) {

    throw new Error(
      '找不到「題庫」工作表。'
    );

  }

  const lHeader =
    String(
      sheet
        .getRange(
          1,
          12
        )
        .getValue() || ''
    ).trim();

  const mHeader =
    String(
      sheet
        .getRange(
          1,
          13
        )
        .getValue() || ''
    ).trim();

  if (
    lHeader !== '題目圖片URL' ||
    mHeader !== '啟用'
  ) {

    throw new Error(
      '目前 L / M 欄不是「題目圖片URL / 啟用」，為避免誤改資料，已停止修正。'
    );

  }

  fixQuestionSheetValidationV121_(
    sheet
  );

  SpreadsheetApp.flush();

  SpreadsheetApp
    .getUi()
    .alert(
      'V1.2.1 修正完成。\\n\\n' +
      'L 欄：已清除核取方塊／資料驗證。\\n' +
      'M 欄：已重新套用啟用核取方塊。\\n\\n' +
      '既有題目與圖片網址未修改。'
    );

}


/**
 * ============================================================
 * V1.3.1d 題目資料傳輸
 * 將 getQuestions() 的結果轉成 JSON 字串送到前端。
 * 不修改題庫、不修改 getQuestions() 的內容。
 * ============================================================
 */
function getQuestionsJson() {

  const questions =
    getQuestions();

  return JSON.stringify(
    questions
  );

}


/**
 * ============================================================
 * V1.4 題庫清單摘要
 * B欄 = 題庫名稱／類別
 * M欄 = 啟用
 * 只讀取，不修改資料。
 * ============================================================
 */
function getQuestionBankSummary() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        QUESTION_SHEET
      );

  if (!sheet) {
    throw new Error(
      '找不到「題庫」工作表。'
    );
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        13
      )
      .getValues();

  const counts = {};

  values.forEach(
    function(row) {

      const category =
        String(
          row[1] || ''
        ).trim();

      if (
        row[12] === true &&
        row[0] !== '' &&
        row[2] !== '' &&
        category !== ''
      ) {
        counts[category] =
          (counts[category] || 0) + 1;
      }
    }
  );

  return Object.keys(counts)
    .map(
      function(category) {
        return {
          category: category,
          count: counts[category]
        };
      }
    );
}


/**
 * ============================================================
 * V1.4.1 穩定初始化
 *
 * 一次回傳：
 * 1. 班級
 * 2. 所有 M=TRUE 的啟用題目
 *
 * 以 JSON 字串傳輸，避免前端同時發出多個初始化請求，
 * 也避免大量物件直接跨 google.script.run 傳輸不穩定。
 * 不修改試算表資料。
 * ============================================================
 */
function getClientBootstrap() {

  return JSON.stringify({
    classes: getClasses(),
    questions: getQuestions()
  });

}
