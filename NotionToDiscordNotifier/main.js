/**
 * @fileoverview NotionからのWebhookを受け取り、変更内容に応じてDiscordに通知し、
 * スプレッドシートに処理ログとページの状態を記録するスクリプト。
 * @version 2.0.0
 */

// ------------------------------------------------------------------------------------
// 設定 (CONFIG)
// ------------------------------------------------------------------------------------
/**
 * スクリプト全体の設定を管理するオブジェクト。
 * @const
 */
const CONFIG = {
  // スクリプトプロパティのキー
  PROP_KEYS: {
    SPREADSHEET_ID: "SPREADSHEET_ID_SECRET",
    LOG_SHEET_NAME: "SHEET_NAME_VALUE", // ログ用シート名
    DISCORD_URL: "DISCORD_WEBHOOK_URL_SECRET",
    DISCORD_URL_SHAROUSHI: "DISCORD_WEBHOOK_URL_SHAROUSHI",
  },
  // シート関連の設定
  SHEETS: {
    LOG: {
      NAME_DEFAULT: "NotionWebhookLog", // ログシート名のデフォルト値
      HEADERS: [
        "受信日時",
        "企業名",
        "ステータス",
        "担当",
        "全体処理ステータス",
        "Discord通知結果",
        "実行ログ・エラー詳細",
        "受信データ(raw)",
        "イベント全体(raw)",
      ],
    },
    STATE: {
      NAME: "ページ状態履歴", // ページ状態を保存するシート名
      HEADERS: ["Page ID", "Last Known Properties (JSON)"],
      COLUMN_INDEX: {
        PAGE_ID: 0, // A列
        PROPERTIES_JSON: 1, // B列
      },
    },
  },
};

// ------------------------------------------------------------------------------------
// メイン処理 (エントリーポイント)
// ------------------------------------------------------------------------------------

/**
 * Notion WebhookからのPOSTリクエストを処理するメイン関数。
 * @param {GoogleAppsScript.Events.DoPost} e - Webhookイベントオブジェクト。
 * @return {GoogleAppsScript.Content.TextOutput} 処理結果を示すJSONレスポンス。
 */
function doPost(e) {
  const result = processWebhook(e);

  return ContentService.createTextOutput(
    JSON.stringify({ status: result.status, message: result.message })
  ).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Webhook処理のビジネスロジック全体を管理する。
 * @param {GoogleAppsScript.Events.DoPost} e - Webhookイベントオブジェクト。
 * @return {{status: string, message: string}} 処理結果。
 */
function processWebhook(e) {
  const executionLogs = ["--- Webhook Execution Start ---"];
  const logData = {
    timestamp: new Date(),
    rawContents: "N/A",
    fullEventString: "N/A",
    notionInfo: null,
    overallStatus: "成功",
    discordSendStatus: [],
  };

  let config;
  let logSheet;

  try {
    // 1. 設定を読み込み
    config = loadConfig(executionLogs);
    const { spreadsheetId, logSheetName, stateSheetName } = config;

    // 2. シートオブジェクトを取得
    logSheet = getOrCreateSheet(
      spreadsheetId,
      logSheetName,
      CONFIG.SHEETS.LOG.HEADERS
    );
    const stateSheet = getOrCreateSheet(
      spreadsheetId,
      stateSheetName,
      CONFIG.SHEETS.STATE.HEADERS
    );

    // 3. Webhookデータを解析
    const webhookData = parseWebhookEvent(e, executionLogs);
    logData.rawContents = webhookData.rawContents;
    logData.fullEventString = webhookData.fullEventString;

    const currentPageData = webhookData.notionPageData;
    if (!currentPageData || !currentPageData.id) {
      throw new Error("Notion page data or Page ID could not be parsed.");
    }
    const pageId = currentPageData.id;

    // 4. 現在の情報を抽出
    const currentInfo = extractNotionInfo(currentPageData, executionLogs);
    logData.notionInfo = currentInfo;

    // 5. 以前の状態を取得し、以前の情報を抽出
    const previousState = getPreviousState(stateSheet, pageId, executionLogs);
    const previousInfo = previousState
      ? extractNotionInfo(
          { properties: previousState.properties },
          executionLogs,
          "(変更前データなし)"
        )
      : null;

    // 6. 状態を比較し、必要に応じてDiscordに通知
    const notificationResults = handleNotifications(
      previousInfo,
      currentInfo,
      config,
      executionLogs
    );
    logData.discordSendStatus = notificationResults.map(
      (r) => `${r.type}: ${r.status}`
    );
    if (notificationResults.some((r) => r.error)) {
      logData.overallStatus = "一部エラー";
    }

    // 7. 現在の状態で「ページ状態履歴」シートを更新
    updateCurrentState(
      stateSheet,
      pageId,
      currentPageData.properties,
      previousState ? previousState.rowIndex : -1,
      executionLogs
    );
  } catch (error) {
    executionLogs.push(
      `Critical Error: ${error.message} \n${error.stack || ""}`
    );
    logData.overallStatus = "致命的エラー";
    Logger.log(`Critical Error in processWebhook: ${error.toString()}`);
  } finally {
    // 8. 処理結果をログシートに記録
    if (logSheet) {
      logToSheet(logSheet, logData, executionLogs);
    } else {
      Logger.log(
        `CRITICAL: Log sheet was not available. Logs:\n${executionLogs.join("\n")}`
      );
    }
  }

  executionLogs.push("--- Webhook Execution End ---");
  Logger.log(executionLogs.join("\n"));

  return { status: logData.overallStatus, message: "Webhook processed." };
}

// ------------------------------------------------------------------------------------
// 設定・準備関数
// ------------------------------------------------------------------------------------

/**
 * スクリプトプロパティから設定値を読み込み、オブジェクトとして返す。
 * @param {string[]} executionLogs - 実行ログ配列。
 * @return {object} 読み込んだ設定値のオブジェクト。
 * @throws {Error} 必須の設定値が見つからない場合。
 */
function loadConfig(executionLogs) {
  const properties = PropertiesService.getScriptProperties();
  const getProp = (key) => properties.getProperty(key);

  const config = {
    spreadsheetId: getProp(CONFIG.PROP_KEYS.SPREADSHEET_ID),
    logSheetName:
      getProp(CONFIG.PROP_KEYS.LOG_SHEET_NAME) ||
      CONFIG.SHEETS.LOG.NAME_DEFAULT,
    stateSheetName: CONFIG.SHEETS.STATE.NAME,
    discordUrl: getProp(CONFIG.PROP_KEYS.DISCORD_URL),
    discordUrlSharoushi: getProp(CONFIG.PROP_KEYS.DISCORD_URL_SHAROUSHI),
  };

  if (!config.spreadsheetId) {
    executionLogs.push(
      `FATAL: Script Property "${CONFIG.PROP_KEYS.SPREADSHEET_ID}" is not set.`
    );
    throw new Error(
      `Script Property "${CONFIG.PROP_KEYS.SPREADSHEET_ID}" is not set.`
    );
  }

  executionLogs.push("Configuration loaded successfully.");
  return config;
}

/**
 * 指定されたIDのスプレッドシートから、指定された名前のシートを取得または作成する。
 * @param {string} spreadsheetId - スプレッドシートID。
 * @param {string} sheetName - シート名。
 * @param {string[]} headers - シートが存在しない場合に設定するヘッダー行。
 * @return {GoogleAppsScript.Spreadsheet.Sheet} シートオブジェクト。
 * @throws {Error} スプレッドシートへのアクセスに失敗した場合。
 */
function getOrCreateSheet(spreadsheetId, sheetName, headers) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      if (headers && headers.length > 0) {
        sheet.appendRow(headers);
      }
    }
    return sheet;
  } catch (error) {
    throw new Error(
      `Failed to access or setup spreadsheet/sheet. ID: ${spreadsheetId}, Name: ${sheetName}. Details: ${error.message}`
    );
  }
}

// ------------------------------------------------------------------------------------
// Webhookデータ処理関数
// ------------------------------------------------------------------------------------

/**
 * WebhookイベントオブジェクトからNotionページデータを解析・抽出する。
 * @param {GoogleAppsScript.Events.DoPost} e - Webhookイベントオブジェクト。
 * @param {string[]} executionLogs - 実行ログ配列。
 * @return {{rawContents: string, fullEventString: string, notionPageData: object|null}} 解析結果。
 */
function parseWebhookEvent(e, executionLogs) {
  executionLogs.push("Parsing webhook event...");
  if (!e) {
    executionLogs.push("ERROR: Event object 'e' is undefined or null.");
    return {
      rawContents: "N/A",
      fullEventString: "Event object 'e' is undefined or null",
      notionPageData: null,
    };
  }

  const fullEventString = JSON.stringify(e);
  if (!e.postData || !e.postData.contents) {
    executionLogs.push("WARN: e.postData or e.postData.contents is missing.");
    return {
      rawContents: "e.postData or e.postData.contents is missing",
      fullEventString,
      notionPageData: null,
    };
  }

  const rawContents = e.postData.contents;
  try {
    const eventData = JSON.parse(rawContents);
    const notionPageData = eventData.data;
    if (notionPageData) {
      executionLogs.push("Notion Page Data parsed successfully.");
      return { rawContents, fullEventString, notionPageData };
    } else {
      executionLogs.push(
        "WARN: 'data' property for Notion page is missing in parsed content."
      );
      return { rawContents, fullEventString, notionPageData: null };
    }
  } catch (parseError) {
    executionLogs.push(
      `ERROR: Parsing received contents as JSON failed: ${parseError.toString()}`
    );
    return { rawContents, fullEventString, notionPageData: null };
  }
}

/**
 * Notionページデータから通知やログに必要な情報を抽出する。
 * @param {object} notionPageData - Notionのページオブジェクト。
 * @param {string[]} executionLogs - 実行ログ配列。
 * @return {{kigyoMei: string, status: string|null, tanto: string|null, sharoushi: string|null, pageUrl: string, lastEditedTime: string}} 抽出された情報。
 */
function extractNotionInfo(notionPageData, executionLogs) {
  const properties = notionPageData.properties;
  if (!properties) {
    executionLogs.push(
      "ERROR: 'properties' object is missing in notionPageData."
    );
    return {
      kigyoMei: "企業名不明",
      status: null,
      tanto: null,
      sharoushi: null,
      pageUrl: "URL不明",
      lastEditedTime: "日時不明",
    };
  }

  const getPlainText = (prop) => prop?.title?.[0]?.plain_text;
  const getStatusName = (prop) => prop?.status?.name;
  const getSelectName = (prop) => prop?.select?.name;

  const info = {
    kigyoMei: getPlainText(properties["企業名"]) || "企業名不明",
    status: getStatusName(properties["商談ステータス"]) || null,
    tanto: getSelectName(properties["担当"]) || null,
    sharoushi: getStatusName(properties["社労士連携"]) || null,
    pageUrl: notionPageData.url || "URL不明",
    lastEditedTime: notionPageData.last_edited_time
      ? Utilities.formatDate(
          new Date(notionPageData.last_edited_time),
          "Asia/Tokyo",
          "yyyy/MM/dd HH:mm:ss"
        )
      : "日時不明",
  };
  executionLogs.push(
    `Extracted Info => 企業名: ${info.kigyoMei}, ステータス: ${info.status || "N/A"}, 担当: ${info.tanto || "N/A"}, 社労士連携: ${info.sharoushi || "N/A"}`
  );
  return info;
}

// ------------------------------------------------------------------------------------
// 状態管理関数 (スプレッドシート)
// ------------------------------------------------------------------------------------

/**
 * 「ページ状態履歴」シートから指定されたページの以前の状態を取得する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} stateSheet - 状態保存用シート。
 * @param {string} pageId - NotionページID。
 * @param {string[]} executionLogs - 実行ログ配列。
 * @return {{rowIndex: number, properties: object}|null} 見つかった場合は行番号とプロパティオブジェクト、なければnull。
 */
function getPreviousState(stateSheet, pageId, executionLogs) {
  executionLogs.push(`Searching for previous state of page ID: ${pageId}`);
  const data = stateSheet.getDataRange().getValues();
  // ヘッダーを除き、下から検索（最新の状態が見つかりやすいため）
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][CONFIG.SHEETS.STATE.COLUMN_INDEX.PAGE_ID] === pageId) {
      try {
        const properties = JSON.parse(
          data[i][CONFIG.SHEETS.STATE.COLUMN_INDEX.PROPERTIES_JSON]
        );
        executionLogs.push(`Previous state found at row ${i + 1}.`);
        return { rowIndex: i + 1, properties };
      } catch (e) {
        executionLogs.push(
          `ERROR: Failed to parse stored JSON for page ${pageId} at row ${i + 1}.`
        );
        return null; // パース失敗時は状態なしとみなす
      }
    }
  }
  executionLogs.push("No previous state found for this page.");
  return null;
}

/**
 * 「ページ状態履歴」シートのレコードを更新または新規作成する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} stateSheet - 状態保存用シート。
 * @param {string} pageId - NotionページID。
 * @param {object} newProperties - 保存する新しいプロパティオブジェクト。
 * @param {number} rowIndex - 更新対象の行番号。新規の場合は-1。
 * @param {string[]} executionLogs - 実行ログ配列。
 */
function updateCurrentState(
  stateSheet,
  pageId,
  newProperties,
  rowIndex,
  executionLogs
) {
  const newPropertiesJson = JSON.stringify(newProperties);
  if (rowIndex > -1) {
    stateSheet
      .getRange(rowIndex, CONFIG.SHEETS.STATE.COLUMN_INDEX.PROPERTIES_JSON + 1)
      .setValue(newPropertiesJson);
    executionLogs.push(`Updated state for page ${pageId} at row ${rowIndex}.`);
  } else {
    stateSheet.appendRow([pageId, newPropertiesJson]);
    executionLogs.push(`Appended new state for page ${pageId}.`);
  }
}

// ------------------------------------------------------------------------------------
// Discord通知関連関数
// ------------------------------------------------------------------------------------

/**
 * 変更内容に応じて、関連するすべての通知を処理するハブ関数。
 * @param {object|null} previousInfo - 変更前の情報オブジェクト。
 * @param {object} currentInfo - 変更後の情報オブジェクト。
 * @param {object} config - 読み込まれた設定オブジェクト。
 * @param {string[]} executionLogs - 実行ログ配列。
 * @return {Array<{type: string, status: string, error: string|null}>} 各通知の送信結果の配列。
 */
function handleNotifications(previousInfo, currentInfo, config, executionLogs) {
  const results = [];

  // --- 意味のある変更があったかどうかを判定 ---
  // isNewPage: 新規作成されたページか
  const isNewPage = !previousInfo;
  // needsStatusNotification: ステータス/担当が「新規設定」または「変更」されたか
  const needsStatusNotification =
    (isNewPage && (currentInfo.status || currentInfo.tanto)) ||
    (!isNewPage &&
      (previousInfo.status !== currentInfo.status ||
        previousInfo.tanto !== currentInfo.tanto));

  // needsSharoushiNotification: 社労士連携が「新規設定」または「変更」されたか
  const needsSharoushiNotification =
    (isNewPage && currentInfo.sharoushi) ||
    (!isNewPage && previousInfo.sharoushi !== currentInfo.sharoushi);
  // -----------------------------------------

  // 1. ステータスまたは担当者の変更通知
  if (needsStatusNotification) {
    executionLogs.push(
      "Status or Tanto change detected. Preparing notification."
    );
    const message = createStatusChangeMessage(previousInfo, currentInfo);
    const result = sendDiscordMessage(
      config.discordUrl,
      { content: message },
      executionLogs
    );
    results.push({ type: "ステータス・担当者変更", ...result });
  }

  // 2. 社労士連携ステータスの変更通知
  if (needsSharoushiNotification) {
    executionLogs.push(
      "Sharoushi status change detected. Preparing notification."
    );
    const message = createSharoushiUpdateMessage(previousInfo, currentInfo); // この関数も必要に応じてnullハンドリングを調整してください
    const result = sendDiscordMessage(
      config.discordUrlSharoushi,
      { content: message },
      executionLogs
    );
    results.push({ type: "社労士連携", ...result });
  }

  if (results.length === 0) {
    executionLogs.push("No significant changes detected for notification.");
  }

  return results;
}

/**
 * ステータス・担当者変更通知用のDiscordメッセージ本文を作成する。
 * @param {object|null} previousInfo - 変更前の情報オブジェクト。
 * @param {object} currentInfo - 変更後の情報オブジェクト。
 * @return {string} Discordメッセージ文字列。
 */
function createStatusChangeMessage(previousInfo, currentInfo) {
  const { kigyoMei, pageUrl, lastEditedTime } = currentInfo;
  let messageBody = `**企業名:** ${kigyoMei}\n`;

  // nullの場合の表示文字列を定義
  const na = "（未設定）";

  const prevStatus = previousInfo ? previousInfo.status || na : na;
  const currentStatus = currentInfo.status || na;
  const prevTanto = previousInfo ? previousInfo.tanto || na : na;
  const currentTanto = currentInfo.tanto || na;

  if (currentStatus !== prevStatus) {
    messageBody += `**商談ステータス:** **\`${prevStatus}\`** → **\`${currentStatus}\`** に変更\n`;
  } else {
    messageBody += `**商談ステータス:** ${currentStatus}\n`;
  }

  if (currentTanto !== prevTanto) {
    messageBody += `**担当:** **\`${prevTanto}\`** → **\`${currentTanto}\`** に変更\n`;
  } else {
    messageBody += `**担当:** ${currentTanto}\n`;
  }

  messageBody += `**最終更新日時:** ${lastEditedTime}\n`;

  return (
    `**【商談ステータス】更新通知** 📢\n` +
    `------------------------------------\n` +
    messageBody +
    `------------------------------------\n` +
    `詳細はこちら: ${pageUrl}`
  );
}

/**
 * 社労士連携ステータス変更通知用のDiscordメッセージ本文を作成する。
 * @param {object|null} previousInfo - 変更前の情報オブジェクト。
 * @param {object} currentInfo - 変更後の情報オブジェクト。
 * @return {string} Discordメッセージ文字列。
 */
function createSharoushiUpdateMessage(previousInfo, currentInfo) {
  const { kigyoMei, sharoushi, pageUrl } = currentInfo;
  const prevSharoushi = previousInfo
    ? previousInfo.sharoushi || "（未設定）"
    : "（変更前データなし）";

  const changeMessage = `**\`${prevSharoushi}\`** → **\`${sharoushi || "（未設定）"}\`** に変更されました。`;

  return (
    `**【社労士連携】更新通知** 🔔\n` +
    `------------------------------------\n` +
    `**企業名:** ${kigyoMei}\n` +
    `**担当:** ${currentInfo.tanto || "（未設定）"}\n` +
    `**連携ステータス:** ${changeMessage}\n` +
    `------------------------------------\n` +
    `詳細はこちら: ${pageUrl}`
  );
}

/**
 * 共通のDiscordメッセージ送信関数。
 * @param {string} webhookUrl - 送信先のDiscord Webhook URL。
 * @param {object} payload - 送信するペイロードオブジェクト（例: {content: "..."}）。
 * @param {string[]} executionLogs - 実行ログ配列。
 * @return {{status: string, error: string|null}} 送信結果。
 */
function sendDiscordMessage(webhookUrl, payload, executionLogs) {
  if (
    !webhookUrl ||
    !webhookUrl.startsWith("https://discord.com/api/webhooks/")
  ) {
    const warning = `WARN: Discord Webhook URL is not configured or invalid. Skipping notification.`;
    executionLogs.push(warning);
    return { status: "URL未設定/不正のためスキップ", error: null };
  }

  try {
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
    };
    UrlFetchApp.fetch(webhookUrl, options);
    executionLogs.push("Successfully sent message to Discord.");
    return { status: "送信成功", error: null };
  } catch (e) {
    executionLogs.push(
      `ERROR: Sending message to Discord failed: ${e.toString()}`
    );
    return { status: "送信エラー", error: e.toString() };
  }
}

// ------------------------------------------------------------------------------------
// ロギング関数
// ------------------------------------------------------------------------------------

/**
 * 処理の最終結果をスプレッドシートに記録する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - ログ記録用シート。
 * @param {object} logData - 記録するログデータのオブジェクト。
 * @param {string[]} executionLogs - 全ての実行ログを含む配列。
 */
function logToSheet(sheet, logData, executionLogs) {
  try {
    const {
      timestamp,
      rawContents,
      fullEventString,
      notionInfo,
      overallStatus,
      discordSendStatus,
    } = logData;
    const { kigyoMei, status, tanto } = notionInfo || {
      kigyoMei: "N/A",
      status: "N/A",
      tanto: "N/A",
    };

    sheet.appendRow([
      timestamp,
      kigyoMei,
      status,
      tanto,
      overallStatus,
      discordSendStatus.join("\n"),
      executionLogs.join("\n"),
      rawContents,
      fullEventString,
    ]);
  } catch (error) {
    Logger.log(
      `CRITICAL: Error appending final log to spreadsheet: ${error.toString()}\n${error.stack}`
    );
  }
}
