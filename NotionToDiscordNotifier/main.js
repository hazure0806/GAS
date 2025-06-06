// ------------------------------------------------------------------------------------
// スクリプトプロパティから設定値を読み込む
// ------------------------------------------------------------------------------------
const PROP_KEY_SPREADSHEET_ID = "SPREADSHEET_ID_SECRET";
const PROP_KEY_SHEET_NAME = "SHEET_NAME_VALUE"; // これはログ用シート名
const PROP_KEY_DISCORD_URL = "DISCORD_WEBHOOK_URL_SECRET";
const PROP_KEY_DISCORD_URL_SHAROUSHI = "DISCORD_WEBHOOK_URL_SHAROUSHI";

function getScriptPropertyValue(key, defaultValue = null) {
  const properties = PropertiesService.getScriptProperties();
  const value = properties.getProperty(key);
  if (value === null && defaultValue === null) {
    Logger.log(`警告: スクリプトプロパティ "${key}" が設定されていません。`);
  }
  return value !== null ? value : defaultValue;
}

const SPREADSHEET_ID = getScriptPropertyValue(PROP_KEY_SPREADSHEET_ID);
const LOG_SHEET_NAME = getScriptPropertyValue(
  PROP_KEY_SHEET_NAME,
  "NotionWebhookLog"
);
const DISCORD_WEBHOOK_URL = getScriptPropertyValue(PROP_KEY_DISCORD_URL);
const DISCORD_WEBHOOK_URL_SHAROUSHI = getScriptPropertyValue(
  PROP_KEY_DISCORD_URL_SHAROUSHI
);

// ★★★ 新しい設定項目 ★★★
const STATE_SHEET_NAME = "ページ状態履歴"; // ページの状態を記憶するシート名
const STATE_HEADERS = ["Page ID", "Last Known Properties (JSON)"]; // 状態保存シートのヘッダー
// ★★★★★★★★★★★★★★

const LOG_HEADERS = [
  "受信日時",
  "受信データ (raw)",
  "イベント全体 (raw)",
  "企業名",
  "ステータス",
  "担当",
  "全体処理ステータス",
  "Discord通知結果",
  "実行ログ・エラー詳細",
];

// ------------------------------------------------------------------------------------
// メイン処理
// ------------------------------------------------------------------------------------
function doPost(e) {
  let executionLogs = [];
  executionLogs.push("--- doPost Execution Start ---");
  let overallStatus = "成功";
  let discordSendStatus = [];
  let timestamp = new Date();
  let webhookData = null;
  let extractedNotionInfo = null;
  let logSheet, stateSheet;

  try {
    if (!SPREADSHEET_ID) {
      throw new Error(
        `Script Property "${PROP_KEY_SPREADSHEET_ID}" is not set.`
      );
    }
    // ログ用シートと状態保存用シートの両方を取得
    logSheet = getOrCreateSheet(SPREADSHEET_ID, LOG_SHEET_NAME, LOG_HEADERS);
    stateSheet = getOrCreateSheet(
      SPREADSHEET_ID,
      STATE_SHEET_NAME,
      STATE_HEADERS
    );

    webhookData = parseWebhookEvent(e, executionLogs);
    if (!webhookData.notionPageData) {
      throw new Error("Notion page data could not be parsed from webhook.");
    }

    const pageId = webhookData.notionPageData.id;
    if (!pageId) throw new Error("Page ID is missing in the webhook data.");

    // 1. 以前の状態を取得
    const previousState = getPreviousState(stateSheet, pageId, executionLogs);
    const previousProperties = previousState ? previousState.properties : null;

    // 2. 現在の状態と以前の状態から、通知に必要な情報を抽出
    extractedNotionInfo = extractNotionInfo(
      webhookData.notionPageData,
      executionLogs
    );
    const previousNotionInfo = previousProperties
      ? extractNotionInfo({ properties: previousProperties }, executionLogs)
      : null;

    // 3. 状態を比較し、Discordに通知
    // ★★★ 通知の分岐ロジック ★★★
    // 1. ステータス変更の通知
    if (
      !previousNotionInfo ||
      previousNotionInfo.status !== extractedNotionInfo.status
    ) {
      const result = sendStateChangeDiscordNotification(
        previousNotionInfo,
        extractedNotionInfo,
        DISCORD_WEBHOOK_URL,
        executionLogs
      );
      discordSendStatus.push(`ステータス通知: ${result.status}`);
      if (result.error) overallStatus = "一部エラー";
    }

    // 2. 社労士連携の通知
    if (
      !previousNotionInfo ||
      previousNotionInfo.sharoushi !== extractedNotionInfo.sharoushi
    ) {
      const result = sendSharoushiDiscordNotification(
        previousNotionInfo,
        extractedNotionInfo,
        DISCORD_WEBHOOK_URL_SHAROUSHI,
        executionLogs
      );
      discordSendStatus.push(`社労士連携通知: ${result.status}`);
      if (result.error) overallStatus = "一部エラー";
    }
    // ★★★★★★★★★★★★★★★★★★

    // 4. 現在の状態で「ページ状態履歴」シートを更新
    updateCurrentState(
      stateSheet,
      pageId,
      webhookData.notionPageData.properties,
      previousState ? previousState.rowIndex : -1,
      executionLogs
    );
  } catch (error) {
    executionLogs.push(`Critical Error in doPost: ${error.toString()}`);
    overallStatus = "致命的エラー";
  } finally {
    if (logSheet) {
      // 5. 処理結果をログシートに記録
      logToSheet(
        logSheet,
        timestamp,
        webhookData,
        extractedNotionInfo,
        overallStatus,
        discordSendStatus,
        executionLogs
      );
    } else {
      Logger.log(
        `CRITICAL: Log sheet was not available. Logs: ${executionLogs.join("\n")}`
      );
    }
  }

  executionLogs.push("--- doPost Execution End ---");
  Logger.log(executionLogs.join("\n"));

  return ContentService.createTextOutput(
    JSON.stringify({ status: overallStatus, message: "Webhook processed." })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------------------------------
// ヘルパー関数群 (★マークは新規または大幅修正された関数)
// ------------------------------------------------------------------------------------
/**
 * ★ 以前の状態を「ページ状態履歴」シートから取得します。
 * @param {Sheet} stateSheet - 状態を保存しているシート。
 * @param {string} pageId - 検索対象のNotionページID。
 * @param {Array<string>} executionLogs - 実行ログ配列。
 * @return {object|null} 見つかった場合は { rowIndex, properties }、見つからなければ null。
 */
function getPreviousState(stateSheet, pageId, executionLogs) {
  executionLogs.push(`Searching for previous state of page ID: ${pageId}`);
  const dataRange = stateSheet.getDataRange();
  const values = dataRange.getValues();
  // ヘッダー行を除き、下から検索（最新の状態が見つかりやすいため）
  for (let i = values.length - 1; i > 0; i--) {
    if (values[i][0] === pageId) {
      // 1列目 (A列) が Page ID
      executionLogs.push(`Previous state found at row ${i + 1}.`);
      try {
        const properties = JSON.parse(values[i][1]); // 2列目 (B列) がプロパティJSON
        return { rowIndex: i + 1, properties: properties };
      } catch (e) {
        executionLogs.push(
          `ERROR: Failed to parse stored JSON for page ${pageId} at row ${i + 1}.`
        );
        return null; // パースに失敗した場合は状態なしとして扱う
      }
    }
  }
  executionLogs.push("No previous state found for this page.");
  return null;
}

/**
 * ★ 現在の状態で「ページ状態履歴」シートを更新または新規作成します。
 * @param {Sheet} stateSheet - 状態を保存しているシート。
 * @param {string} pageId - 対象のNotionページID。
 * @param {object} newProperties - 保存する新しいプロパティオブジェクト。
 * @param {number} rowIndex - 更新対象の行番号。見つかっていない場合は -1。
 * @param {Array<string>} executionLogs - 実行ログ配列。
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
    // 既存の行を更新
    stateSheet.getRange(rowIndex, 2).setValue(newPropertiesJson); // B列を更新
    executionLogs.push(`Updated state for page ${pageId} at row ${rowIndex}.`);
  } else {
    // 新しい行を追加
    stateSheet.appendRow([pageId, newPropertiesJson]);
    executionLogs.push(`Appended new state for page ${pageId}.`);
  }
}

/**
 * ★ 変更前後の情報をもとにDiscordへ通知します。
 * @param {object|null} previousInfo - 変更前の情報。
 * @param {object} currentInfo - 変更後の情報。
 * @param {string} discordWebhookUrl - Discord Webhook URL。
 * @param {Array<string>} executionLogs - 実行ログ配列。
 * @return {object} 送信結果。
 */
function sendStateChangeDiscordNotification(
  previousInfo,
  currentInfo,
  discordWebhookUrl,
  executionLogs
) {
  // (変更なし、以前のコードのまま)
  executionLogs.push("Preparing state-change Discord notification...");
  let result = { status: "未実行", error: null };

  if (!currentInfo) {
    result.status = "スキップ（現在情報なし）";
    return result;
  }

  const { kigyoMei, status, tanto, pageUrl, lastEditedTime } = currentInfo;
  let messageBody = "";

  if (previousInfo) {
    const prevStatus = previousInfo.status || "（不明）";
    const prevTanto = previousInfo.tanto || "（不明）";

    messageBody = `**企業名:** ${kigyoMei}\n`;
    if (currentInfo.status !== prevStatus) {
      messageBody += `**商談ステータス:** **\`${prevStatus}\`** → **\`${currentInfo.status}\`** に変更\n`;
    } else {
      messageBody += `**商談ステータス:** ${currentInfo.status}\n`;
    }
    if (currentInfo.tanto !== prevTanto) {
      messageBody += `**担当:** **\`${prevTanto}\`** → **\`${currentInfo.tanto}\`** に変更\n`;
    } else {
      messageBody += `**担当:** ${currentInfo.tanto}\n`;
    }
    messageBody += `**最終更新日時:** ${lastEditedTime}\n`;
  } else {
    messageBody =
      `**企業名:** ${kigyoMei}\n` +
      `**商談ステータス:** ${status}\n` +
      `**担当:** ${tanto}\n` +
      `**最終更新日時:** ${lastEditedTime}\n`;
  }

  const discordMessageContent =
    `**Notion顧客情報 更新通知** 📢\n` +
    `------------------------------------\n` +
    messageBody +
    `------------------------------------\n` +
    `詳細はこちら: ${pageUrl}`;

  if (
    discordWebhookUrl &&
    discordWebhookUrl.startsWith("https://discord.com/api/webhooks/")
  ) {
    try {
      const payload = JSON.stringify({ content: discordMessageContent });
      const options = {
        method: "post",
        contentType: "application/json",
        payload: payload,
      };
      UrlFetchApp.fetch(discordWebhookUrl, options);
      executionLogs.push("Successfully sent message to Discord.");
      result.status = "送信成功";
    } catch (discordError) {
      executionLogs.push(
        `ERROR: Sending message to Discord failed: ${discordError.toString()}`
      );
      result.status = "送信エラー";
      result.error = discordError.toString();
    }
  } else {
    executionLogs.push(
      "WARN: DISCORD_WEBHOOK_URL is not configured or invalid. Skipping notification."
    );
    result.status = "URL未設定/不正のためスキップ";
  }
  return result;
}

// --- 以下の関数はほぼ変更なし ---
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
      `GAS could not access or setup spreadsheet: ${error.message}`
    );
  }
}

function parseWebhookEvent(e, executionLogs) {
  // (変更なし、以前のコードのまま)
  executionLogs.push("Parsing webhook event...");
  let rawContents = "N/A";
  let fullEventString = "N/A";
  let notionPageData = null;

  if (e) {
    executionLogs.push("Event object 'e' received.");
    try {
      fullEventString = JSON.stringify(e);
    } catch (stringifyError) {
      executionLogs.push(
        `Could not stringify 'e' object: ${stringifyError.toString()}`
      );
      fullEventString = "Could not stringify 'e' object.";
    }

    if (e.postData && e.postData.contents) {
      rawContents = e.postData.contents;
      executionLogs.push(
        `Received e.postData.contents (length: ${rawContents.length})`
      );
      try {
        const eventData = JSON.parse(rawContents);
        if (eventData && eventData.data) {
          notionPageData = eventData.data;
          executionLogs.push("Notion Page Data parsed successfully.");
        } else {
          executionLogs.push(
            "WARN: 'data' property for Notion page is missing in parsed content."
          );
        }
      } catch (parseError) {
        executionLogs.push(
          `ERROR: Parsing receivedContents as JSON failed: ${parseError.toString()}`
        );
      }
    } else {
      executionLogs.push("WARN: e.postData or e.postData.contents is missing.");
      rawContents = "e.postData or e.postData.contents is missing";
    }
  } else {
    executionLogs.push("ERROR: Event object 'e' is undefined or null.");
    fullEventString = "Event object 'e' is undefined or null";
  }
  return { rawContents, fullEventString, notionPageData };
}

/**
 * ★ Notionページデータから「社労士連携」カラムの値も含めて抽出します。
 */
function extractNotionInfo(notionPageData, executionLogs) {
  executionLogs.push("Extracting Notion info...");
  let kigyoMei = "取得失敗";
  let status = "取得失敗";
  let tanto = "取得失敗";
  let sharoushi = null; // ★ 社労士連携カラムの値
  const pageUrl = notionPageData.url || "URL不明";
  const lastEditedTime = notionPageData.last_edited_time
    ? new Date(notionPageData.last_edited_time).toLocaleString("ja-JP")
    : "日時不明";

  const properties = notionPageData.properties;
  if (!properties) {
    executionLogs.push(
      "ERROR: 'properties' object is missing in notionPageData."
    );
    return { kigyoMei, status, tanto, sharoushi, pageUrl, lastEditedTime };
  }

  if (properties["企業名"]?.title?.[0]?.plain_text) {
    kigyoMei = properties["企業名"].title[0].plain_text;
  }
  if (properties["商談ステータス"]?.status?.name) {
    status = properties["商談ステータス"].status.name;
  }
  if (properties["担当"]?.select?.name) {
    tanto = properties["担当"].select.name;
  }

  // ★★★ 「社労士連携」カラムの抽出 ★★★
  const sharoushiProp = properties["社労士連携"];
  if (sharoushiProp && sharoushiProp.status) {
    // 型が「ステータス」なので、.status.name で値（例：「連携済み」）を取得します
    sharoushi = sharoushiProp.status.name;
  } else {
    executionLogs.push(
      "WARN: Failed to extract '社労士連携' status or property is empty."
    );
    sharoushi = sharoushiProp ? "（ステータス未設定）" : "カラムなし";
  }

  executionLogs.push(
    `Extracted => 企業名: ${kigyoMei}, ステータス: ${status}, 担当: ${tanto}, 社労士連携: ${sharoushi}`
  );
  return { kigyoMei, status, tanto, sharoushi, pageUrl, lastEditedTime };
}

/**
 * ★ 社労士連携の変更を通知するための新しい関数
 */
function sendSharoushiDiscordNotification(
  previousInfo,
  currentInfo,
  discordWebhookUrl,
  executionLogs
) {
  executionLogs.push("Preparing Sharoushi notification...");
  let result = { status: "未処理", error: null };

  if (!currentInfo) {
    result.status = "スキップ（現在情報なし）";
    return result;
  }

  const { kigyoMei, sharoushi, pageUrl } = currentInfo;
  const prevSharoushi = previousInfo
    ? previousInfo.sharoushi
    : "（変更前データなし）";

  // チェックボックスの場合を想定したメッセージ例
  let changeMessage = "";
  if (sharoushi === true) {
    changeMessage = `**\`未チェック\`** → **\`チェック済み\`** に変更されました。`;
  } else if (sharoushi === false) {
    changeMessage = `**\`チェック済み\`** → **\`未チェック\`** に変更されました。`;
  } else {
    // セレクトやステータスの場合
    changeMessage = `**\`${prevSharoushi}\`** → **\`${sharoushi}\`** に変更されました。`;
  }

  const discordMessageContent =
    `**【社労士連携】**\n` +
    `------------------------------------\n` +
    `**企業名:** ${kigyoMei}\n` +
    `**連携ステータス:** ${changeMessage}\n` +
    `------------------------------------\n` +
    `詳細はこちら: ${pageUrl}`;

  // (Discordへの送信ロジックは流用)
  if (
    discordWebhookUrl &&
    discordWebhookUrl.startsWith("https://discord.com/api/webhooks/")
  ) {
    try {
      const payload = JSON.stringify({ content: discordMessageContent });
      const options = {
        method: "post",
        contentType: "application/json",
        payload: payload,
      };
      UrlFetchApp.fetch(discordWebhookUrl, options);
      executionLogs.push(
        "Successfully sent Sharoushi notification to Discord."
      );
      result.status = "送信成功";
    } catch (discordError) {
      executionLogs.push(
        `ERROR: Sending Sharoushi notification to Discord failed: ${discordError.toString()}`
      );
      result.status = "送信エラー";
      result.error = discordError.toString();
    }
  } else {
    executionLogs.push(
      "WARN: Sharoushi DISCORD_WEBHOOK_URL is not configured. Skipping notification."
    );
    result.status = "URL未設定/不正のためスキップ";
  }
  return result;
}

function logToSheet(
  sheet,
  timestamp,
  webhookData,
  notionInfo,
  overallStatus,
  discordSendStatus,
  executionLogs
) {
  // (引数を少し整理)
  const executionLogsString = executionLogs.join("\n");
  try {
    const kigyoMei = notionInfo ? notionInfo.kigyoMei : "N/A";
    const status = notionInfo ? notionInfo.status : "N/A";
    const tanto = notionInfo ? notionInfo.tanto : "N/A";
    const rawContents = webhookData ? webhookData.rawContents : "N/A";
    const fullEventString = webhookData ? webhookData.fullEventString : "N/A";

    sheet.appendRow([
      timestamp,
      rawContents,
      fullEventString,
      kigyoMei,
      status,
      tanto,
      overallStatus,
      discordSendStatus,
      executionLogsString,
    ]);
  } catch (error) {
    Logger.log(
      `CRITICAL: Error appending final log to spreadsheet: ${error.toString()}`
    );
  }
}
