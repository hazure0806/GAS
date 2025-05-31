// ------------------------------------------------------------------------------------
// スクリプトプロパティから設定値を読み込む
// ------------------------------------------------------------------------------------
const PROP_KEY_SPREADSHEET_ID = "SPREADSHEET_ID_SECRET";
const PROP_KEY_SHEET_NAME = "SHEET_NAME_VALUE";
const PROP_KEY_DISCORD_URL = "DISCORD_WEBHOOK_URL_SECRET";

/**
 * スクリプトプロパティから指定されたキーの値を取得します。
 * 見つからない場合や値がnullの場合は、defaultValueを返します。
 * defaultValueもnullで値が見つからない場合はnullを返しますが、警告ログを出力します。
 * @param {string} key 取得するプロパティのキー。
 * @param {any} [defaultValue=null] プロパティが見つからない場合に返すデフォルト値。
 * @return {string|null} プロパティの値、またはデフォルト値、またはnull。
 */
function getScriptPropertyValue(key, defaultValue = null) {
  const properties = PropertiesService.getScriptProperties();
  const value = properties.getProperty(key);
  if (value === null && defaultValue === null) {
    Logger.log(`警告: スクリプトプロパティ "${key}" が設定されていません。`);
  }
  return value !== null ? value : defaultValue;
}

const SPREADSHEET_ID = getScriptPropertyValue(PROP_KEY_SPREADSHEET_ID);
const SHEET_NAME = getScriptPropertyValue(
  PROP_KEY_SHEET_NAME,
  "NotionWebhookLog"
);
const DISCORD_WEBHOOK_URL = getScriptPropertyValue(PROP_KEY_DISCORD_URL);

const HEADERS = [
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
/**
 * Handles HTTP POST requests to the web app.
 * Expected to be called by Notion's webhook automation.
 * @param {GoogleAppsScript.Events.DoPost} e The event parameter for a POST request.
 * @return {GoogleAppsScript.Content.TextOutput} A TextOutput response.
 */
function doPost(e) {
  let executionLogs = [];
  executionLogs.push("--- doPost Execution Start ---");

  let overallStatus = "成功";
  let discordSendStatus = "未処理";
  let errorForSheet = "";

  let sheet;
  let timestamp = new Date();
  let webhookData = {
    rawContents: "N/A",
    fullEventString: "N/A",
    notionPageData: null,
  };
  let extractedNotionInfo = null;

  try {
    if (!SPREADSHEET_ID) {
      throw new Error(
        `Script Property "${PROP_KEY_SPREADSHEET_ID}" is not set or empty.`
      );
    }
    sheet = getOrCreateSheet(SPREADSHEET_ID, SHEET_NAME, HEADERS);
    webhookData = parseWebhookEvent(e, executionLogs);

    if (!webhookData.notionPageData) {
      executionLogs.push("Notion page data is not available.");
      overallStatus = "Notionデータなし";
      discordSendStatus = "スキップ（Notionデータなし）";
      logToSheet(
        sheet,
        timestamp,
        webhookData.rawContents,
        webhookData.fullEventString,
        null,
        overallStatus,
        discordSendStatus,
        executionLogs
      );
      return ContentService.createTextOutput(
        JSON.stringify({
          status: "error",
          message: "Notion page data not found.",
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    extractedNotionInfo = extractNotionInfo(
      webhookData.notionPageData,
      executionLogs
    );
    const discordResult = sendDiscordNotificationAndGetStatus(
      extractedNotionInfo,
      DISCORD_WEBHOOK_URL,
      executionLogs
    );
    discordSendStatus = discordResult.status;

    if (discordResult.error) {
      errorForSheet = "Discord送信エラー: " + discordResult.error;
      overallStatus = overallStatus === "成功" ? "一部エラー" : overallStatus;
    }
  } catch (error) {
    executionLogs.push(
      `Critical Error in doPost: ${error.toString()}${error.stack ? "\nStack: " + error.stack : ""}`
    );
    overallStatus = "致命的エラー";
    errorForSheet = "全体処理エラー: " + error.toString();
  } finally {
    if (sheet) {
      logToSheet(
        sheet,
        timestamp,
        webhookData.rawContents,
        webhookData.fullEventString,
        extractedNotionInfo,
        overallStatus,
        discordSendStatus,
        executionLogs
      );
    } else {
      Logger.log(
        `CRITICAL: Sheet object was not available. Logs: ${executionLogs.join("\n")}`
      );
    }
  }

  executionLogs.push("--- doPost Execution End ---");
  Logger.log(executionLogs.join("\n"));

  return ContentService.createTextOutput(
    JSON.stringify({
      status: overallStatus,
      message: errorForSheet || "Webhook processed.",
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------------------------------
// ヘルパー関数群
// ------------------------------------------------------------------------------------

/**
 * 指定されたIDのスプレッドシートを開き、指定された名前のシートを取得または作成します。
 * @param {string} spreadsheetId 操作対象のスプレッドシートID。
 * @param {string} sheetName 操作対象のシート名。
 * @param {Array<string>} headers シートが新規作成された場合に追加するヘッダー行。
 * @return {GoogleAppsScript.Spreadsheet.Sheet} Apps ScriptのSheetオブジェクト。
 * @throws {Error} スプレッドシートのアクセスや設定に失敗した場合
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
      `GAS could not access or setup spreadsheet: ${error.message}`
    );
  }
}

/**
 * Webhookイベントオブジェクト(e)を解析し、主要なデータを抽出します。
 * @param {object} e doPostから渡されるイベントオブジェクト。
 * @param {Array<string>} executionLogs 実行ログを格納する配列。
 * @return {object} 解析されたデータを含むオブジェクト { rawContents, fullEventString, notionPageData }。
 */
function parseWebhookEvent(e, executionLogs) {
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
 * Notionのページデータオブジェクトから指定された情報を抽出します。
 * @param {object} notionPageData Notionのページデータ。
 * @param {Array<string>} executionLogs 実行ログを格納する配列。
 * @return {object} 抽出された情報 { kigyoMei, status, tanto, pageUrl, lastEditedTime }。
 */
function extractNotionInfo(notionPageData, executionLogs) {
  executionLogs.push("Extracting Notion info...");
  let kigyoMei = "取得失敗";
  let status = "取得失敗";
  let tanto = "取得失敗";
  const pageUrl = notionPageData.url || "URL不明";
  const lastEditedTime = notionPageData.last_edited_time
    ? new Date(notionPageData.last_edited_time).toLocaleString("ja-JP")
    : "日時不明";

  if (notionPageData.properties?.["企業名"]?.title?.[0]?.plain_text) {
    kigyoMei = notionPageData.properties["企業名"].title[0].plain_text;
  } else {
    executionLogs.push("WARN: Failed to extract '企業名'.");
  }
  if (notionPageData.properties?.["商談ステータス"]?.status?.name) {
    status = notionPageData.properties["商談ステータス"].status.name;
  } else {
    executionLogs.push("WARN: Failed to extract '商談ステータス'.");
  }
  if (notionPageData.properties?.["担当"]?.select?.name) {
    tanto = notionPageData.properties["担当"].select.name;
  } else {
    executionLogs.push("WARN: Failed to extract '担当'.");
  }

  executionLogs.push(
    `Extracted => 企業名: ${kigyoMei}, ステータス: ${status}, 担当: ${tanto}`
  );
  return { kigyoMei, status, tanto, pageUrl, lastEditedTime };
}

/**
 * 抽出されたNotion情報からDiscordへの通知メッセージを作成し送信し、結果を返します。
 * @param {object} notionInfo extractNotionInfoから返されるオブジェクト。
 * @param {string} discordWebhookUrl 送信先のDiscord Webhook URL。
 * @param {Array<string>} executionLogs 実行ログを格納する配列。
 * @return {object} 送信結果 { status: string, error: string|null }。
 */
function sendDiscordNotificationAndGetStatus(
  notionInfo,
  discordWebhookUrl,
  executionLogs
) {
  executionLogs.push("Preparing Discord notification...");
  let result = { status: "未実行", error: null };

  if (!notionInfo) {
    executionLogs.push(
      "Notion info is not available, skipping Discord notification."
    );
    result.status = "スキップ（Notion情報なし）";
    return result;
  }
  const { kigyoMei, status, tanto, pageUrl, lastEditedTime } = notionInfo;
  const discordMessageContent =
    `**Notion顧客情報 更新通知** 📢\n` +
    `------------------------------------\n` +
    `**企業名:** ${kigyoMei}\n` +
    `**商談ステータス:** ${status}\n` +
    `**担当:** ${tanto}\n` +
    `**最終更新日時:** ${lastEditedTime}\n` +
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

/**
 * 指定されたシートにデータを追記します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 書き込み対象のApps ScriptのSheetオブジェクト。
 * @param {Date} timestamp 受信日時。
 * @param {string} rawContents Webhookの生データコンテンツ。
 * @param {string} fullEventString イベントオブジェクト全体の文字列。
 * @param {object | null} notionInfo extractNotionInfoから返されるオブジェクト、またはデータがない場合はnull。
 * @param {string} overallStatus 全体処理のステータス。
 * @param {string} discordSendStatus Discord通知の結果ステータス。
 * @param {Array<string>} executionLogs 実行ログ（文字列の配列）。
 */
function logToSheet(
  sheet,
  timestamp,
  rawContents,
  fullEventString,
  notionInfo,
  overallStatus,
  discordSendStatus,
  executionLogs
) {
  const executionLogsString = executionLogs.join("\n");
  try {
    const kigyoMei = notionInfo ? notionInfo.kigyoMei : "N/A";
    const status = notionInfo ? notionInfo.status : "N/A";
    const tanto = notionInfo ? notionInfo.tanto : "N/A";

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
    Logger.log(
      `Data attempted to log: Timestamp: ${timestamp}, Overall: ${overallStatus}, Discord: ${discordSendStatus}, Logs: ${executionLogsString}`
    );
  }
}
