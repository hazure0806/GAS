/**
 * @fileoverview 領収書PDFをGoogle Driveから取得し、
 * スプレッドシートのリストに基づいてGmailの下書きを自動作成するスクリプトです。
 * @version 1.2.0 (複数行処理対応・ロギング強化版)
 */

// ------------------------------------------------------------------------------------
// 設定 (CONFIG)
// ------------------------------------------------------------------------------------
function getConfig() {
  const properties = PropertiesService.getScriptProperties();
  const config = {
    PDF_PARENT_FOLDER_ID: properties.getProperty("PDF_PARENT_FOLDER_ID"),
  };
  if (!config.PDF_PARENT_FOLDER_ID) {
    throw new Error(
      "スクリプトプロパティ「PDF_PARENT_FOLDER_ID」が設定されていません。"
    );
  }
  return config;
}

// ------------------------------------------------------------------------------------
// メニュー追加機能
// ------------------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("領収書メール作成")
    .addItem("選択した行の下書きを作成", "createDraftsForSelectedRows")
    .addToUi();
}

// ------------------------------------------------------------------------------------
// メイン処理
// ------------------------------------------------------------------------------------
function createDraftsForSelectedRows() {
  const ui = SpreadsheetApp.getUi();
  Logger.log("--- 関数 createDraftsForSelectedRows 開始 ---");
  try {
    const CONFIG = getConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    if (!sheet.getName().startsWith("送信リスト_")) {
      ui.alert("「送信リスト_YYYY」という名前のシートで実行してください。");
      return;
    }

    const settingsSheet = ss.getSheetByName("メール設定");
    if (!settingsSheet)
      throw new Error("「メール設定」シートが見つかりません。");
    const subjectTemplate = settingsSheet.getRange("B1").getValue();
    const bodyTemplate = settingsSheet.getRange("B2").getValue();

    const selection = sheet.getSelection();
    const selectedRanges = selection.getActiveRangeList().getRanges();
    if (selectedRanges.length === 0) {
      ui.alert("下書きを作成したい行を選択してください。");
      return;
    }

    Logger.log(`選択された範囲の数: ${selectedRanges.length}`);
    for (let i = 0; i < selectedRanges.length; i++) {
      Logger.log(`範囲 #${i + 1}: ${selectedRanges[i].getA1Notation()}`);
    }

    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const colIndex = getColumnIndexes(headers);

    // 領収書フォルダ
    const parentFolder = DriveApp.getFolderById(CONFIG.PDF_PARENT_FOLDER_ID);
    // 成功件数カウント
    let successCount = 0;
    // 月フォルダのキャッシュ
    const monthFolderCache = {};

    // 選択された各「範囲」に対してループ
    for (const range of selectedRanges) {
      const startRow = range.getRow();
      const numRows = range.getNumRows();

      Logger.log(
        `処理中の範囲: ${range.getA1Notation()}, 開始行: ${startRow}, 行数: ${numRows}`
      );

      // 範囲内の各「行」に対してループ
      for (let i = 0; i < numRows; i++) {
        const currentRow = startRow + i;
        if (currentRow === 1) {
          Logger.log(`行 ${currentRow}: ヘッダー行のためスキップします。`);
          continue; // ヘッダー行はスキップ
        }

        Logger.log(`--- ▼ 行 ${currentRow} の処理を開始します ▼ ---`);

        const statusCell = sheet.getRange(
          currentRow,
          colIndex["処理ステータス"] + 1
        );
        if (statusCell.getValue() === "作成済み") {
          Logger.log(
            `行 ${currentRow}: ステータスが「作成済み」のためスキップします。`
          );
          continue;
        }

        const rowValues = sheet
          .getRange(currentRow, 1, 1, headers.length)
          .getValues()[0];
        const clientName = rowValues[colIndex["契約主名"]];
        Logger.log(
          `行 ${currentRow}: 契約主名「${clientName}」のデータを取得しました。`
        );

        const email = rowValues[colIndex["メールアドレス"]];
        const patientName = rowValues[colIndex["患者様名"]];
        const pdfFileName = rowValues[colIndex["PDFファイル名"]];
        const targetMonthFolder = rowValues[colIndex["対象月 (フォルダ名)"]];

        if (!email || !pdfFileName || !targetMonthFolder) {
          statusCell.setValue(
            "エラー: 必須情報（アドレス, ファイル名, 対象月）が不足しています。"
          );
          Logger.log(
            `行 ${currentRow}: 必須情報が不足しているためスキップします。`
          );
          continue;
        }

        try {
          let monthFolder = monthFolderCache[targetMonthFolder];
          if (!monthFolder) {
            Logger.log(`キャッシュにないため、月フォルダ「${targetMonthFolder}」を検索します。`);
            const monthFolders = parentFolder.getFoldersByName(targetMonthFolder);
            if (!monthFolders.hasNext()) {
              throw new Error(`月フォルダ「${targetMonthFolder}」が見つかりません。`);
            }
            monthFolder = monthFolders.next();
            monthFolderCache[targetMonthFolder] = monthFolder; // 見つけたら記憶
          }
          
          const files = monthFolder.getFilesByName(pdfFileName);
          if (!files.hasNext()) {
            throw new Error(`PDFファイル「${pdfFileName}」が見つかりません。`);
          }
          const pdfFile = files.next();

          const subject = subjectTemplate.replace(/{契約主名}/g, clientName);
          const body = bodyTemplate
            .replace(/{契約主名}/g, clientName)
            .replace(/{患者様名}/g, patientName);

          GmailApp.createDraft(email, subject, body, {
            attachments: [pdfFile.getAs("application/pdf")],
          });
          Logger.log(`行 ${currentRow}: Gmailの下書きを作成しました。`);

          statusCell.setValue("作成済み");
          successCount++;
        } catch (e) {
          Logger.log(`行 ${currentRow}: エラー発生 - ${e.message}`);
          statusCell.setValue(`エラー: ${e.message}`);
        }
        Logger.log(`--- ▲ 行 ${currentRow} の処理を終了します ▲ ---`);
      }
    }

    ui.alert(
      `${successCount}件のメール下書きを作成しました。Gmailをご確認ください。`
    );
    Logger.log("--- 関数 createDraftsForSelectedRows 正常終了 ---");
  } catch (error) {
    Logger.log(
      `!!! 致命的なエラーが発生しました: ${error.message}\n${error.stack}`
    );
    ui.alert(`処理中にエラーが発生しました。\n詳細: ${error.message}`);
  }
}

// ------------------------------------------------------------------------------------
// ヘルパー関数
// ------------------------------------------------------------------------------------
function getColumnIndexes(headers) {
  const indexes = {};
  headers.forEach((header, i) => {
    if (header) {
      indexes[header.trim()] = i;
    }
  });
  return indexes;
}

function findPdfFile(parentFolder, monthFolderName, pdfFileName) {
  const monthFolders = parentFolder.getFoldersByName(monthFolderName);
  if (!monthFolders.hasNext()) {
    throw new Error(`月フォルダ「${monthFolderName}」が見つかりません。`);
  }
  const monthFolder = monthFolders.next();
  const files = monthFolder.getFilesByName(pdfFileName);
  if (!files.hasNext()) {
    throw new Error(`PDFファイル「${pdfFileName}」が見つかりません。`);
  }
  return files.next();
}
