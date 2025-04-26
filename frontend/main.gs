// フロントエンド側のmain.gs
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Shopee設定")
    .addItem("事前設定", "showSettingsForm")
    .addItem("Shopee認証", "requestShopeeAuth")
    .addItem("バックエンド設定", "showBackendSettingForm")
    .addItem("バックエンド接続テスト", "testBackendConnection") // 直接テスト
    .addItem("バックエンド関数テスト", "processTestEcho") // callBackendFunction経
    .addToUi();
  ui.createMenu("Shopee出品")
    .addItem("Amazonデータ収集", "showShopSelectionDialog2")
    .addItem("出品データ生成(収集データ変更時)", "showShopSelectionDialog3")
    .addItem("Shopee出品", "showShopSelectionDialog4")
    .addItem("一括出品(Amazon→Shopee)", "showShopSelectionDialog")
    .addToUi();
  ui.createMenu("Shopee出品(バリエーション)")
    .addItem("Step1", "showShopSelectionDialog_t1_step1")
    .addItem("Step2", "showShopSelectionDialog_t1_step2")
    .addItem("Shopee出品", "showShopSelectionDialog_t1_update")
    .addToUi();
}

// バックエンドURL設定画面を表示
function showBackendSettingForm() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const template = HtmlService.createTemplateFromFile("BackendSettingForm");
  template.backendUrl = scriptProperties.getProperty("BACKEND_URL") || "";
  const htmlOutput = template.evaluate().setWidth(400).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "バックエンド設定");
}

function showShopSelectionDialog() {
  const html = HtmlService.createHtmlOutputFromFile("ShopSelectionDialog")
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "一括出品");
}

function showShopSelectionDialog2() {
  const html = HtmlService.createHtmlOutputFromFile("ShopSelectionDialog2")
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "Amazonデータ収集");
}

function showShopSelectionDialog3() {
  const html = HtmlService.createHtmlOutputFromFile("ShopSelectionDialog3")
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "出品データ生成");
}

function showShopSelectionDialog4() {
  const html = HtmlService.createHtmlOutputFromFile("ShopSelectionDialog4")
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "Shopee出品");
}

function showShopSelectionDialog_t1_step1() {
  const html = HtmlService.createHtmlOutputFromFile(
    "ShopSelectionDialog_t1_step1"
  )
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "Step1");
}

function showShopSelectionDialog_t1_step2() {
  const html = HtmlService.createHtmlOutputFromFile(
    "ShopSelectionDialog_t1_step2"
  )
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "Step2");
}

function showShopSelectionDialog_t1_update() {
  const html = HtmlService.createHtmlOutputFromFile(
    "ShopSelectionDialog_t1_update"
  )
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "Shopee出品");
}

// チェックされた行を処理
function processCheckedShops(selectedShopsIndices) {
  // 現在のスプレッドシートデータを取得
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
  const lastRow = sheet.getLastRow();
  const checkRange = sheet.getRange(2, 1, lastRow - 1);
  const asinRange = sheet.getRange(2, 2, lastRow - 1);
  const checkValues = checkRange.getValues();
  const asinValues = asinRange.getValues();

  // チェックされた行を検索
  const rowsToProcess = [];
  for (let i = 0; i < checkValues.length; i++) {
    if (checkValues[i][0] && asinValues[i][0]) {
      // チェック行のASINとデータを収集
      const rowNum = i + 2;
      const asin = asinValues[i][0];

      // 処理に必要な行データを収集
      const rowData = {
        title: sheet.getRange(rowNum, 9).getValue(),
        description: sheet.getRange(rowNum, 10).getValue(),
        price: sheet.getRange(rowNum, 4).getValue(),
        weight: sheet.getRange(rowNum, 5).getValue(),
        length: sheet.getRange(rowNum, 6).getValue(),
        width: sheet.getRange(rowNum, 7).getValue(),
        height: sheet.getRange(rowNum, 8).getValue(),
        imageIds: [],
      };

      // 画像IDを取得
      for (let j = 0; j < 8; j++) {
        const imageId = sheet.getRange(rowNum, 15 + j).getValue();
        if (imageId && imageId !== "N/A") {
          rowData.imageIds.push(imageId);
        }
      }

      rowsToProcess.push({
        rowNum: rowNum,
        asin: asin,
        data: rowData,
      });
    }
  }

  if (rowsToProcess.length === 0) {
    SpreadsheetApp.getUi().alert("処理対象の行がありません。");
    return;
  }

  try {
    // バックエンドAPIを呼び出して処理
    const result = callBackendFunction("processCheckedShops", {
      selectedShopsIndices: selectedShopsIndices,
      rowsToProcess: rowsToProcess,
    });

    if (result.status === "error") {
      SpreadsheetApp.getUi().alert("エラー: " + result.error);
      return;
    }

    // スプレッドシートを結果で更新
    result.result.processedRows.forEach((row) => {
      // チェックボックスをオフに
      sheet.getRange(row.rowNum, 1).setValue(false);

      // バックエンドから返されたデータをセット
      Object.keys(row.data).forEach((colKey) => {
        const colNum = parseInt(colKey);
        sheet.getRange(row.rowNum, colNum).setValue(row.data[colKey]);
      });
    });

    SpreadsheetApp.getUi().alert("処理が完了しました。");
    return result;
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー: " + error.message);
  }
}

// Amazonデータ収集
function processMakeData() {
  // チェック行取得ロジック
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
  const lastRow = sheet.getLastRow();
  const checkRange = sheet.getRange(2, 1, lastRow - 1);
  const asinRange = sheet.getRange(2, 2, lastRow - 1);
  const checkValues = checkRange.getValues();
  const asinValues = asinRange.getValues();

  // チェックされた行を検索
  const rowsToProcess = [];
  for (let i = 0; i < checkValues.length; i++) {
    if (checkValues[i][0] && asinValues[i][0]) {
      rowsToProcess.push({
        rowNum: i + 2,
        asin: asinValues[i][0],
      });
    }
  }

  if (rowsToProcess.length === 0) {
    SpreadsheetApp.getUi().alert("処理対象の行がありません。");
    return;
  }

  // // UI更新
  // const ui = SpreadsheetApp.getUi();
  // ui.alert("バックエンドにデータ収集を依頼しています...");

  try {
    // バックエンドAPIを呼び出して処理
    const backendUrl =
      PropertiesService.getScriptProperties().getProperty("BACKEND_URL");
    if (!backendUrl) {
      ui.alert(
        "バックエンドURLが設定されていません。「バックエンド設定」から設定してください。"
      );
      return;
    }

    // ASINリストだけを送信
    const asinList = rowsToProcess.map((row) => row.asin);

    const payload = {
      function: "processKeepaProducts",
      parameters: {
        asinList: asinList,
      },
      apiKey: "YOUR_API_KEY", // 本番環境では適切に保護する
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    // バックエンドAPI呼び出し
    const response = UrlFetchApp.fetch(backendUrl, options);

    if (response.getResponseCode() !== 200) {
      throw new Error(
        `バックエンドエラー (${response.getResponseCode()}): ${response.getContentText()}`
      );
    }

    const result = JSON.parse(response.getContentText());

    if (result.status === "error") {
      throw new Error(result.error);
    }

    // バックエンドからの結果を処理
    if (result.result && result.result.asinData) {
      // 各ASINのデータをスプレッドシートに反映
      rowsToProcess.forEach((row, index) => {
        const asinData = result.result.asinData[row.asin];
        if (asinData) {
          // チェックボックスをオフに
          sheet.getRange(row.rowNum, 1).setValue(false);

          // データを設定
          for (const colKey in asinData) {
            const colNum = parseInt(colKey);
            sheet.getRange(row.rowNum, colNum).setValue(asinData[colKey]);
          }
        }
      });

      ui.alert("データ収集が完了しました。");
    } else {
      throw new Error("バックエンドからの応答に有効なデータがありません。");
    }
  } catch (error) {
    ui.alert("エラー: " + error.message);
  }
}

function processCalcData() {
  // チェック行取得ロジック
  // ...
  const result = callBackendFunction("processCalcData", {
    rowsToProcess: rowsToProcess,
  });
  // 結果処理ロジック
  // ...
}

function processMakeDataStep1() {
  // チェック行取得ロジック
  // ...
  const result = callBackendFunction("processMakeDataStep1", {
    rowsToProcess: rowsToProcess,
  });
  // 結果処理ロジック
  // ...
}

// バックエンドとの通信ユーティリティを修正
function callBackendFunction(functionName, params) {
  // このチェックを追加
  // バックエンド接続が必要な場合のみバックエンドURLをチェック
  const needsBackend = ["refreshAllTokens"].includes(functionName);

  if (needsBackend) {
    const backendUrl =
      PropertiesService.getUserProperties().getProperty("BACKEND_URL");
    if (!backendUrl) {
      throw new Error(
        "バックエンドURLが設定されていません。「バックエンド設定」から設定してください。"
      );
    }

    const payload = {
      function: functionName,
      parameters: params,
      apiKey: "YOUR_API_KEY", // 認証用APIキー（本番では適切に管理）
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      headers: {
        "Content-Type": "application/json",
      },
    };

    try {
      const response = UrlFetchApp.fetch(backendUrl, options);
      if (response.getResponseCode() !== 200) {
        throw new Error(
          `バックエンドエラー (${response.getResponseCode()}): ${response.getContentText()}`
        );
      }
      return JSON.parse(response.getContentText());
    } catch (error) {
      console.error("バックエンド通信エラー:", error);
      throw new Error(`バックエンド通信エラー: ${error.message}`);
    }
  } else {
    // バックエンド接続が不要な関数はフロントエンドで処理
    return { status: "success", result: {} };
  }
}

// バックエンドURLを保存する関数
function saveBackendUrl(data) {
  try {
    if (!data || !data.backendUrl) {
      return "エラー: バックエンドURLが入力されていません。";
    }

    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty("BACKEND_URL", data.backendUrl);

    // 保存確認のためにプロパティを読み込み
    const savedUrl = scriptProperties.getProperty("BACKEND_URL");

    if (savedUrl === data.backendUrl) {
      return "バックエンドURLが保存されました: " + savedUrl;
    } else {
      return "エラー: バックエンドURLの保存に失敗しました。";
    }
  } catch (error) {
    console.error("バックエンドURL保存エラー:", error);
    return "エラー: " + error.message;
  }
}

// スクリプトプロパティの内容を確認する関数（デバッグ用）
function checkAllProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const properties = scriptProperties.getProperties();

  let message = "現在のスクリプトプロパティ：\n\n";
  for (const key in properties) {
    const value = properties[key];
    if (key.includes("KEY") || key.includes("ID") || key.includes("SECRET")) {
      message += `${key}: ********\n`;
    } else {
      message += `${key}: ${value}\n`;
    }
  }

  SpreadsheetApp.getUi().alert(message);
}

// バックエンド接続テスト関数
function testBackendConnection() {
  const ui = SpreadsheetApp.getUi();

  try {
    // バックエンドURLの取得
    const scriptProperties = PropertiesService.getScriptProperties();
    const backendUrl = scriptProperties.getProperty("BACKEND_URL");

    if (!backendUrl) {
      ui.alert(
        "バックエンドURLが設定されていません。「バックエンド設定」から設定してください。"
      );
      return;
    }

    // テスト用のペイロードを作成
    const payload = {
      function: "testEcho",
      parameters: {
        message: "テスト通信",
        timestamp: new Date().toISOString(),
      },
    };

    // APIリクエストのオプション
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    ui.alert("バックエンドにテスト通信を送信します: " + backendUrl);

    // バックエンドAPI呼び出し
    const response = UrlFetchApp.fetch(backendUrl, options);

    // レスポンスの詳細をログに出力
    console.log(`応答コード: ${response.getResponseCode()}`);
    console.log(`応答ヘッダー: ${JSON.stringify(response.getAllHeaders())}`);
    console.log(`応答内容: ${response.getContentText()}`);

    // 応答コードを確認
    if (response.getResponseCode() !== 200) {
      ui.alert(
        `エラー: HTTP ${response.getResponseCode()}\n${response
          .getContentText()
          .substring(0, 300)}`
      );
      return;
    }

    // JSONとして解析
    try {
      const result = JSON.parse(response.getContentText());

      // 結果を表示
      if (result.status === "success") {
        ui.alert(
          "テスト成功!\n\n" +
            `受信データ: ${JSON.stringify(result.result, null, 2)}`
        );
      } else {
        ui.alert(
          "バックエンドからエラーが返されました:\n\n" +
            `エラー: ${result.error || "不明なエラー"}`
        );
      }
    } catch (jsonError) {
      ui.alert(
        "JSONデータの解析に失敗しました:\n\n" +
          `エラー: ${jsonError.message}\n\n` +
          `受信データ: ${response.getContentText().substring(0, 300)}`
      );
    }
  } catch (error) {
    console.error("テスト通信エラー:", error);
    ui.alert("通信エラー: " + error.message);
  }
}

// callBackendFunction関数を使用する場合のテスト関数
function processTestEcho() {
  const ui = SpreadsheetApp.getUi();

  try {
    // テスト用パラメータ
    const testParams = {
      message: "テスト通信",
      timestamp: new Date().toISOString(),
    };

    // バックエンド関数を呼び出す
    const result = callBackendFunction("testEcho", testParams);

    // 結果の表示
    if (result.status === "success") {
      ui.alert(
        "テスト成功!\n\n" +
          `受信データ: ${JSON.stringify(result.result, null, 2)}`
      );
    } else {
      ui.alert(
        "バックエンドからエラーが返されました:\n\n" +
          `エラー: ${result.error || "不明なエラー"}`
      );
    }
  } catch (error) {
    console.error("テスト実行エラー:", error);
    ui.alert("エラー: " + error.message);
  }
}
