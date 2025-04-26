/**
 * 主なワークフローを処理（データ収集、出品データ生成、履歴保存）。
 */
function processCalcData() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");

  const lastRow = sheet.getLastRow();
  const checkRange = sheet.getRange(2, 1, lastRow - 1); // A列（チェックボックス）の範囲
  const asinRange = sheet.getRange(2, 2, lastRow - 1); // B列（ASIN）の範囲

  const checkValues = checkRange.getValues(); // A列のチェックボックス状態
  const asinValues = asinRange.getValues(); // B列のASIN

  const rowsToProcess = [];

  // チェックされた行をリストに追加
  for (let i = 0; i < checkValues.length; i++) {
    const isChecked = checkValues[i][0]; // チェックボックス
    const asin = asinValues[i][0]; // ASIN

    if (isChecked && asin) {
      rowsToProcess.push(i + 2); // 行番号を保存 (スプレッドシート上では2行目から)
    }
  }

  // 処理対象が無ければ終了
  if (rowsToProcess.length === 0) {
    Logger.log("処理対象の行がありません。");
    return;
  }

  // 各行に対して1回だけ実行する処理
  rowsToProcess.forEach(function (row) {
    // 各行ごとに1回だけ実行する関数
    processShopeeListingRow(row); // 出品データ生成
  });
}

function processAddItems(selectedShopsIndices) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const shopsData = JSON.parse(scriptProperties.getProperty("SHOPS_DATA"));

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");

  const lastRow = sheet.getLastRow();
  const checkRange = sheet.getRange(2, 1, lastRow - 1); // A列（チェックボックス）の範囲
  const asinRange = sheet.getRange(2, 2, lastRow - 1); // B列（ASIN）の範囲

  const checkValues = checkRange.getValues(); // A列のチェックボックス状態
  const asinValues = asinRange.getValues(); // B列のASIN

  const rowsToProcess = [];

  // チェックされた行をリストに追加
  for (let i = 0; i < checkValues.length; i++) {
    const isChecked = checkValues[i][0]; // チェックボックス
    const asin = asinValues[i][0]; // ASIN

    if (isChecked && asin) {
      rowsToProcess.push(i + 2); // 行番号を保存 (スプレッドシート上では2行目から)
    }
  }

  // 処理対象が無ければ終了
  if (rowsToProcess.length === 0) {
    Logger.log("処理対象の行がありません。");
    return;
  }

  // 各行に対して1回だけ実行する処理
  rowsToProcess.forEach(function (row) {
    // 各行ごとに1回だけ実行する関数
    uploadImageToShopee(row); // 画像のアップロード

    // 選択された店舗数分実行する処理
    selectedShopsIndices.forEach(function (index) {
      const shop = shopsData[index]; // 選択された店舗データ

      // 店舗数分実行する関数
      processAddShopeeItem(
        row,
        shop.shop_id,
        shop.access_token,
        shop.region,
        shop.shop_name
      ); // 店舗ごとの出品処理
    });

    // 全店舗の処理が完了したタイミングで、チェックボックスをオフにし、処理日付を設定
    sheet.getRange(row, 1).setValue(false); // A列のチェックボックスをオフ
    const currentDate = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy/MM/dd"
    );
    sheet.getRange(row, 62).setValue(currentDate); // BJ列に処理日付を設定
  });
}

/**
 * Keepa APIの残りトークン数を取得
 * @param {string} key - Keepa APIキー
 * @return {number} - 残りトークン数
 */
function processCheckedShops(selectedShopsIndices) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const shopsData = JSON.parse(scriptProperties.getProperty("SHOPS_DATA"));

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");

  const lastRow = sheet.getLastRow();
  const checkRange = sheet.getRange(2, 1, lastRow - 1); // A列（チェックボックス）の範囲
  const asinRange = sheet.getRange(2, 2, lastRow - 1); // B列（ASIN）の範囲

  const checkValues = checkRange.getValues(); // A列のチェックボックス状態
  const asinValues = asinRange.getValues(); // B列のASIN

  const rowsToProcess = [];

  // チェックされた行をリストに追加
  for (let i = 0; i < checkValues.length; i++) {
    const isChecked = checkValues[i][0]; // チェックボックス
    const asin = asinValues[i][0]; // ASIN

    if (isChecked && asin) {
      rowsToProcess.push(i + 2); // 行番号を保存 (スプレッドシート上では2行目から)
    }
  }

  // 処理対象が無ければ終了
  if (rowsToProcess.length === 0) {
    Logger.log("処理対象の行がありません。");
    return;
  }

  // 各行に対して1回だけ実行する処理
  rowsToProcess.forEach(function (row) {
    // 各行ごとに1回だけ実行する関数
    processKeepaProductRow(row); // Amazonデータ取得
    processShopeeListingRow(row); // 出品データ生成
    uploadImageToShopee(row); // 画像のアップロード

    // 選択された店舗数分実行する処理
    selectedShopsIndices.forEach(function (index) {
      const shop = shopsData[index]; // 選択された店舗データ

      // 店舗数分実行する関数
      processAddShopeeItem(
        row,
        shop.shop_id,
        shop.access_token,
        shop.region,
        shop.shop_name
      ); // 店舗ごとの出品処理
    });

    // 全店舗の処理が完了したタイミングで、チェックボックスをオフにし、処理日付を設定
    sheet.getRange(row, 1).setValue(false); // A列のチェックボックスをオフ
    const currentDate = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy/MM/dd"
    );
    sheet.getRange(row, 62).setValue(currentDate); // BJ列に処理日付を設定
  });
}

function testKeepaProductRow() {
  const row = 23;
  const response = processKeepaProductRow(row);
}

/**
 * Keepa APIを呼び出してAmazonの商品情報
 */
function processKeepaProductRow(row) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODUCTS);
  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);

  // 基本設定シートからAPIキーとフォルダIDを取得
  const apiKey = settingsSheet.getRange("C4").getValue();
  // const apiKey = PropertiesService.getScriptProperties().getProperty('KEEPA_API_KEY');
  const folderId = settingsSheet.getRange("C5").getValue();

  if (!sheet || !apiKey || !folderId) {
    Logger.log("出品一覧シート、APIキー、またはフォルダIDが見つかりません。");
    return;
  }

  const asin = sheet.getRange(row, COL_ASIN).getValue(); // メインASIN

  if (!asin) {
    Logger.log(`行${row}にASINがありません。`);
    return;
  }

  // 処理開始前にC列～AD列をクリア
  sheet.getRange(row, COL_DATA_DATE, 1, VARIATION_DATAS).clearContent(); // C列からAD列をクリア

  let productData = postKeepaProducts(apiKey, asin);
  if (!productData) {
    Logger.log("APIトークンが無くなりました。");
    return;
  }

  // productDataがJSON文字列なので、オブジェクトに変換
  productData = JSON.parse(productData);

  // データが正しく取得できたか確認
  Logger.log(JSON.stringify(productData)); // ログでデータ確認用

  // 画像をダウンロードしてGoogle Driveに保存
  const imageIds = [];
  if (productData.image) {
    const images = productData.image.split(",");
    Logger.log(images);
    for (let j = 0; j < images.length && j < 8; j++) {
      const fileId = downloadImageAndSave(images[j], folderId);
      imageIds.push(fileId || "N/A");
    }
  }

  // シートにプロット
  const currentDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy/MM/dd"
  );
  sheet.getRange(row, 3).setValue(currentDate); // C列：データ取得日
  sheet.getRange(row, 4).setValue(productData.price || "N/A"); // D列：新品最安値
  sheet
    .getRange(row, 5)
    .setValue(productData.weight ? productData.weight / 1000 : "N/A"); // E列：重量 (kg)
  sheet
    .getRange(row, 6)
    .setValue(productData.length ? productData.length / 10 : "N/A"); // F列：長さ (cm)
  sheet
    .getRange(row, 7)
    .setValue(productData.width ? productData.width / 10 : "N/A"); // G列：幅 (cm)
  sheet
    .getRange(row, 8)
    .setValue(productData.height ? productData.height / 10 : "N/A"); // H列：高さ (cm)
  sheet.getRange(row, 9).setValue(productData.title || "N/A"); // I列：商品名

  // featuresとdescriptionの結合（改行を減らす）
  const features = productData.features
    ? productData.features.join("\n").replace(/\n+/g, "\n")
    : "";
  const description = productData.description
    ? productData.description.replace(/\n+/g, "\n")
    : "";
  sheet.getRange(row, 10).setValue(features + "\n" + description || "N/A"); // J列：商品説明

  // 翻訳用の数式をセット
  sheet.getRange(row, 11).setFormula(`=GOOGLETRANSLATE(I${row}, "ja", "en")`); // K列：商品名 (英)
  sheet.getRange(row, 12).setFormula(`=GOOGLETRANSLATE(J${row}, "ja", "en")`); // L列：商品説明 (英)
  sheet
    .getRange(row, 13)
    .setFormula(`=GOOGLETRANSLATE(I${row}, "ja", "zh-CN")`); // M列：商品名 (中)
  sheet
    .getRange(row, 14)
    .setFormula(`=GOOGLETRANSLATE(J${row}, "ja", "zh-CN")`); // N列：商品説明 (中)

  // 画像ファイルIDをプロット、N/Aの場合はイメージ欄もN/Aにする
  for (let j = 0; j < 8; j++) {
    const imageId = imageIds[j] || "N/A";
    sheet.getRange(row, 15 + j).setValue(imageId); // O～V列：画像①～⑧
    var imageFormula = "N/A";
    if (j === 0) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&O${row})`;
    } else if (j === 1) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&P${row})`;
    } else if (j === 2) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&Q${row})`;
    } else if (j === 3) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&R${row})`;
    } else if (j === 4) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&S${row})`;
    } else if (j === 5) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&T${row})`;
    } else if (j === 6) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&U${row})`;
    } else if (j === 7) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&V${row})`;
    }
    sheet.getRange(row, 23 + j).setValue(imageFormula); // W～AD列：イメージ①～⑧
  }
}

/**
 * 出品情報を作成する
 */
function processShopeeListingRow(row) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("基本設定");
  const additionalSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("追加設定");

  const currentDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd"
  );

  // 各シートから必要な情報を取得
  const initWeight = parseFloat(settingsSheet.getRange("C7").getValue()); // 初期重量（未取得時）
  const addWeight = parseFloat(settingsSheet.getRange("C8").getValue()); // 重量加算
  const initStock = parseFloat(settingsSheet.getRange("C9").getValue()); // 初期在庫

  const rateMap = {}; // 各国レート
  rateMap["SG"] = parseFloat(settingsSheet.getRange("F22").getValue()); // SGレート
  rateMap["MY"] = parseFloat(settingsSheet.getRange("F23").getValue()); // MYレート
  rateMap["PH"] = parseFloat(settingsSheet.getRange("F24").getValue()); // PHレート
  rateMap["TH"] = parseFloat(settingsSheet.getRange("F25").getValue()); // THレート
  rateMap["TW"] = parseFloat(settingsSheet.getRange("F26").getValue()); // TWレート
  rateMap["VN"] = parseFloat(settingsSheet.getRange("F27").getValue()); // VNレート
  rateMap["BR"] = parseFloat(settingsSheet.getRange("F23").getValue()); // VNレート

  const asin = sheet.getRange(row, 2).getValue(); // ASIN
  const productWeight =
    sheet.getRange(row, 5).getValue() === "N/A"
      ? initWeight
      : parseFloat(sheet.getRange(row, 5).getValue()) + addWeight; // 重量 (kg)

  const priceMap = priceCalc(row);

  // プロット前にクリア
  sheet.getRange(row, 31, 1, 22).clearContent();

  // 英語の商品名を255文字以内にする
  const item_name_en_bef = additionalSheet.getRange("D6").getValue();
  const item_name_en_aft = additionalSheet.getRange("D7").getValue();
  var item_name_en = sheet.getRange(row, 11).getValue();
  if (item_name_en_bef) {
    item_name_en = item_name_en_bef + " " + item_name_en;
  }
  if (item_name_en_aft) {
    item_name_en = item_name_en + " " + item_name_en_aft;
  }
  const item_name_en_255 =
    item_name_en.length > 255 ? item_name_en.substring(0, 255) : item_name_en;

  // 英語の商品説明を3000文字以内にする
  const description_en_bef = additionalSheet.getRange("D9").getValue();
  const description_en_aft = additionalSheet.getRange("D10").getValue();
  var description_en = sheet.getRange(row, 12).getValue();
  if (description_en_bef) {
    description_en = description_en_bef + "\n" + description_en;
  }
  if (description_en_aft) {
    description_en = description_en + "\n" + description_en_aft;
  }
  const description_en_3000 =
    description_en.length > 3000
      ? description_en.substring(0, 3000)
      : description_en;

  // 中国語の商品名を60文字以内にする
  const item_name_ch_bef = additionalSheet.getRange("E6").getValue();
  const item_name_ch_aft = additionalSheet.getRange("E7").getValue();
  var item_name_ch = sheet.getRange(row, 13).getValue();
  if (item_name_ch_bef) {
    item_name_ch = item_name_ch_bef + " " + item_name_ch;
  }
  if (item_name_ch_aft) {
    item_name_ch = item_name_ch + " " + item_name_ch_aft;
  }
  const item_name_ch_60 =
    item_name_ch.length > 60 ? item_name_ch.substring(0, 60) : item_name_ch;

  // 中国語の商品説明を3000文字以内にする
  const description_ch_bef = additionalSheet.getRange("E9").getValue();
  const description_ch_aft = additionalSheet.getRange("E10").getValue();
  var description_ch = sheet.getRange(row, 14).getValue();
  if (description_ch_bef) {
    description_ch = description_ch_bef + "\n" + description_ch;
  }
  if (description_ch_aft) {
    description_ch = description_ch + "\n" + description_ch_aft;
  }
  const description_ch_3000 =
    description_ch.length > 3000
      ? description_ch.substring(0, 3000)
      : description_ch;

  // シートにプロット
  sheet.getRange(row, 31).setValue(`${asin}-${currentDate}`); // AE列：SKU
  sheet.getRange(row, 32).setValue(priceMap["SG"]);
  sheet.getRange(row, 33).setValue(priceMap["MY"]);
  sheet.getRange(row, 34).setValue(priceMap["PH"]);
  sheet.getRange(row, 35).setValue(priceMap["TH"]);
  sheet.getRange(row, 36).setValue(priceMap["TW"]);
  sheet.getRange(row, 37).setValue(priceMap["VN"]);
  sheet.getRange(row, 38).setValue(priceMap["BR"]);
  sheet.getRange(row, 39).setValue(Math.round(priceMap["SG"] * rateMap["SG"]));
  sheet.getRange(row, 40).setValue(Math.round(priceMap["MY"] * rateMap["MY"]));
  sheet.getRange(row, 41).setValue(Math.round(priceMap["PH"] * rateMap["PH"]));
  sheet.getRange(row, 42).setValue(Math.round(priceMap["TH"] * rateMap["TH"]));
  sheet.getRange(row, 43).setValue(Math.round(priceMap["TW"] * rateMap["TW"]));
  sheet.getRange(row, 44).setValue(Math.round(priceMap["VN"] * rateMap["VN"]));
  sheet.getRange(row, 45).setValue(item_name_en_255); // 255文字で切る
  sheet.getRange(row, 46).setValue(description_en_3000);
  sheet.getRange(row, 47).setValue(item_name_ch_60); // 60文字で切る
  sheet.getRange(row, 48).setValue(description_ch_3000);
  sheet.getRange(row, 49).setValue(productWeight);
  sheet
    .getRange(row, 50)
    .setValue(
      sheet.getRange(row, 6).getValue() === "N/A"
        ? settingsSheet.getRange("I7").getValue()
        : sheet.getRange(row, 6).getValue()
    );
  sheet
    .getRange(row, 51)
    .setValue(
      sheet.getRange(row, 7).getValue() === "N/A"
        ? settingsSheet.getRange("I8").getValue()
        : sheet.getRange(row, 7).getValue()
    );
  sheet
    .getRange(row, 52)
    .setValue(
      sheet.getRange(row, 8).getValue() === "N/A"
        ? settingsSheet.getRange("I9").getValue()
        : sheet.getRange(row, 8).getValue()
    );
  sheet.getRange(row, 53).setValue(initStock);
}

function priceCalc(row) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("基本設定");
  const slsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SLS送料表");

  // 各シートから必要な情報を取得
  const initWeight = parseFloat(settingsSheet.getRange("C7").getValue()); // 初期重量（未取得時）
  const addWeight = parseFloat(settingsSheet.getRange("C8").getValue()); // 重量加算
  const initLength = parseFloat(settingsSheet.getRange("I7").getValue()); // 初期長さ（未取得時）
  const initWidth = parseFloat(settingsSheet.getRange("I8").getValue()); // 初期幅（未取得時）
  const initHeight = parseFloat(settingsSheet.getRange("I9").getValue()); // 初期高さ（未取得時）
  const profitRatio = settingsSheet.getRange("C11").getValue(); // 利益率
  const calcWeight =
    sheet.getRange(row, 5).getValue() === "N/A"
      ? initWeight
      : parseFloat(sheet.getRange(row, 5).getValue()) + addWeight; // 計算用重量 (kg)
  const calcLength =
    sheet.getRange(row, 6).getValue() === "N/A"
      ? initLength
      : parseFloat(sheet.getRange(row, 6).getValue());
  const calcWitdh =
    sheet.getRange(row, 7).getValue() === "N/A"
      ? initWidth
      : parseFloat(sheet.getRange(row, 7).getValue());
  const calcHeight =
    sheet.getRange(row, 8).getValue() === "N/A"
      ? initHeight
      : parseFloat(sheet.getRange(row, 8).getValue());
  const calcSize = calcLength + calcWitdh + calcHeight;
  const domCost = getDomShippingCost(calcSize, calcWeight); // 国内送料を取得
  const fixedCost =
    parseInt(domCost ? domCost : settingsSheet.getRange("C13").getValue()) +
    parseInt(settingsSheet.getRange("C14").getValue()); // 固定費(国内送料+その他経費)
  const payoneerRatio = settingsSheet.getRange("C15").getValue(); // Payoneer手数料
  const promoRatio = settingsSheet.getRange("C16").getValue(); // プロモーション割引
  const voucherRatio = settingsSheet.getRange("C17").getValue(); // バウチャー割引

  const rateMap = {}; // 各国レート
  rateMap["SG"] = parseFloat(settingsSheet.getRange("F22").getValue()); // SGレート
  rateMap["MY"] = parseFloat(settingsSheet.getRange("F23").getValue()); // MYレート
  rateMap["PH"] = parseFloat(settingsSheet.getRange("F24").getValue()); // PHレート
  rateMap["TH"] = parseFloat(settingsSheet.getRange("F25").getValue()); // THレート
  rateMap["TW"] = parseFloat(settingsSheet.getRange("F26").getValue()); // TWレート
  rateMap["VN"] = parseFloat(settingsSheet.getRange("F27").getValue()); // VNレート
  rateMap["BR"] = parseFloat(settingsSheet.getRange("F28").getValue()); // VNレート

  const deliveryMap = {}; // 各国配送設定
  deliveryMap["SG"] = parseInt(settingsSheet.getRange("C32").getValue());
  deliveryMap["MY"] = parseInt(settingsSheet.getRange("C33").getValue());
  deliveryMap["PH"] = parseInt(settingsSheet.getRange("C34").getValue());
  deliveryMap["TH"] = parseInt(settingsSheet.getRange("C35").getValue());
  deliveryMap["TW"] = parseInt(settingsSheet.getRange("C36").getValue());
  deliveryMap["VN"] = parseInt(settingsSheet.getRange("C37").getValue());
  deliveryMap["BR"] = parseInt(settingsSheet.getRange("C38").getValue());

  const salesMap = {}; // 各国販売手数料
  salesMap["SG"] = settingsSheet.getRange("F32").getValue();
  salesMap["MY"] = settingsSheet.getRange("G32").getValue();
  salesMap["PH"] = settingsSheet.getRange("H32").getValue();
  salesMap["TH"] = settingsSheet.getRange("I32").getValue();
  salesMap["TW"] = settingsSheet.getRange("J32").getValue();
  salesMap["VN"] = settingsSheet.getRange("K32").getValue();
  salesMap["BR"] = settingsSheet.getRange("L32").getValue();

  const paymentMap = {}; // 各国決済手数料
  paymentMap["SG"] = settingsSheet.getRange("F33").getValue();
  paymentMap["MY"] = settingsSheet.getRange("G33").getValue();
  paymentMap["PH"] = settingsSheet.getRange("H33").getValue();
  paymentMap["TH"] = settingsSheet.getRange("I33").getValue();
  paymentMap["TW"] = settingsSheet.getRange("J33").getValue();
  paymentMap["VN"] = settingsSheet.getRange("K33").getValue();
  paymentMap["BR"] = settingsSheet.getRange("L33").getValue();

  const ccbfssMap = {}; // CCB/FSS
  ccbfssMap["SG"] = settingsSheet.getRange("F34").getValue();
  ccbfssMap["MY"] = settingsSheet.getRange("G34").getValue();
  ccbfssMap["PH"] = settingsSheet.getRange("H34").getValue();
  ccbfssMap["TH"] = settingsSheet.getRange("I34").getValue();
  ccbfssMap["TW"] = settingsSheet.getRange("J34").getValue();
  ccbfssMap["VN"] = settingsSheet.getRange("K34").getValue();
  ccbfssMap["BR"] = settingsSheet.getRange("L34").getValue();

  const mdvMap = {}; // MDV
  mdvMap["SG"] = settingsSheet.getRange("F35").getValue();
  mdvMap["MY"] = settingsSheet.getRange("G35").getValue();
  mdvMap["PH"] = settingsSheet.getRange("H35").getValue();
  mdvMap["TH"] = settingsSheet.getRange("I35").getValue();
  mdvMap["TW"] = settingsSheet.getRange("J35").getValue();
  mdvMap["VN"] = settingsSheet.getRange("K35").getValue();
  mdvMap["BR"] = settingsSheet.getRange("L35").getValue();

  const amazonPrice = parseInt(sheet.getRange(row, 4).getValue()); // 新品最安値
  const calcWeightG = calcWeight * 1000; // 計算用重量 (g)
  const slsMap = getSlsShippingCost(calcWeightG); // 各国SLS送料

  const totalCommMap = {}; // 換算後合計手数料
  totalCommMap["SG"] =
    (1 - voucherRatio) * (1 - salesMap["SG"] - ccbfssMap["SG"] - mdvMap["SG"]) -
    paymentMap["SG"];
  totalCommMap["MY"] =
    (1 - voucherRatio) * (1 - salesMap["MY"] - ccbfssMap["MY"] - mdvMap["MY"]) -
    paymentMap["MY"];
  totalCommMap["PH"] =
    (1 - voucherRatio) * (1 - salesMap["PH"] - ccbfssMap["PH"] - mdvMap["PH"]) -
    paymentMap["PH"];
  totalCommMap["TH"] =
    (1 - voucherRatio) * (1 - salesMap["TH"] - ccbfssMap["TH"] - mdvMap["TH"]) -
    paymentMap["TH"];
  totalCommMap["TW"] =
    (1 - voucherRatio) * (1 - salesMap["TW"] - ccbfssMap["TW"] - mdvMap["TW"]) -
    paymentMap["TW"];
  totalCommMap["VN"] =
    (1 - voucherRatio) * (1 - salesMap["VN"] - ccbfssMap["VN"] - mdvMap["VN"]) -
    paymentMap["VN"];
  totalCommMap["BR"] =
    (1 - voucherRatio) * (1 - salesMap["BR"] - ccbfssMap["BR"] - mdvMap["BR"]) -
    paymentMap["BR"];

  const priceMap = {}; // 出品価格
  priceMap["SG"] =
    Math.round(
      ((slsMap["SG"] * (1 - payoneerRatio) +
        (amazonPrice + fixedCost) / rateMap["SG"]) /
        (totalCommMap["SG"] * (1 - payoneerRatio) -
          profitRatio * (1 - voucherRatio)) /
        (1 - promoRatio)) *
        100
    ) / 100;
  priceMap["MY"] =
    Math.round(
      ((slsMap["MY"] * (1 - payoneerRatio) +
        (amazonPrice + fixedCost) / rateMap["MY"]) /
        (totalCommMap["MY"] * (1 - payoneerRatio) -
          profitRatio * (1 - voucherRatio)) /
        (1 - promoRatio)) *
        100
    ) / 100;
  priceMap["PH"] = Math.round(
    (slsMap["PH"] * (1 - payoneerRatio) +
      (amazonPrice + fixedCost) / rateMap["PH"]) /
      (totalCommMap["PH"] * (1 - payoneerRatio) -
        profitRatio * (1 - voucherRatio)) /
      (1 - promoRatio)
  );
  priceMap["TH"] = Math.round(
    (slsMap["TH"] * (1 - payoneerRatio) +
      (amazonPrice + fixedCost) / rateMap["TH"]) /
      (totalCommMap["TH"] * (1 - payoneerRatio) -
        profitRatio * (1 - voucherRatio)) /
      (1 - promoRatio)
  );
  priceMap["TW"] =
    Math.round(
      ((slsMap["TW"] * (1 - payoneerRatio) +
        (amazonPrice + fixedCost) / rateMap["TW"]) /
        (totalCommMap["TW"] * (1 - payoneerRatio) -
          profitRatio * (1 - voucherRatio)) /
        (1 - promoRatio)) *
        100
    ) / 100;
  priceMap["VN"] = Math.round(
    (slsMap["VN"] * (1 - payoneerRatio) +
      (amazonPrice + fixedCost) / rateMap["VN"]) /
      (totalCommMap["VN"] * (1 - payoneerRatio) -
        profitRatio * (1 - voucherRatio)) /
      (1 - promoRatio)
  );
  priceMap["BR"] = Math.round(
    (slsMap["BR"] * (1 - payoneerRatio) +
      (amazonPrice + fixedCost) / rateMap["BR"]) /
      (totalCommMap["BR"] * (1 - payoneerRatio) -
        profitRatio * (1 - voucherRatio)) /
      (1 - promoRatio)
  );

  return priceMap;
}

function getSlsShippingCost(calcWeight) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SLS送料表");
  if (!sheet) {
    Logger.log("SLS送料表シートが見つかりません。");
    return;
  }

  const countries = ["SG", "MY", "PH", "TH", "TW", "VN", "BR"]; // 国の略称
  const slsMap = {};

  // 一度に必要な範囲のデータを全て取得（3行目から最下行まで、12列分のデータ）
  const lastRow = sheet.getLastRow();
  const allData = sheet.getRange(3, 1, lastRow - 2, 14).getValues();

  // 各国ごとの処理
  countries.forEach((country, index) => {
    const weightCol = index * 2; // 各国の重量列（0ベース）
    const costCol = weightCol + 1; // 各国の送料列（0ベース）

    let foundCost = null;
    let lastValidCost = null; // 最下行の送料を保存

    for (let i = 0; i < allData.length; i++) {
      const weight = allData[i][weightCol]; // 重量
      const cost = allData[i][costCol]; // 送料

      // 空のセルが出たらその国のデータはここで終了
      if (weight === "" || cost === "") {
        break; // この国の処理を終了
      }

      lastValidCost = cost; // 常に最下行の送料を保持

      // calcWeight がその重量に達したか、それより小さい場合はその送料を取得
      if (calcWeight <= weight) {
        foundCost = cost;
        break; // 一致または超えた場合、その送料を取得してループを終了
      }
    }

    // 最下行を超えた場合、最下行の送料を使用
    if (foundCost === null) {
      foundCost = lastValidCost; // 最下行の送料を使用
    }

    // slsMapに国別の送料を格納
    slsMap[country] = foundCost;
  });

  return slsMap;
}

function testGetDomShippingCost() {
  const size = 220;
  const weight = 4.5;
  const response = getDomShippingCost(size, weight);
  Logger.log(response);
}

function getDomShippingCost(size, weight) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("追加設定");
  if (!sheet) {
    Logger.log("追加設定シートが見つかりません。");
    return null;
  }

  const startRow = 17;
  const sizeRange = sheet.getRange(`B${startRow}:B`).getValues(); // B列のサイズ
  const weightRange = sheet.getRange(`C${startRow}:C`).getValues(); // C列の重量
  const costRange = sheet.getRange(`D${startRow}:D`).getValues(); // D列のコスト

  let lastRow = startRow;
  // 最下行を見つける（D列にデータがある最終行を特定）
  while (costRange[lastRow - startRow][0] !== "") {
    lastRow++;
  }

  // 両方の条件（sizeとweight）を満たす行の次のcostを取得
  for (let i = 0; i < lastRow - startRow - 1; i++) {
    // 次の行のコストが存在するため -1
    const rowSize = sizeRange[i][0];
    const rowWeight = weightRange[i][0];

    // sizeとweightが両方とも行の値以下であるかをチェック
    if (size <= rowSize && weight <= rowWeight) {
      // 条件を満たした行のcostを返す
      return parseInt(costRange[i][0], 10); // その行のコストを返す
    }
  }

  // すべての行を満たさない場合、最後の行のcostを返す
  return parseInt(costRange[lastRow - startRow - 1][0], 10);
}

/**
 * バリエーションを含めたバージョン
 */
function uploadImageToShopee(row) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");

  const currentDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy/MM/dd"
  );
  sheet.getRange(row, 55).setValue(currentDate); // BA列：画像登録日

  // メイン商品の画像を処理
  for (let i = 0; i < 8; i++) {
    const fileId = sheet.getRange(row, 15 + i).getValue(); // P列以降に格納された画像ファイルID

    if (fileId !== "N/A" && fileId !== "") {
      const imageId = postUploadImage(fileId);
      if (imageId) {
        sheet.getRange(row, 56 + i).setValue(imageId); // BB列以降にアップロードした画像IDを設定
      } else {
        Logger.log(`画像アップロード失敗: row=${row}, col=${15 + i}`);
      }
    }
  }

  // バリエーション商品の画像を処理
  let i = 0; // ループのインデックス
  variationasin_col = COL_VARIATION_OUTPUT_START;
  while (true) {
    const asinCell = sheet.getRange(row, variationasin_col + VARIATION_ASIN); // バリエーションASINのセル
    const imageCell = sheet.getRange(
      row,
      variationasin_col + VARIATION_IMAGE_ID
    ); // バリエーション画像IDのセル

    const asin = asinCell.getValue();
    if (!asin) break;
    const fileId = imageCell.getValue();

    // ASINが存在しない場合、ループを終了
    if (!asin || asin === "" || asin === "N/A") {
      break;
    }

    // ファイルIDが存在する場合のみ画像アップロード
    if (fileId && fileId !== "" && fileId !== "N/A") {
      const imageId = postUploadImage(fileId);
      if (imageId) {
        sheet
          .getRange(row, variationasin_col + VARIATION_IMAGE_ID_CON)
          .setValue(imageId); // 画像IDを設定
      } else {
        Logger.log(
          `バリエーション画像アップロード失敗: row=${row}, variation_index=${i}`
        );
      }
    }

    variationasin_col += VARIATION_DATAS; // 次のバリエーションを処理
  }
}

// /**
//  * 画像を Shopee にアップロードし、画像IDを取得
//  * @param {string} fileId - アップロードする画像ファイルID
//  * @returns {string|null} - アップロード成功時の画像ID、失敗時はnull
//  */
// function postUploadImage(fileId) {
//   try {
//     const url = 'https://partner.shopeemobile.com/api/v2/media_space/upload_image';
//     const imageBlob = DriveApp.getFileById(fileId).getBlob(); // Google Driveからファイルを取得

//     const options = {
//       method: 'post',
//       contentType: 'multipart/form-data',
//       payload: {
//         image: imageBlob
//       }
//     };

//     const response = UrlFetchApp.fetch(url, options);
//     const responseData = JSON.parse(response.getContentText());

//     if (responseData.error) {
//       Logger.log(`画像アップロードエラー: ${responseData.message}`);
//       return null;
//     }

//     return responseData.image_id;
//   } catch (e) {
//     Logger.log(`画像アップロードに失敗しました: ${e.message}`);
//     return null;
//   }
// }
function processAddShopeeItem(row, shop_id, access_token, region, shop_name) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
  const histSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品履歴");
  const setupSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("基本設定");

  const item_name_en = sheet.getRange(row, 44).getValue(); // AR列：商品名 (英)
  const item_name_ch = sheet.getRange(row, 46).getValue(); // AT列：商品名 (中)
  const cover_image = sheet.getRange(row, 54).getValue(); // BB列：画像ID

  let item_name = "";
  if (region === "TW") {
    item_name = item_name_ch;
  } else {
    item_name = item_name_en;
  }

  const category_id = getCategoryRecommend(
    shop_id,
    access_token,
    item_name,
    cover_image[0]
  );
  //Logger.log(category_id);
  const attributeList = getMandatoryAttributes(
    shop_id,
    access_token,
    category_id
  );
  //Logger.log(attributeList);
  const formattedAttribute = attributeList.map((item) => ({
    attribute_id: item.attribute_id,
    attribute_value_list: [
      {
        value_id: item.first_value.value_id,
        original_value_name: item.first_value.value_name,
      },
    ],
  }));
  //Logger.log(formattedAttribute);

  // ブランドリストを取得し、1つ目のブランドを設定
  const brandList = getBrandList(shop.shop_id, shop.access_token, categoryId);
  const brandName = brandList.length > 0 ? brandList[0].brand_name : "NoBrand";
  Logger.log(`ブランド名: ${brandName}`);

  // 商品価格の設定
  const priceMap = {}; // 商品価格
  priceMap["SG"] = parseFloat(sheet.getRange(row, 32).getValue());
  priceMap["MY"] = parseFloat(sheet.getRange(row, 33).getValue());
  priceMap["PH"] = parseInt(sheet.getRange(row, 34).getValue());
  priceMap["TH"] = parseInt(sheet.getRange(row, 35).getValue());
  priceMap["TW"] = parseFloat(sheet.getRange(row, 36).getValue());
  priceMap["VN"] = parseInt(sheet.getRange(row, 37).getValue());
  priceMap["BR"] = parseInt(sheet.getRange(row, 38).getValue());

  let description = "";
  if (region === "TW") {
    description = sheet.getRange(row, 49).getValue();
  } else {
    description = sheet.getRange(row, 47).getValue();
  }

  const logisticMap = {};
  logisticMap["SG"] = parseInt(setupSheet.getRange("C32").getValue());
  logisticMap["MY"] = parseInt(setupSheet.getRange("C33").getValue());
  logisticMap["PH"] = parseInt(setupSheet.getRange("C34").getValue());
  logisticMap["TH"] = parseInt(setupSheet.getRange("C35").getValue());
  logisticMap["TW"] = parseInt(setupSheet.getRange("C36").getValue());
  logisticMap["VN"] = parseInt(setupSheet.getRange("C37").getValue());
  logisticMap["VN"] = parseInt(setupSheet.getRange("C38").getValue());

  const imageRange = sheet.getRange(row, 56, 1, 8); // 1行目（指定行）の54列目(BB)から8列(BB~BI)
  const imageValues = imageRange.getValues()[0]; // 1行だけなので[0]で取得
  const imageList = imageValues.filter(
    (value) => value !== null && value !== ""
  );

  const item_sku = sheet.getRange(row, 31).getValue();

  const payload = {
    original_price: priceMap[region], // 商品価格
    description: description, // 商品説明
    weight: parseFloat(sheet.getRange(row, 50).getValue()), // 重量 (kg)
    item_name: item_name, // 商品名
    dimension: {
      package_length: parseInt(sheet.getRange(row, 51).getValue()), // 長さ (cm)
      package_width: parseInt(sheet.getRange(row, 52).getValue()), // 幅 (cm)
      package_height: parseInt(sheet.getRange(row, 53).getValue()), // 高さ (cm)
    },
    logistic_info: [
      {
        enabled: true, // 配送方法有効化
        logistic_id: logisticMap[region], // ロジスティックID
      },
    ],
    category_id: parseInt(category_id), // カテゴリID
    image: {
      image_id_list: imageList,
    },
    brand: {
      brand_id: parseInt("0", 10), // ブランドID
      original_brand_name: "NoBrand", // ブランド名
    },
    item_sku: item_sku, // SKU
    seller_stock: [
      {
        stock: parseInt(sheet.getRange(row, 54).getValue()), // 在庫
      },
    ],
    attribute_list: formattedAttribute,
  };
  const response = postAddItem(shop_id, access_token, payload);
  const responseData = JSON.parse(response.getContentText());
  // 出品結果を履歴にプロット
  const timestamp = new Date();
  const logData = [
    timestamp,
    item_sku,
    region,
    shop_name,
    responseData?.response?.item_id ? responseData.response.item_id : "",
    responseData.error,
    responseData.message,
    responseData.warning,
    payload,
    response,
  ];

  histSheet.appendRow(logData);
}

///////////////////////////////////////////////////////////////////////////////////////////
/**
 * Shopee 商品追加およびバリエーション作成のテストプログラム
 */
function testAddProductWithVariationsStep3() {
  // selectedShops = [0,1,2,3,4,5,6,7,8]
  selectedShops = [8];
  processMakeDataStep3(selectedShops);
}
function testfetchAllCategoriesAndBrands() {
  // selectedShops = [0,1,2,3,4,5,6]
  selectedShops = [2, 3, 4, 5, 6, 7, 8];
  fetchAllCategoriesAndBrands(selectedShops);
}
function testAddProductWithVariations() {
  // selectedShops = [0,1,2,3,4,5,6,7,8]
  selectedShops = [2, 3, 4, 5, 6, 7, 8];
  addProductWithVariations(selectedShops);
}

/**
 * Shopee 商品追加およびバリエーション作成のプログラム
 * Step1～3まで完了して商品画像のアップロードとブランドリストの更新を完了してから実行してください
 */
function addProductWithVariations(selectedShopsIndices) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
  const histSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品履歴");
  const setupSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("基本設定");

  const lastRow = sheet.getLastRow();
  const checkRange = sheet.getRange(2, 1, lastRow - 1); // チェックボックス列
  const checkValues = checkRange.getValues(); // チェックボックス状態

  // チェックされた行を収集
  const rowsToProcess = checkValues.reduce((acc, val, index) => {
    if (val[0]) acc.push(index + 2); // チェックされた行を追加
    return acc;
  }, []);

  if (rowsToProcess.length === 0) {
    Logger.log("処理対象の行がありません。");
    return;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  const shopsData = JSON.parse(scriptProperties.getProperty("SHOPS_DATA"));

  rowsToProcess.forEach((row) => {
    try {
      selectedShopsIndices.forEach((index) => {
        try {
          const shop = shopsData[index];
          Logger.log(`${index}番目の処理 ${shop.shop_name}`);
          const { tierVariation, models } = getVariations(
            sheet,
            row,
            shop.region
          );
          // バリエーションの組み合わせ数を計算
          const expectedModelsCount =
            tierVariation[0].option_list.length *
            (tierVariation[1]?.option_list.length || 1);
          if (models.length !== expectedModelsCount) {
            throw new Error(
              `モデルの数がバリエーションの組み合わせ数と一致しません: ${models.length} != ${expectedModelsCount}`
            );
          }
          // ペイロードをログに記録
          Logger.log(`Tier Variation: ${JSON.stringify(tierVariation)}`);
          Logger.log(`Models: ${JSON.stringify(models)}`);

          const {
            itemName,
            description,
            priceMap,
            weight,
            dimensions,
            imageIds,
            itemSku,
          } = getProductDetails(sheet, setupSheet, row, shop.region);

          if (!Array.isArray(imageIds) || imageIds.length === 0) {
            Logger.log(`画像IDが無効です (行: ${row})`);
            throw new Error("画像IDが無効です");
            // return; // 処理を中断
          }
          const categoryId = getCategoryRecommend(
            shop.shop_id,
            shop.access_token,
            itemName,
            imageIds[0]
          )[0];
          // 必須情報の取得
          const attributeList = getMandatoryAttributes(
            shop.shop_id,
            shop.access_token,
            categoryId
          );
          const formattedAttribute = attributeList.map((item) => ({
            attribute_id: item.attribute_id,
            attribute_value_list: [
              {
                value_id: item.first_value.value_id,
                original_value_name: item.first_value.value_name,
              },
            ],
          }));
          //Logger.log(formattedAttribute);
          const logisticInfo = getValidLogistics(
            shop.shop_id,
            shop.access_token
          );

          // Shopeeブランドリストをスプレッドシートから取得し、カテゴリーIDでフィルタリング
          const allShopeeBrands =
            getShopeeBrandsFromSheetByCategory(categoryId);
          if (!allShopeeBrands || allShopeeBrands.length === 0) {
            Logger.log(
              `カテゴリーID "${categoryId}" に一致するShopeeブランドリストを取得できませんでした。`
            );
            throw new Error("Shopeeブランドリストが空です");
            // return;
          }

          Logger.log(
            `取得したブランド数 (カテゴリーID: ${categoryId}): ${allShopeeBrands.length}`
          );

          // ブランドマッチング
          const keepaBrand = sheet
            .getRange(row, COL_BRAND_FROM_KEEPA)
            .getValue();
          const nearestBrand = matchBrandWithShopee(
            keepaBrand,
            allShopeeBrands
          );
          if (!nearestBrand) {
            Logger.log(
              `Keepaブランド "${keepaBrand}" に一致するShopeeブランドが見つかりませんでした (行: ${row})。`
            );
            throw new Error("ブランドマッチングに失敗しました");
          }

          Logger.log(
            `Keepaブランド "${keepaBrand}" に最も近いShopeeブランドは "${nearestBrand.brand_id}、${nearestBrand.original_brand}" です。`
          );
          // メイン商品の登録
          let mainPayload = {
            original_price: priceMap[shop.region],
            description,
            weight,
            item_name: itemName,
            dimension: {
              package_length: dimensions.package_length || 1,
              package_width: dimensions.package_width || 1,
              package_height: dimensions.package_height || 1,
            },
            category_id: categoryId,
            image: { image_id_list: imageIds },
            brand: {
              brand_id: nearestBrand.brand_id,
              original_brand_name: nearestBrand.brand_name,
            },
            // seller_stock: [{ stock: 10 }], // デフォルト在庫数
            seller_stock: [
              {
                stock: parseInt(sheet.getRange(row, 52).getValue()), // 在庫
              },
            ],
            logistic_info: logisticInfo,
            attribute_list: formattedAttribute,
          };

          let mainResponse = postAddItem(
            shop.shop_id,
            shop.access_token,
            mainPayload
          );
          let mainResult = JSON.parse(mainResponse.getContentText());

          // ブランドが無効な場合、"No Brand" を使用して再試行
          if (
            mainResult.error &&
            mainResult.message.includes(
              "The brand is invalid, please check and update"
            )
          ) {
            Logger.log(
              `ブランドが無効のため、"No Brand" を使用して再試行します。`
            );
            mainPayload.brand = {
              brand_id: 0, // Shopeeで "No Brand" に対応する ID
              original_brand_name: "No Brand",
            };

            mainResponse = postAddItem(
              shop.shop_id,
              shop.access_token,
              mainPayload
            );
            mainResult = JSON.parse(mainResponse.getContentText());
          }
          // 再試行後もエラーがある場合は例外をスロー
          if (mainResult.error) {
            Logger.log(
              `${shop.shop_name} メイン商品登録エラー: ${mainResult.message}`
            );
            // throw new Error(mainResult.message);
          }

          Logger.log(
            `メイン商品が登録されました (Item ID: ${mainResult.response.item_id})`
          );

          // バリエーションの登録
          if (tierVariation.length > 0 && models.length > 0) {
            try {
              const variationResponse = initTierVariation(
                shop.shop_id,
                shop.access_token,
                mainResult.response.item_id,
                tierVariation,
                models
              );

              if (variationResponse.error) {
                Logger.log(
                  `${shop.shop_name} バリエーション登録エラー: ${variationResponse.message}`
                );
              } else {
                Logger.log(
                  `${shop.shop_name} バリエーションが登録されました (Item ID: ${mainResult.response.item_id})`
                );
              }
              // 出品結果を履歴にプロット
              const timestamp = new Date();
              const logData = [
                timestamp,
                itemSku,
                shop.region,
                shop.name,
                variationResponse?.response?.item_id
                  ? variationResponse.response.item_id
                  : "",
                variationResponse.error,
                variationResponse.message,
                variationResponse.warning,
                mainPayload,
                variationResponse,
              ];
              histSheet.appendRow(logData);
            } catch (error) {
              Logger.log(
                `${shop.shop_name} バリエーション登録エラー (行: ${row}): ${error.message}`
              );
            }
          } else {
            // 出品結果を履歴にプロット
            const timestamp = new Date();
            const logData = [
              timestamp,
              itemSku,
              shop.region,
              shop.name,
              mainResult?.response?.item_id ? mainResult.response.item_id : "",
              mainResult.error,
              mainResult.message,
              mainResult.warning,
              payload,
              mainResult.response,
            ];
            histSheet.appendRow(logData);
          }
        } catch (shopError) {
          Logger.log(
            `ショップ処理中にエラーが発生しました (行: ${row}, ショップ: ${shopsData[index].shop_name}): ${shopError.message}`
          );
        }
      });

      // チェックボックスをオフにし、処理日付を設定
      sheet.getRange(row, 1).setValue(false);
      sheet
        .getRange(row, COL_SET_DAY)
        .setValue(
          Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            "yyyy/MM/dd"
          )
        );
    } catch (error) {
      Logger.log(`エラーが発生しました (行: ${row}): ${error.message}`);
    }
  });
}

/**
 * 商品画像のアップロードとブランドIDの取得
 */
function processMakeDataStep3(selectedShopsIndices) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
  const histSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品履歴");
  const setupSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("基本設定");
  const shopeeBrandSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    BRAND_CACHE_SHEET_NAME
  );

  const lastRow = sheet.getLastRow();
  const checkRange = sheet.getRange(2, 1, lastRow - 1); // チェックボックス列
  const checkValues = checkRange.getValues(); // チェックボックス状態

  // チェックされた行を収集
  const rowsToProcess = checkValues.reduce((acc, val, index) => {
    if (val[0]) acc.push(index + 2); // チェックされた行を追加
    return acc;
  }, []);

  if (rowsToProcess.length === 0) {
    Logger.log("処理対象の行がありません。");
    return;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  const shopsData = JSON.parse(scriptProperties.getProperty("SHOPS_DATA"));

  rowsToProcess.forEach((row) => {
    try {
      //画像イメージをShopeeにアップロード
      uploadImageToShopee(row);

      selectedShopsIndices.forEach((index) => {
        try {
          const shop = shopsData[index];

          const {
            itemName,
            description,
            priceMap,
            weight,
            dimensions,
            imageIds,
            itemSku,
          } = getProductDetails(sheet, setupSheet, row, shop.region);

          if (!Array.isArray(imageIds) || imageIds.length === 0) {
            Logger.log(`画像IDが無効です (行: ${row})`);
            throw new Error("画像IDが無効です");
            // return; // 処理を中断
          }
          const categoryId = getCategoryRecommend(
            shop.shop_id,
            shop.access_token,
            itemName,
            imageIds[0]
          )[0];
          // 必須情報の取得

          // CatetoryごとのShopeeブランドリストをすべて取得
          // fetchAllShopeeBrands(shop.shop_id, shop.access_token, categoryId);
          fetchBrandListFromShopee(shop.shop_id, shop.access_token, categoryId); // ブランドリストを取得
          // saveBrandListToSheet(sheet, categoryId, brandList); // ブランドリストを保存
          // translateShopeeBrandsWithDeepL(brandList);
        } catch (shopError) {
          Logger.log(
            `ショップ処理中にエラーが発生しました (行: ${row}, ショップ: ${shopsData[index].shop_name}): ${shopError.message}`
          );
        }
      });
    } catch (error) {
      Logger.log(`エラーが発生しました (行: ${row}): ${error.message}`);
    }
  });
}
// 補助関数群
/**
 * 商品詳細を取得する
 */
function getProductDetails(sheet, setupSheet, row, region) {
  const itemNameEn = sheet
    .getRange(row, COL_MAIN_PRODUCT_NAME_EN_CON)
    .getValue(); // 商品名 (英)
  const item_name_en_120 =
    itemNameEn.length > 120 ? itemNameEn.substring(0, 120) : itemNameEn;
  const itemNameCh = sheet
    .getRange(row, COL_MAIN_PRODUCT_NAME_ZH_CON)
    .getValue(); // 商品名 (中)
  const itemName =
    region === "BR"
      ? item_name_en_120
      : region === "TW"
      ? itemNameCh
      : itemNameEn;
  const description =
    region === "TW"
      ? sheet.getRange(row, 49).getValue()
      : sheet.getRange(row, 47).getValue();
  const priceMap = getPriceMap(sheet, row);
  const weight = parseFloat(sheet.getRange(row, 50).getValue());
  const dimensions = {
    length: parseInt(sheet.getRange(row, 51).getValue()),
    width: parseInt(sheet.getRange(row, 52).getValue()),
    height: parseInt(sheet.getRange(row, 53).getValue()),
  };
  const imageIds = getImageIds(sheet, row);
  const itemSku = sheet.getRange(row, 31).getValue();

  return {
    itemName,
    description,
    priceMap,
    weight,
    dimensions,
    imageIds,
    itemSku,
  };
}

/**
 * メイン商品の価格Mapを取得
 */
function getPriceMap(sheet, row) {
  return {
    SG: parseFloat(sheet.getRange(row, 32).getValue()),
    MY: parseFloat(sheet.getRange(row, 33).getValue()),
    PH: parseInt(sheet.getRange(row, 34).getValue()),
    TH: parseInt(sheet.getRange(row, 35).getValue()),
    TW: parseFloat(sheet.getRange(row, 36).getValue()),
    VN: parseInt(sheet.getRange(row, 37).getValue()),
    BR: parseFloat(sheet.getRange(row, 38).getValue()),
  };
}

/**
 * メイン画像の画像IDを収集
 */
function getImageIds(sheet, row) {
  const imageIds = [];

  // メイン商品の画像を収集
  const mainImageRange = sheet.getRange(row, 56, 1, 8); // 54列目から8列
  const mainImageValues = mainImageRange.getValues()[0]; // 1行だけなので[0]で取得
  imageIds.push(
    ...mainImageValues.filter((value) => value !== null && value !== "")
  );

  return imageIds;
}

/**
 * バリエーションデータ取得ロジック
 */
function getVariations(sheet, row, region) {
  const variationType = sheet.getRange(row, COL_VARIATION_TYPE).getValue(); // バリエーションタイプ
  const tierVariation = [];
  const models = [];

  if (variationType === "1つ" || variationType === "2つ") {
    const regionindex = region === "TW" ? 2 : 1;
    // バリエーション1
    const variation1NameOrg = sheet
      .getRange(row, COL_VARIATION1_NAME + regionindex)
      .getValue();
    const variation1Name =
      variation1NameOrg.length > 14
        ? variation1NameOrg.substring(0, 14)
        : variation1NameOrg;
    // 全列データを取得
    // const variation1Values = sheet.getRange(row, COL_VARIATION1_VALUES_START, 3, 10).getValues()[0].filter(v => v);
    const allValuesvariation1 = sheet
      .getRange(row, COL_VARIATION1_VALUES_START, 1, 3 * VARIATION_COLUMNS)
      .getValues()[0]; // 30列分取得（必要に応じて範囲を変更）

    // 3列飛びで値を抽出
    const variation1Values = allValuesvariation1.filter(
      (_, index) => index % 3 === regionindex && allValuesvariation1[index]
    );

    Logger.log(`バリエーション1の値 : ${variation1Values}`);

    // バリエーション1の画像IDを取得（40列飛びで）
    let variation1ImageCol =
      COL_VARIATION_OUTPUT_START + VARIATION_IMAGE_ID_CON;
    const variation1ImageIds = [];
    variation1Values.forEach(() => {
      variation1ImageIds.push(
        sheet.getRange(row, variation1ImageCol).getValue()
      );
      variation1ImageCol += VARIATION_DATAS; // 40列飛び
    });

    tierVariation.push({
      name: variation1Name,
      option_list: variation1Values.map((value, index) => ({
        option: value,
        image: variation1ImageIds[index]
          ? { image_id: variation1ImageIds[index] } // 期待される形式
          : null, // 画像がない場合は null
      })),
    });

    // バリエーション2（存在する場合）
    let variation2Values = [];
    let variation2ImageIds = [];
    if (variationType === "2つ") {
      const variation2NameOrg = sheet
        .getRange(row, COL_VARIATION2_NAME + regionindex)
        .getValue();
      const variation2Name =
        variation2NameOrg.length > 14
          ? variation2NameOrg.substring(0, 14)
          : variation2NameOrg;
      // variation2Values = sheet.getRange(row, COL_VARIATION2_VALUES_START, 3, 10).getValues()[0].filter(v => v);
      const allValuesvariation2 = sheet
        .getRange(row, COL_VARIATION2_VALUES_START, 1, 3 * VARIATION_COLUMNS)
        .getValues()[0]; // 30列分取得（必要に応じて範囲を変更）

      // 3列飛びで値を抽出
      variation2Values = allValuesvariation2.filter(
        (_, index) => index % 3 === regionindex && allValuesvariation2[index]
      );

      Logger.log(`バリエーション2の値 : ${variation2Values}`);

      // バリエーション2の画像IDを取得（40列飛びで）
      let variation2ImageCol = variation1ImageCol; // バリエーション2の画像列開始位置
      variation2Values.forEach(() => {
        variation2ImageIds.push(
          sheet.getRange(row, variation2ImageCol).getValue()
        );
        variation2ImageCol += VARIATION_DATAS; // 40列飛び
      });

      tierVariation.push({
        name: variation2Name,
        option_list: variation2Values.map((value, index) => ({
          option: value,
          image: variation2ImageIds[index]
            ? { image_id: variation2ImageIds[index] } // 期待される形式
            : null, // 画像がない場合は null
        })),
      });
    }

    // 組み合わせの生成（BL列以降のデータを取得）
    let currentCol = COL_VARIATION_OUTPUT_START + VARIATION_COL; // BL列
    const variationCombinations = [];
    while (sheet.getRange(row, currentCol).getValue()) {
      variationCombinations.push(sheet.getRange(row, currentCol).getValue());
      currentCol += VARIATION_DATAS;
    }

    // モデル情報の基準列
    let priceCol = COL_VARIATION_OUTPUT_START;
    let stockCol = COL_VARIATION_OUTPUT_START + VARIATION_ZAIKO_CON;
    let skuCol = COL_VARIATION_OUTPUT_START + VARIATION_SKU;
    let asinCol = COL_VARIATION_OUTPUT_START + VARIATION_ASIN;
    let variationCol = COL_VARIATION_OUTPUT_START + VARIATION_COL;
    let weightCol = COL_VARIATION_OUTPUT_START + VARIATION_WEIGHT;
    let lengthCol = COL_VARIATION_OUTPUT_START + VARIATION_LENGTH;
    let widthCol = COL_VARIATION_OUTPUT_START + VARIATION_WIDTH;
    let heightCol = COL_VARIATION_OUTPUT_START + VARIATION_HEIGHT;

    let currentRow = row;
    // // バリエーション値を配列として取得
    // const variation1ValuesCol = sheet.getRange(row, COL_VARIATION1_VALUES_START, 1, 10).getValues()[0];
    // const variation2ValuesCol = variationType === "2つ"
    //   ? sheet.getRange(row, COL_VARIATION2_VALUES_START, 1, 10).getValues()[0]
    //   : [];

    while (sheet.getRange(currentRow, asinCol).getValue()) {
      // 現在の行のデータを取得
      let priceMap = getPriceMapVariation(sheet, currentRow, priceCol);
      let stock = sheet.getRange(currentRow, stockCol).getValue();
      let sku = sheet.getRange(currentRow, skuCol).getValue();
      let variationCombination = sheet
        .getRange(currentRow, variationCol)
        .getValue();
      let weight = sheet.getRange(currentRow, weightCol).getValue();
      let length = sheet.getRange(currentRow, lengthCol).getValue();
      let width = sheet.getRange(currentRow, widthCol).getValue();
      let height = sheet.getRange(currentRow, heightCol).getValue();

      const combinationIndex =
        variationCombinations.indexOf(variationCombination);

      if (combinationIndex === -1) {
        throw new Error(
          `Invalid combination at row ${currentRow}: ${combinationString} is not in the generated combinations.`
        );
      }

      // ティアインデックスを計算
      const tierIndex =
        variationType === "2つ"
          ? [
              Math.floor(combinationIndex / variation2Values.length),
              combinationIndex % variation2Values.length,
            ]
          : [combinationIndex];

      // ティアインデックスをログに出力
      Logger.log(
        `Row ${currentRow}: combinationIndex = ${combinationIndex}, tierIndex = ${JSON.stringify(
          tierIndex
        )}`
      );

      // モデルデータを追加
      models.push({
        tier_index: tierIndex,
        original_price: priceMap[region],
        normal_stock: stock,
        model_sku: sku,
        weight: weight, // バリエーションごとの重量
        dimension: {
          // バリエーションごとの梱包サイズ
          package_length: length || 1,
          package_width: width || 1,
          package_height: height || 1,
        },
      });

      // 40列飛ばし
      priceCol += VARIATION_DATAS;
      stockCol += VARIATION_DATAS;
      skuCol += VARIATION_DATAS;
      asinCol += VARIATION_DATAS;
      variationCol += VARIATION_DATAS;
      weightCol += VARIATION_DATAS;
      lengthCol += VARIATION_DATAS;
      widthCol += VARIATION_DATAS;
      heightCol += VARIATION_DATAS;
    }
  }
  // 最大価格の取得と価格の調整
  const maxPrice = Math.max(...models.map((model) => model.original_price));
  models.forEach((model) => {
    if (model.original_price < maxPrice * 0.2) {
      model.original_price = Math.round(maxPrice * 0.2); // 20%に設定
    }
  });
  return { tierVariation, models };
}

/**
 * ShopeeのAddItem APIに必要なペイロードを生成する
 * @param {Object} productDetails - 商品の詳細情報
 * @returns {Object} ペイロードオブジェクト
 */
function createPayload({
  itemName,
  description,
  price,
  weight,
  dimensions,
  logisticInfo,
  categoryId,
  imageIds,
  formattedAttribute,
  itemSku,
  tierVariation,
  models,
}) {
  // 必須フィールドのチェック
  if (
    !itemName ||
    !description ||
    !price ||
    !weight ||
    !categoryId ||
    !imageIds ||
    imageIds.length === 0
  ) {
    throw new Error("Missing required fields in payload.");
  }

  // ペイロードオブジェクト
  const payload = {
    original_price: price, // 商品価格
    description: description, // 商品説明
    weight: weight, // 商品重量
    item_name: itemName, // 商品名
    dimension: {
      package_length: dimensions.package_length, // 長さ
      package_width: dimensions.package_width, // 幅
      package_height: dimensions.package_height, // 高さ
    },
    category_id: categoryId, // カテゴリID
    image: {
      image_id_list: imageIds, // アップロードされた画像IDリスト
    },
    seller_stock: [
      {
        stock: models ? undefined : 0, // モデルがある場合はモデルで在庫管理
      },
    ],
    logistic_info: logisticInfo, // 配送情報
    attribute_list: formattedAttribute, // 必須属性
    item_sku: itemSku, // SKU
    tier_variation: tierVariation || [], // バリエーション情報（任意）
    model: models || [], // モデル情報（任意）
  };

  Logger.log("Generated Payload: " + JSON.stringify(payload, null, 2)); // JSON文字列化してログ出力
  return payload;
}

function logResponse(histSheet, row, shop, response, payload) {
  const responseData = JSON.parse(response.getContentText());
  const logData = [
    new Date(),
    payload.item_sku,
    shop.region,
    shop.shop_name,
    responseData?.response?.item_id || "",
    responseData.error || "",
    responseData.message || "",
    responseData.warning || "",
    JSON.stringify(payload),
    JSON.stringify(response),
  ];
  histSheet.appendRow(logData);
}

/**
 * 指定行からバリエーション情報を取得
 *
 * @param {number} row - スプレッドシートの行番号
 * @return {Object} バリエーションデータ
 */
function getVariationData(sheet, row, nameCol, startCol, endCol) {
  const name = sheet.getRange(row, nameCol).getValue(); // バリエーション名
  const values = [];
  for (let col = startCol; col <= endCol; col += 3) {
    const value = sheet.getRange(row, col).getValue(); // 各列の値を取得
    if (value) values.push(value); // 値が存在する場合のみ追加
  }
  return { name, values };
}

/**
 * Keepa APIを呼び出してAmazonの商品情報
 */
function processKeepaProductVariationRow(row) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODUCTS);
  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);

  // 基本設定シートからAPIキーとフォルダIDを取得
  const apiKey = settingsSheet.getRange("C4").getValue();
  // const apiKey = PropertiesService.getScriptProperties().getProperty('KEEPA_API_KEY');
  const folderId = settingsSheet.getRange("C5").getValue();

  if (!sheet || !apiKey || !folderId) {
    Logger.log("出品一覧シート、APIキー、またはフォルダIDが見つかりません。");
    return;
  }

  // バリエーション分ループ
  // 1バリエーション10商品までループ
  const asin = sheet.getRange(row, COL_ASIN).getValue(); // メインASIN

  if (!asin) {
    Logger.log(`行${row}にASINがありません。`);
    return;
  }

  // 処理開始前にC列～AD列をクリア
  sheet.getRange(row, COL_DATA_DATE, 1, VARIATION_DATAS).clearContent(); // C列からAD列をクリア

  let productData = postKeepaProducts(apiKey, asin);
  if (!productData) {
    Logger.log("APIトークンが無くなりました。");
    return;
  }

  // productDataがJSON文字列なので、オブジェクトに変換
  productData = JSON.parse(productData);

  // データが正しく取得できたか確認
  Logger.log(JSON.stringify(productData)); // ログでデータ確認用

  // 画像をダウンロードしてGoogle Driveに保存
  const imageIds = [];
  if (productData.image) {
    const images = productData.image.split(",");
    Logger.log(images);
    for (let j = 0; j < images.length && j < 8; j++) {
      const fileId = downloadImageAndSave(images[j], folderId);
      imageIds.push(fileId || "N/A");
    }
  }

  // シートにプロット
  const currentDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy/MM/dd"
  );
  sheet.getRange(row, 3).setValue(currentDate); // C列：データ取得日
  sheet.getRange(row, 4).setValue(productData.price || "N/A"); // D列：新品最安値
  sheet
    .getRange(row, 5)
    .setValue(productData.weight ? productData.weight / 1000 : "N/A"); // E列：重量 (kg)
  sheet
    .getRange(row, 6)
    .setValue(productData.length ? productData.length / 10 : "N/A"); // F列：長さ (cm)
  sheet
    .getRange(row, 7)
    .setValue(productData.width ? productData.width / 10 : "N/A"); // G列：幅 (cm)
  sheet
    .getRange(row, 8)
    .setValue(productData.height ? productData.height / 10 : "N/A"); // H列：高さ (cm)
  sheet.getRange(row, 9).setValue(productData.title || "N/A"); // I列：商品名

  // featuresとdescriptionの結合（改行を減らす）
  const features = productData.features
    ? productData.features.join("\n").replace(/\n+/g, "\n")
    : "";
  const description = productData.description
    ? productData.description.replace(/\n+/g, "\n")
    : "";
  sheet.getRange(row, 10).setValue(features + "\n" + description || "N/A"); // J列：商品説明

  // 翻訳用の数式をセット
  sheet.getRange(row, 11).setFormula(`=GOOGLETRANSLATE(I${row}, "ja", "en")`); // K列：商品名 (英)
  sheet.getRange(row, 12).setFormula(`=GOOGLETRANSLATE(J${row}, "ja", "en")`); // L列：商品説明 (英)
  sheet
    .getRange(row, 13)
    .setFormula(`=GOOGLETRANSLATE(I${row}, "ja", "zh-CN")`); // M列：商品名 (中)
  sheet
    .getRange(row, 14)
    .setFormula(`=GOOGLETRANSLATE(J${row}, "ja", "zh-CN")`); // N列：商品説明 (中)

  // 画像ファイルIDをプロット、N/Aの場合はイメージ欄もN/Aにする
  for (let j = 0; j < 8; j++) {
    const imageId = imageIds[j] || "N/A";
    sheet.getRange(row, 15 + j).setValue(imageId); // O～V列：画像①～⑧
    var imageFormula = "N/A";
    if (j === 0) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&O${row})`;
    } else if (j === 1) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&P${row})`;
    } else if (j === 2) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&Q${row})`;
    } else if (j === 3) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&R${row})`;
    } else if (j === 4) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&S${row})`;
    } else if (j === 5) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&T${row})`;
    } else if (j === 6) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&U${row})`;
    } else if (j === 7) {
      imageFormula = `=IMAGE("https://drive.google.com/uc?id="&V${row})`;
    }
    sheet.getRange(row, 23 + j).setValue(imageFormula); // W～AD列：イメージ①～⑧
  }
}
/**
 * バリエーション情報に基づき、データを処理するメイン関数
 */
function processMakeDataWithVariations() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODUCTS);
  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);

  const lastRow = sheet.getLastRow();
  const keepaApiKey = settingsSheet.getRange("C4").getValue(); // Keepa APIキー
  const folderId = settingsSheet.getRange("C5").getValue(); // Google Drive フォルダID

  for (let row = 2; row <= lastRow; row++) {
    const variationType = sheet.getRange(row, COL_VARIATION_TYPE).getValue(); // バリエーション選択列

    if (!variationType || variationType === "なし") continue;

    // バリエーション1とバリエーション2の値を取得
    const variation1Values = sheet
      .getRange(row, COL_VARIATION1_VALUES_START, 1, MAX_VARIATIONS)
      .getValues()
      .flat()
      .filter((v) => v);
    const variation2Values = sheet
      .getRange(row, COL_VARIATION2_VALUES_START, 1, MAX_VARIATIONS)
      .getValues()
      .flat()
      .filter((v) => v);

    let outputRow = row; // データ出力開始行
    for (let value1 of variation1Values) {
      for (let value2 of variation2Values) {
        const combination = `${value1}-${value2}`;
        const asin = sheet.getRange(outputRow, COL_ASIN).getValue(); // ASIN列

        if (!asin) {
          Logger.log(`ASINがありません。組み合わせ: ${combination}`);
          outputRow++;
          continue;
        }

        // Keepa APIで商品情報を取得
        const productData = fetchKeepaProductData(keepaApiKey, asin);
        if (!productData) {
          Logger.log(`Keepa APIからデータを取得できません。ASIN: ${asin}`);
          outputRow++;
          continue;
        }

        // データをスプレッドシートに書き込み
        sheet.getRange(outputRow, COL_OUTPUT_START).setValue(combination); // 組み合わせ
        sheet.getRange(outputRow, COL_ASIN).setValue(asin); // ASIN
        sheet
          .getRange(outputRow, COL_PRICE)
          .setValue(productData.price || "N/A"); // 新品最安値
        sheet
          .getRange(outputRow, COL_WEIGHT)
          .setValue(productData.weight || "N/A"); // 重量
        sheet
          .getRange(outputRow, COL_LENGTH)
          .setValue(productData.length || "N/A"); // 長さ
        sheet
          .getRange(outputRow, COL_WIDTH)
          .setValue(productData.width || "N/A"); // 幅
        sheet
          .getRange(outputRow, COL_HEIGHT)
          .setValue(productData.height || "N/A"); // 高さ
        sheet
          .getRange(outputRow, COL_TITLE_EN)
          .setValue(productData.title || "N/A"); // 商品名 (英)
        sheet
          .getRange(outputRow, COL_TITLE_ZH)
          .setValue(productData.title_zh || "N/A"); // 商品名 (中)

        // 画像情報をGoogle Driveに保存して取得
        const imageIds = [];
        if (productData.image) {
          const images = productData.image.split(",");
          for (let i = 0; i < images.length && i < 8; i++) {
            const fileId = downloadImageAndSave(images[i], folderId);
            imageIds.push(fileId || "N/A");
          }
        }
        sheet.getRange(outputRow, COL_IMAGE1).setValue(imageIds[0] || "N/A"); // 画像①
        sheet
          .getRange(outputRow, COL_IMAGE1_DISPLAY)
          .setFormula(
            `=IMAGE("https://drive.google.com/uc?id="&CS${outputRow})`
          ); // イメージ①

        // SKUや価格はカスタマイズ可能
        sheet
          .getRange(outputRow, COL_SKU)
          .setValue(`SKU-${asin}-${value1}-${value2}`); // SKU
        sheet.getRange(outputRow, COL_STOCK).setValue(100); // 在庫数

        outputRow++;
      }
    }
  }
}

/**
 * バリエーションデータを取得する関数
 * @param {Object} sheet - スプレッドシートオブジェクト
 * @param {number} row - 対象行番号
 * @param {number} nameCol - バリエーション名称の列番号
 * @param {number} startCol - バリエーション値の開始列番号
 * @param {number} endCol - バリエーション値の終了列番号
 * @return {Object} バリエーション名と値を含むオブジェクト
 */
// function getVariationData(sheet, row, nameCol, startCol, endCol) {
//   const name = sheet.getRange(row, nameCol).getValue(); // バリエーション名
//   const values = [];
//   for (let col = startCol; col <= endCol; col++) {
//     const value = sheet.getRange(row, col).getValue(); // 各列の値を取得
//     if (value) values.push(value); // 値が存在する場合のみ追加
//   }
//   return { name, values };
// }

/**
 * バリエーションなしの商品を処理する関数
 * @param {number} row - 処理対象の行番号
 */
function processMainProduct(row) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
  const itemName = sheet.getRange(row, 44).getValue(); // 商品名
  Logger.log(`バリエーションなしの商品: ${itemName}`);
  // ここに商品登録ロジックを追加
}

// /**
//  * バリエーションが1つの商品を処理する関数
//  * @param {number} row - 処理対象の行番号
//  * @param {Object} variation1 - 1つ目のバリエーションデータ
//  */
// function processProductWithVariation1(row, variation1) {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
//   const itemName = sheet.getRange(row, 44).getValue(); // 商品名
//   const combinations = generateVariations(variation1);

//   // BL列以降にバリエーションの組み合わせを出力
//   let colIndex = 64; // BL列 (スプレッドシートは1始まり)
//   combinations.forEach(combination => {
//     sheet.getRange(row, colIndex).setValue(`${variation1.name}: ${combination[variation1.name]}`);
//     colIndex++;
//   });

//   Logger.log(`バリエーション1: ${itemName}, ${variation1.name}, ${variation1.values}`);
// }

// /**
//  * バリエーションが2つの商品を処理する関数
//  * @param {number} row - 処理対象の行番号
//  * @param {Object} variation1 - 1つ目のバリエーションデータ
//  * @param {Object} variation2 - 2つ目のバリエーションデータ
//  */
// function processProductWithVariation2(row, variation1, variation2) {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
//   const itemName = sheet.getRange(row, 44).getValue(); // 商品名
//   const combinations = generateVariations(variation1, variation2);

//   // BL列以降にバリエーションの組み合わせを出力
//   let colIndex = 64; // BL列 (スプレッドシートは1始まり)
//   combinations.forEach(combination => {
//     sheet.getRange(row, colIndex).setValue(
//       `${variation1.name}: ${combination[variation1.name]}, ${variation2.name}: ${combination[variation2.name]}`
//     );
//     colIndex++;
//   });

//   Logger.log(`バリエーション2: ${itemName}, ${variation1.name}, ${variation1.values}, ${variation2.name}, ${variation2.values}`);
// }

/**
 * バリエーションを生成する関数
 * @param {Object} variation1 - 1つ目のバリエーションデータ
 * @param {Object} variation2 - 2つ目のバリエーションデータ（省略可能）
 * @return {Array} 生成されたバリエーションデータ
 */
// function generateVariations(variation1, variation2 = null) {
//   const variations = [];

//   if (!variation2) {
//     variation1.values.forEach(value => {
//       variations.push({ [variation1.name]: value });
//     });
//   } else {
//     variation1.values.forEach(value1 => {
//       variation2.values.forEach(value2 => {
//         variations.push({ [variation1.name]: value1, [variation2.name]: value2 });
//       });
//     });
//   }

//   return variations;
// }

/**
 * 生成されたバリエーションデータをテストする関数
 */
function testGenerateVariations() {
  const variation1 = { name: "Color", values: ["Red", "Blue", "Green"] };
  const variation2 = { name: "Size", values: ["S", "M", "L"] };

  const oneVariation = generateVariations(variation1);
  Logger.log(`バリエーション1: ${JSON.stringify(oneVariation)}`);

  const twoVariations = generateVariations(variation1, variation2);
  Logger.log(`バリエーション2: ${JSON.stringify(twoVariations)}`);
}

/**
 * バリエーション選択用のプルダウンを作成する
 */

function setDropdownForVariationSelection() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
  const range = sheet.getRange("BK2:BK100"); // BK列の2行目から100行目に適用
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["なし", "1つ", "2つ"], true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

// シグネチャ生成関数
function generateSign(
  path,
  partnerId,
  shopId,
  accessToken,
  timestamp,
  partnerKey
) {
  const baseString = `${partnerId}${path}${timestamp}${accessToken}${shopId}`;
  const signatureBytes = Utilities.computeHmacSha256Signature(
    baseString,
    partnerKey
  );
  const sign = signatureBytes
    .map((byte) => {
      const v = (byte < 0 ? byte + 256 : byte).toString(16);
      return v.length === 1 ? "0" + v : v;
    })
    .join("");
  return sign;
}

function initTierVariation(
  shop_id,
  access_token,
  item_id,
  tier_variation,
  model
) {
  const partner_id = PARTNER_ID; // Shopee Partner ID
  const partner_key = PARTNER_KEY; // Shopee Partner Key
  const timestamp = Math.floor(Date.now() / 1000); // 現在のタイムスタンプ
  const path = "/api/v2/product/init_tier_variation";

  // シグネチャを生成
  const sign = generateSign(
    path,
    partner_id,
    shop_id,
    access_token,
    timestamp,
    partner_key
  );

  // URLの構築
  const url = `https://partner.shopeemobile.com${path}?partner_id=${partner_id}&shop_id=${shop_id}&timestamp=${timestamp}&sign=${sign}&access_token=${access_token}`;

  // `item_id` の型を確認し、数値型に変換
  if (typeof item_id !== "number") {
    item_id = parseInt(item_id, 10); // 数値型に変換
    if (isNaN(item_id)) {
      throw new Error("Invalid item_id: item_id must be a valid number.");
    }
  }

  // ペイロード
  const payload = {
    item_id: item_id,
    tier_variation: tier_variation,
    model: model.map((m) => ({
      tier_index: m.tier_index,
      original_price: m.original_price,
      normal_stock: m.normal_stock,
      model_sku: m.model_sku,
      weight: m.weight, // バリエーションごとの重量
      dimension: {
        // バリエーションごとの梱包サイズ
        package_length: Math.round(m.dimension.package_length || 1),
        package_width: Math.round(m.dimension.package_width || 1),
        package_height: Math.round(m.dimension.package_height || 1),
      },
      seller_stock: [
        {
          stock: m.normal_stock, // 必須フィールド: 在庫数を指定
        },
      ],
    })),
  };

  // リクエストオプション
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
  Logger.log("initTierVariation Payload: " + JSON.stringify(payload, null, 2)); // JSON文字列化してログ出力

  try {
    // APIを呼び出す
    const response = UrlFetchApp.fetch(url, options);

    // レスポンスをパース
    const responseData = JSON.parse(response.getContentText());

    // エラーが含まれている場合は例外をスロー
    if (responseData.error) {
      throw new Error(`Error in initTierVariation: ${responseData.message}`);
    }

    Logger.log(`initTierVariation response: ${JSON.stringify(responseData)}`);
    return responseData;
  } catch (error) {
    // エラーをログに出力して再スロー
    Logger.log(`Error in initTierVariation: ${error.message}`);
    throw error;
  }
}

/**
 * 配送情報を追加
 */
function getValidLogistics(shop_id, access_token) {
  const partner_id = PARTNER_ID; // 正しいPartner IDを設定
  const partner_key = PARTNER_KEY; // 正しいPartner Keyを設定
  const path = "/api/v2/logistics/get_channel_list";
  const timestamp = Math.floor(Date.now() / 1000);

  // シグネチャを生成
  const sign = generateSign(
    path,
    partner_id,
    shop_id,
    access_token,
    timestamp,
    partner_key
  );

  // クエリパラメータ付きのURLを作成
  const url = `https://partner.shopeemobile.com${path}?partner_id=${partner_id}&shop_id=${shop_id}&timestamp=${timestamp}&sign=${sign}&access_token=${access_token}`;

  const options = {
    method: "get",
    contentType: "application/json",
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const logisticsData = JSON.parse(response.getContentText());

    Logger.log(`Response: ${JSON.stringify(logisticsData)}`);

    if (logisticsData.error) {
      throw new Error(`Error fetching logistics: ${logisticsData.message}`);
    }

    // 有効な配送チャネルを抽出
    const validLogistics = logisticsData.response.logistics_channel_list
      .filter((channel) => channel.enabled) // 有効なチャネルのみ
      .map((channel) => ({
        logistic_id: channel.logistics_channel_id,
        enabled: true,
        is_free: false, // 配送料はバイヤー負担
      }));

    if (validLogistics.length === 0) {
      throw new Error("No valid logistics channels found.");
    }

    return validLogistics;
  } catch (error) {
    Logger.log(`Error in getValidLogistics: ${error.message}`);
    throw error;
  }
}

/**
 * Step1では、各行のデータを取得し、バリエーション情報を整理して返します。
 */
function prepareVariationData(row) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODUCTS);
  const data = []; // データ格納用の配列

  const variationType = sheet.getRange(row, COL_VARIATION_TYPE).getValue(); // BK列（バリエーション選択）

  if (!variationType || variationType === "なし") {
    // バリエーションなしの場合、データを追加
    data.push({
      row,
      variationType: "なし",
      mainProduct: true, // メイン商品のフラグ
      variation1: null,
      variation2: null,
    });
  }

  // バリエーション情報を取得
  const variation1 = getVariationData(
    sheet,
    row,
    COL_VARIATION1_NAME,
    COL_VARIATION1_VALUES_START,
    COL_VARIATION1_VALUES_END
  );
  const variation2 =
    variationType === "2つ"
      ? getVariationData(
          sheet,
          row,
          COL_VARIATION2_NAME,
          COL_VARIATION2_VALUES_START,
          COL_VARIATION2_VALUES_END
        )
      : null;

  // データ配列に追加
  data.push({
    row,
    variationType,
    mainProduct: false, // バリエーション商品のフラグ
    variation1,
    variation2,
  });

  return data; // 整理されたデータを返す
}

/**
 * Step2では、Step1で準備したデータを基に処理を実行します。
 */
function processPreparedData(preparedData) {
  preparedData.forEach((data) => {
    if (data.mainProduct) {
      // メイン商品のみを処理
      processMainProduct(data.row);
    } else if (data.variationType === "1つ") {
      // バリエーションが1つの場合を処理
      processProductWithVariation1(data.row, data.variation1);
    } else if (data.variationType === "2つ") {
      // バリエーションが2つの場合を処理
      processProductWithVariation2(data.row, data.variation1, data.variation2);
    }
  });
}
/**
 * バリエーションがない商品の処理（例としてログを出力）。
 */
function processMainProduct(row) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
  const itemName = sheet.getRange(row, COL_MAIN_PRODUCT_NAME_JP).getValue(); // 商品名
  Logger.log(`バリエーションなしの商品を処理します: ${itemName}`);
}

/**
 * バリエーションが1つの場合の処理。
 */
function processProductWithVariation1(row, variation1) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
  const itemName = sheet.getRange(row, COL_MAIN_PRODUCT_NAME_JP).getValue(); // 商品名

  // バリエーションの組み合わせを生成
  const combinations = generateVariations(variation1);

  // BL列以降に出力
  let baseColIndex = COL_VARIATION_OUTPUT_START; // CJ列スタート
  combinations.forEach((combination, index) => {
    const currentColIndex = baseColIndex + index * VARIATION_DATAS; // 現在の列位置
    sheet
      .getRange(row, currentColIndex)
      .setValue(`${variation1.name}: ${combination[variation1.name]}`);
  });

  Logger.log(
    `バリエーション1の商品を処理しました: ${itemName}, ${variation1.name}, ${variation1.values}`
  );
}
/**
 * バリエーションが2つの場合の処理。
 */
function processProductWithVariation2(row, variation1, variation2) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
  const itemName = sheet.getRange(row, COL_MAIN_PRODUCT_NAME_JP).getValue(); // 商品名

  // バリエーションの組み合わせを生成
  const combinations = generateVariations(variation1, variation2);

  // CJ列以降に25列ずつ間隔を空けて出力
  let baseColIndex = COL_VARIATION_OUTPUT_START; // CJ列スタート
  combinations.forEach((combination, index) => {
    const currentColIndex = baseColIndex + index * VARIATION_DATAS; // 現在の列位置

    // バリエーション情報を出力
    sheet
      .getRange(row, currentColIndex)
      .setValue(
        `${variation1.name}: ${combination[variation1.name]}, ${
          variation2.name
        }: ${combination[variation2.name]}`
      );
  });

  Logger.log(
    `バリエーション2の商品を処理しました: ${itemName}, ${variation1.name}, ${variation1.values}, ${variation2.name}, ${variation2.values}`
  );
}

/**
 * バリエーションの組み合わせを生成する関数。
 */
function generateVariations(variation1, variation2 = null) {
  const variations = [];

  if (!variation2) {
    variation1.values.forEach((value) => {
      variations.push({ [variation1.name]: value });
    });
  } else {
    variation1.values.forEach((value1) => {
      variation2.values.forEach((value2) => {
        variations.push({
          [variation1.name]: value1,
          [variation2.name]: value2,
        });
      });
    });
  }

  return variations;
}

/**
 * ASINを引数として受け取り、Keepa APIから商品データを取得。
 * 商品情報をオブジェクト形式で返します。
 */
function fetchProductDataFromKeepa(asin) {
  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);

  // 基本設定シートからAPIキーとフォルダIDを取得
  const apiKey = settingsSheet.getRange("C4").getValue();
  const folderId = settingsSheet.getRange("C5").getValue();

  if (!apiKey || !folderId) {
    Logger.log("APIキーまたはフォルダIDが見つかりません。");
    return null;
  }

  // Keepa APIから商品情報を取得
  let productData = postKeepaProducts(apiKey, asin);
  if (!productData) {
    Logger.log("APIトークンが無くなりました。");
    return null;
  }

  // JSON文字列をオブジェクトに変換
  productData = JSON.parse(productData);

  // 画像をダウンロードしてGoogle Driveに保存
  const imageIds = [];
  if (productData.image) {
    const images = productData.image.split(",");
    Logger.log(images);
    for (let j = 0; j < images.length && j < 8; j++) {
      const fileId = downloadImageAndSave(images[j], folderId);
      imageIds.push(fileId || "N/A");
    }
  }

  // 商品データを返却
  return {
    asin: asin,
    imageIds: imageIds,
    price: productData.price || "N/A",
    weight: productData.weight ? productData.weight / 1000 : "N/A",
    length: productData.length ? productData.length / 10 : "N/A",
    width: productData.width ? productData.width / 10 : "N/A",
    height: productData.height ? productData.height / 10 : "N/A",
    title: productData.title || "N/A",
    features: productData.features
      ? productData.features.join("\n").replace(/\n+/g, "\n")
      : "",
    description: productData.description
      ? productData.description.replace(/\n+/g, "\n")
      : "",
  };
}

/**
 * メイン商品のデータを指定された行に書き込みます。
 * 翻訳用の数式や画像表示も自動設定します。
 */
function writeMainProductDataToSheet(sheet, row, data) {
  const currentDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy/MM/dd"
  );

  sheet.getRange(row, COL_DATA_DATE).setValue(currentDate); // データ取得日
  sheet.getRange(row, COL_MAIN_PRICE).setValue(data.price); // 新品最安値
  sheet.getRange(row, COL_MAIN_WEIGHT).setValue(data.weight); // 重量 (kg)
  sheet.getRange(row, COL_MAIN_LENGTH).setValue(data.length); // 長さ (cm)
  sheet.getRange(row, COL_MAIN_WIDTH).setValue(data.width); // 幅 (cm)
  sheet.getRange(row, COL_MAIN_HEIGHT).setValue(data.height); // 高さ (cm)
  sheet.getRange(row, COL_MAIN_PRODUCT_NAME_JP).setValue(data.title); // 商品名

  // 商品説明
  const combinedDescription = `${data.features}\n${data.description}`;
  sheet.getRange(row, COL_MAIN_DESCRIPTION_JP).setValue(combinedDescription);

  // 翻訳用の数式をセット
  sheet
    .getRange(row, COL_MAIN_PRODUCT_NAME_EN)
    .setFormula(`=GOOGLETRANSLATE(I${row}, "ja", "en")`); // 商品名 (英)
  sheet
    .getRange(row, COL_MAIN_DESCRIPTION_EN)
    .setFormula(`=GOOGLETRANSLATE(J${row}, "ja", "en")`); // 商品説明 (英)
  sheet
    .getRange(row, COL_MAIN_PRODUCT_NAME_ZH)
    .setFormula(`=GOOGLETRANSLATE(I${row}, "ja", "zh-CN")`); // 商品名 (中)
  sheet
    .getRange(row, COL_MAIN_DESCRIPTION_ZH)
    .setFormula(`=GOOGLETRANSLATE(J${row}, "ja", "zh-CN")`); // 商品説明 (中)

  // 画像ファイルIDをプロット
  for (let j = 0; j < 8; j++) {
    const imageId = data.imageIds[j] || "N/A";
    sheet.getRange(row, COL_MAIN_IMAGE_1 + j).setValue(imageId); // 画像列
    if (imageId !== "N/A") {
      sheet
        .getRange(row, COL_MAIN_IMAGE_1_DISPLAY + j)
        .setFormula(
          `=IMAGE("https://drive.google.com/uc?id="&${
            COL_MAIN_IMAGE_1 + j
          }${row})`
        );
    }
  }
}

/**
 * バリエーション商品のデータを指定された列に書き込みます。
 * 画像表示や商品情報も列単位で処理します。
 */
function writeVariationProductDataToSheet(sheet, row, col, data) {
  // 商品情報
  sheet.getRange(row, col + VARIATION_PRICE).setValue(data.price); // 新品最安値
  sheet.getRange(row, col + VARIATION_WEIGHT).setValue(data.weight); // 重量 (kg)
  sheet.getRange(row, col + VARIATION_LENGTH).setValue(data.length); // 長さ (cm)
  sheet.getRange(row, col + VARIATION_WIDTH).setValue(data.width); // 幅 (cm)
  sheet.getRange(row, col + VARIATION_HEIGHT).setValue(data.height); // 高さ (cm)
  sheet.getRange(row, col + VARIATION_NAME).setValue(data.title); // 商品名(日)
  sheet
    .getRange(row, col + VATIATION_DESCRIPTION_JP)
    .setValue(data.description); // 商品説明
  // 画像
  // for (let j = 0; j < data.imageIds.length && j < 8; j++) {
  //   const imageId = data.imageIds[j] || "N/A";
  //   sheet.getRange(row, col + j).setValue(imageId);
  //   if (imageId !== "N/A") {
  //     sheet.getRange(row, col + 8 + j).setFormula(`=IMAGE("https://drive.google.com/uc?id="&${col + j}${row})`);
  //   }
  // }
  // バリエーション品はトップ画像のみ
  const imageId = data.imageIds[0] || "N/A";
  sheet.getRange(row, col + VARIATION_IMAGE_ID).setValue(imageId);
  if (imageId !== "N/A") {
    sheet
      .getRange(row, col + VARIATION_IMAGE)
      .setFormula(`=IMAGE("https://drive.google.com/uc?id="&${col}${row})`);
  }
}

/**
 * 各国の出品価格を計算
 */
function calculatePricesForCountries(sheet, row, col, data) {
  const rates = getExchangeRates(); // 各国通貨レート取得
  const basePrice = data.price; // ベース価格

  // 各国の価格を計算して書き込み
  sheet.getRange(row, col + 29).setValue(basePrice * rates.SG); // シンガポール
  sheet.getRange(row, col + 30).setValue(basePrice * rates.MY); // マレーシア
  sheet.getRange(row, col + 31).setValue(basePrice * rates.PH); // フィリピン
  sheet.getRange(row, col + 32).setValue(basePrice * rates.TH); // タイ
  sheet.getRange(row, col + 33).setValue(basePrice * rates.TW); // 台湾
  sheet.getRange(row, col + 34).setValue(basePrice * rates.VN); // ベトナム
  sheet.getRange(row, col + 35).setValue(basePrice * rates.BR); // ブラジル
}

/**
 * Step1はバリエーション組み合わせを生成する
 */
function processMakeDataStep1() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");

  const lastRow = sheet.getLastRow();
  const checkRange = sheet.getRange(2, COL_CHECKBOX, lastRow - 1); // チェックボックス列の範囲
  const asinRange = sheet.getRange(2, COL_ASIN, lastRow - 1); // メインASIN列の範囲
  const checkValues = checkRange.getValues(); // チェックボックス状態
  const asinValues = asinRange.getValues(); // ASIN

  const rowsToProcess = [];

  // チェックされた行をリストに追加
  for (let i = 0; i < checkValues.length; i++) {
    const isChecked = checkValues[i][0]; // チェックボックス
    const asin = asinValues[i][0]; // ASIN

    if (isChecked && asin) {
      rowsToProcess.push(i + 2); // 行番号を保存 (スプレッドシート上では2行目から)
    }
  }

  // 処理対象が無ければ終了
  if (rowsToProcess.length === 0) {
    Logger.log("処理対象の行がありません。");
    return;
  }

  // 各行に対して1回だけ実行する処理
  rowsToProcess.forEach(function (row) {
    // 各行ごとに1回だけ実行する関数
    // Step1-1: 必要なデータを取得・準備
    const preparedData = prepareVariationData(row);

    // Step1-2: データを使って処理を実行
    processPreparedData(preparedData);
  });
}

/**
 * 出品情報を作成する(brand情報を除く)
 * バリエーション名をユーザがつける場合
 */
function processMakeDataStep2() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODUCTS);
  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("基本設定");
  const additionalSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("追加設定");
  const lastRow = sheet.getLastRow();
  const checkRange = sheet.getRange(2, COL_CHECKBOX, lastRow - 1); // チェックボックス列の範囲
  const asinRange = sheet.getRange(2, COL_ASIN, lastRow - 1); // メインASIN列の範囲
  const checkValues = checkRange.getValues(); // チェックボックス状態
  const asinValues = asinRange.getValues(); // ASIN

  // 基本設定シートからAPIキーとフォルダIDを取得
  const apiKey = settingsSheet.getRange("C4").getValue();
  // const apiKey = PropertiesService.getScriptProperties().getProperty('KEEPA_API_KEY');
  const folderId = settingsSheet.getRange("C5").getValue();

  const rowsToProcess = [];

  // チェックされた行をリストに追加
  for (let i = 0; i < checkValues.length; i++) {
    const isChecked = checkValues[i][0]; // チェックボックス
    const asin = asinValues[i][0]; // ASIN

    if (isChecked && asin) {
      rowsToProcess.push(i + 2); // 行番号を保存 (スプレッドシート上では2行目から)
    }
  }

  // 処理対象が無ければ終了
  if (rowsToProcess.length === 0) {
    Logger.log("処理対象の行がありません。");
    return;
  }

  // 各行に対して1回だけ実行する処理
  rowsToProcess.forEach(function (row) {
    if (!sheet || !apiKey || !folderId) {
      Logger.log("出品一覧シート、APIキー、またはフォルダIDが見つかりません。");
      return;
    }

    const asin = sheet.getRange(row, COL_ASIN).getValue(); // メインASIN

    if (!asin) {
      Logger.log(`行${row}にASINがありません。`);
      return;
    }

    // 処理開始前にC列～AD列をクリア
    sheet.getRange(row, COL_DATA_DATE, 1, VARIATION_DATAS).clearContent(); // C列からAD列をクリア

    let productData = postKeepaProducts(apiKey, asin);
    if (!productData) {
      Logger.log("APIトークンが無くなりました。");
      return;
    }

    // productDataがJSON文字列なので、オブジェクトに変換
    productData = JSON.parse(productData);

    // データが正しく取得できたか確認
    Logger.log(JSON.stringify(productData)); // ログでデータ確認用

    // 画像をダウンロードしてGoogle Driveに保存
    const imageIds = [];
    if (productData.image) {
      const images = productData.image.split(",");
      Logger.log(images);
      for (let j = 0; j < images.length && j < 8; j++) {
        const fileId = downloadImageAndSave(images[j], folderId);
        imageIds.push(fileId || "N/A");
      }
    }

    // シートにプロット
    const currentDate = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy/MM/dd"
    );
    sheet.getRange(row, 3).setValue(currentDate); // C列：データ取得日
    sheet.getRange(row, 4).setValue(productData.price || "N/A"); // D列：新品最安値
    sheet
      .getRange(row, 5)
      .setValue(productData.weight ? productData.weight / 1000 : "N/A"); // E列：重量 (kg)
    sheet
      .getRange(row, 6)
      .setValue(productData.length ? productData.length / 10 : "N/A"); // F列：長さ (cm)
    sheet
      .getRange(row, 7)
      .setValue(productData.width ? productData.width / 10 : "N/A"); // G列：幅 (cm)
    sheet
      .getRange(row, 8)
      .setValue(productData.height ? productData.height / 10 : "N/A"); // H列：高さ (cm)
    sheet.getRange(row, 9).setValue(productData.title || "N/A"); // I列：商品名

    // featuresとdescriptionの結合（改行を減らす）
    const features = productData.features
      ? productData.features.join("\n").replace(/\n+/g, "\n")
      : "";
    const description = productData.description
      ? productData.description.replace(/\n+/g, "\n")
      : "";
    sheet.getRange(row, 10).setValue(features + "\n" + description || "N/A"); // J列：商品説明

    // 翻訳用の数式をセット
    sheet.getRange(row, 11).setFormula(`=GOOGLETRANSLATE(I${row}, "ja", "en")`); // K列：商品名 (英)
    sheet.getRange(row, 12).setFormula(`=GOOGLETRANSLATE(J${row}, "ja", "en")`); // L列：商品説明 (英)
    sheet
      .getRange(row, 13)
      .setFormula(`=GOOGLETRANSLATE(I${row}, "ja", "zh-CN")`); // M列：商品名 (中)
    sheet
      .getRange(row, 14)
      .setFormula(`=GOOGLETRANSLATE(J${row}, "ja", "zh-CN")`); // N列：商品説明 (中)

    // 画像ファイルIDをプロット、N/Aの場合はイメージ欄もN/Aにする
    for (let j = 0; j < 8; j++) {
      const imageId = imageIds[j] || "N/A";
      sheet.getRange(row, 15 + j).setValue(imageId); // O～V列：画像①～⑧
      var imageFormula = "N/A";
      if (j === 0) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&O${row})`;
      } else if (j === 1) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&P${row})`;
      } else if (j === 2) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&Q${row})`;
      } else if (j === 3) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&R${row})`;
      } else if (j === 4) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&S${row})`;
      } else if (j === 5) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&T${row})`;
      } else if (j === 6) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&U${row})`;
      } else if (j === 7) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&V${row})`;
      }
      sheet.getRange(row, 23 + j).setValue(imageFormula); // W～AD列：イメージ①～⑧
    }

    // 各シートから必要な情報を取得
    const initWeight = parseFloat(settingsSheet.getRange("C7").getValue()); // 初期重量（未取得時）
    const addWeight = parseFloat(settingsSheet.getRange("C8").getValue()); // 重量加算
    const initStock = parseFloat(settingsSheet.getRange("C9").getValue()); // 初期在庫

    const rateMap = {}; // 各国レート
    rateMap["SG"] = parseFloat(settingsSheet.getRange("F22").getValue()); // SGレート
    rateMap["MY"] = parseFloat(settingsSheet.getRange("F23").getValue()); // MYレート
    rateMap["PH"] = parseFloat(settingsSheet.getRange("F24").getValue()); // PHレート
    rateMap["TH"] = parseFloat(settingsSheet.getRange("F25").getValue()); // THレート
    rateMap["TW"] = parseFloat(settingsSheet.getRange("F26").getValue()); // TWレート
    rateMap["VN"] = parseFloat(settingsSheet.getRange("F27").getValue()); // VNレート
    rateMap["BR"] = parseFloat(settingsSheet.getRange("F28").getValue()); // VNレート

    const productWeight =
      sheet.getRange(row, 5).getValue() === "N/A"
        ? initWeight
        : parseFloat(sheet.getRange(row, 5).getValue()) + addWeight; // 重量 (kg)

    const priceMap = priceCalc(row);
    // プロット前にクリア
    sheet.getRange(row, 31, 1, 22).clearContent();

    // シートにプロット
    // sheet.getRange(row, 31).setValue(`${asin}-${currentDate}`); // AE列：SKU
    sheet.getRange(row, 31).setValue(`${asin}`); // AE列：SKU
    sheet.getRange(row, 32).setValue(priceMap["SG"]);
    sheet.getRange(row, 33).setValue(priceMap["MY"]);
    sheet.getRange(row, 34).setValue(priceMap["PH"]);
    sheet.getRange(row, 35).setValue(priceMap["TH"]);
    sheet.getRange(row, 36).setValue(priceMap["TW"]);
    sheet.getRange(row, 37).setValue(priceMap["VN"]);
    sheet.getRange(row, 38).setValue(priceMap["BR"]);
    sheet
      .getRange(row, 39)
      .setValue(Math.round(priceMap["SG"] * rateMap["SG"]));
    sheet
      .getRange(row, 40)
      .setValue(Math.round(priceMap["MY"] * rateMap["MY"]));
    sheet
      .getRange(row, 41)
      .setValue(Math.round(priceMap["PH"] * rateMap["PH"]));
    sheet
      .getRange(row, 42)
      .setValue(Math.round(priceMap["TH"] * rateMap["TH"]));
    sheet
      .getRange(row, 43)
      .setValue(Math.round(priceMap["TW"] * rateMap["TW"]));
    sheet
      .getRange(row, 44)
      .setValue(Math.round(priceMap["VN"] * rateMap["VN"]));
    sheet.getRange(row, 50).setValue(productWeight);
    sheet
      .getRange(row, 51)
      .setValue(
        sheet.getRange(row, 6).getValue() === "N/A"
          ? settingsSheet.getRange("I7").getValue()
          : sheet.getRange(row, 6).getValue()
      );
    sheet
      .getRange(row, 52)
      .setValue(
        sheet.getRange(row, 7).getValue() === "N/A"
          ? settingsSheet.getRange("I8").getValue()
          : sheet.getRange(row, 7).getValue()
      );
    sheet
      .getRange(row, 53)
      .setValue(
        sheet.getRange(row, 8).getValue() === "N/A"
          ? settingsSheet.getRange("I9").getValue()
          : sheet.getRange(row, 8).getValue()
      );
    sheet.getRange(row, 54).setValue(initStock);

    //////////////////////////////////////////////
    const keepaBrand = productData.brand || "N/A";
    sheet.getRange(row, COL_BRAND_FROM_KEEPA).setValue(keepaBrand); // Keepaから取得したブランドを格納

    let colStart = COL_VARIATION_OUTPUT_START; // CK列 (バリエーションASIN開始列)
    const variationDataList = [];

    while (sheet.getRange(row, colStart).getValue()) {
      const variationAsin = sheet.getRange(row, colStart + 1).getValue();
      if (variationAsin != "") {
        // プロット前にクリア
        sheet
          .getRange(row, colStart + 2, 1, VARIATION_DATAS - 2)
          .clearContent();
        const variationData = fetchProductDataFromKeepa(variationAsin); // バリエーションデータ取得

        if (variationData) {
          writeVariationProductDataToSheet(sheet, row, colStart, variationData);

          // 各国の出品価格を計算
          const priceMap = priceCalcVariation(row, colStart);

          // シートにプロット
          sheet
            .getRange(row, colStart + VARIATION_SKU)
            .setValue(`${variationAsin}`); // AE列：SKU
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG)
            .setValue(priceMap["SG"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 1)
            .setValue(priceMap["MY"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 2)
            .setValue(priceMap["PH"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 3)
            .setValue(priceMap["TH"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 4)
            .setValue(priceMap["TW"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 5)
            .setValue(priceMap["VN"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 6)
            .setValue(priceMap["BR"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN)
            .setValue(Math.round(priceMap["SG"] * rateMap["SG"]));
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN + 1)
            .setValue(Math.round(priceMap["MY"] * rateMap["MY"]));
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN + 2)
            .setValue(Math.round(priceMap["PH"] * rateMap["PH"]));
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN + 3)
            .setValue(Math.round(priceMap["TH"] * rateMap["TH"]));
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN + 4)
            .setValue(Math.round(priceMap["TW"] * rateMap["TW"]));
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN + 5)
            .setValue(Math.round(priceMap["VN"] * rateMap["VN"]));
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN + 6)
            .setValue(Math.round(priceMap["BR"] * rateMap["BR"]));

          sheet
            .getRange(row, colStart + VARIATION_WEIGHT_CON)
            .setValue(productWeight);
          sheet
            .getRange(row, colStart + VARIATION_LENGTH_CON)
            .setValue(
              sheet.getRange(row, 6).getValue() === "N/A"
                ? settingsSheet.getRange("I7").getValue()
                : sheet.getRange(row, 6).getValue()
            );
          sheet
            .getRange(row, colStart + VARIATION_WIDTH_CON)
            .setValue(
              sheet.getRange(row, 7).getValue() === "N/A"
                ? settingsSheet.getRange("I8").getValue()
                : sheet.getRange(row, 7).getValue()
            );
          sheet
            .getRange(row, colStart + VARIATION_HEIGHT_CON)
            .setValue(
              sheet.getRange(row, 8).getValue() === "N/A"
                ? settingsSheet.getRange("I9").getValue()
                : sheet.getRange(row, 8).getValue()
            );
          sheet
            .getRange(row, colStart + VARIATION_ZAIKO_CON)
            .setValue(initStock);

          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG)
            .setValue(priceMap["SG"]); // SG価格
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 1)
            .setValue(priceMap["MY"]); // MY価格
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 2)
            .setValue(priceMap["PH"]); // PH価格
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 3)
            .setValue(priceMap["TH"]); // TH価格
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 4)
            .setValue(priceMap["TW"]); // TW価格
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 5)
            .setValue(priceMap["VN"]); // VN価格
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 6)
            .setValue(priceMap["BR"]); // VN価格

          // バリエーションデータを配列に追加
          variationDataList.push(variationData);
        }
      }

      colStart += VARIATION_DATAS; // 28列間隔で次のバリエーションASIN列に進む
    }

    // ChatGPTを使用してメイン商品名とメイン商品説明を生成
    const productName = generateMainProductNameWithChatGPT(variationDataList);
    const productDescription =
      generateMainProductDescriptionWithChatGPT(variationDataList);

    // 日本語、英語、簡体語でメイン商品名を出力
    sheet.getRange(row, COL_MAIN_PRODUCT_NAME_JP).setValue(productName.ja);
    // 英語の商品名を255文字以内にする
    const item_name_en_bef = additionalSheet.getRange("D6").getValue();
    const item_name_en_aft = additionalSheet.getRange("D7").getValue();
    var item_name_en = productName.en;
    if (item_name_en_bef) {
      item_name_en = item_name_en_bef + " " + item_name_en;
    }
    if (item_name_en_aft) {
      item_name_en = item_name_en + " " + item_name_en_aft;
    }
    const item_name_en_255 =
      item_name_en.length > 255 ? item_name_en.substring(0, 255) : item_name_en;

    // 英語の商品説明を3000文字以内にする
    const description_en_bef = additionalSheet.getRange("D9").getValue();
    const description_en_aft = additionalSheet.getRange("D10").getValue();
    var description_en = productDescription.en;
    if (description_en_bef) {
      description_en = description_en_bef + "\n" + description_en;
    }
    if (description_en_aft) {
      description_en = description_en + "\n" + description_en_aft;
    }
    const description_en_3000 =
      description_en.length > 3000
        ? description_en.substring(0, 3000)
        : description_en;

    // 中国語の商品名を60文字以内にする
    const item_name_ch_bef = additionalSheet.getRange("E6").getValue();
    const item_name_ch_aft = additionalSheet.getRange("E7").getValue();
    var item_name_ch = productName.zh;
    if (item_name_ch_bef) {
      item_name_ch = item_name_ch_bef + " " + item_name_ch;
    }
    if (item_name_ch_aft) {
      item_name_ch = item_name_ch + " " + item_name_ch_aft;
    }
    const item_name_ch_60 =
      item_name_ch.length > 60 ? item_name_ch.substring(0, 60) : item_name_ch;

    // 中国語の商品説明を3000文字以内にする
    const description_ch_bef = additionalSheet.getRange("E9").getValue();
    const description_ch_aft = additionalSheet.getRange("E10").getValue();
    var description_ch = productDescription.zh;
    if (description_ch_bef) {
      description_ch = description_ch_bef + "\n" + description_ch;
    }
    if (description_ch_aft) {
      description_ch = description_ch + "\n" + description_ch_aft;
    }
    const description_ch_3000 =
      description_ch.length > 3000
        ? description_ch.substring(0, 3000)
        : description_ch;

    sheet.getRange(row, COL_MAIN_PRODUCT_NAME_EN).setValue(item_name_en_255);
    sheet.getRange(row, COL_MAIN_PRODUCT_NAME_ZH).setValue(item_name_ch_60);

    // 日本語、英語、簡体語で商品説明を出力
    sheet
      .getRange(row, COL_MAIN_DESCRIPTION_JP)
      .setValue(productDescription.ja);
    sheet.getRange(row, COL_MAIN_DESCRIPTION_EN).setValue(description_en_3000);
    sheet.getRange(row, COL_MAIN_DESCRIPTION_ZH).setValue(description_ch_3000);

    sheet
      .getRange(row, COL_MAIN_PRODUCT_NAME_EN_CON)
      .setValue(item_name_en_255);
    sheet.getRange(row, COL_MAIN_PRODUCT_NAME_ZH_CON).setValue(item_name_ch_60);

    // 日本語、英語、簡体語で商品説明を出力
    sheet
      .getRange(row, COL_MAIN_DESCRIPTION_EN_CON)
      .setValue(description_en_3000);
    sheet
      .getRange(row, COL_MAIN_DESCRIPTION_ZH_CON)
      .setValue(description_ch_3000);
  });
}

function processMakeDataStep1to2_Step1ChatGPT() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODUCTS);
  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("基本設定");
  const additionalSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("追加設定");
  const lastRow = sheet.getLastRow();
  const checkRange = sheet.getRange(2, COL_CHECKBOX, lastRow - 1); // チェックボックス列の範囲
  const asinRange = sheet.getRange(2, COL_ASIN, lastRow - 1); // メインASIN列の範囲
  const checkValues = checkRange.getValues(); // チェックボックス状態
  const asinValues = asinRange.getValues(); // ASIN

  // 基本設定シートからAPIキーとフォルダIDを取得
  const apiKey = settingsSheet.getRange("C4").getValue();
  // const apiKey = PropertiesService.getScriptProperties().getProperty('KEEPA_API_KEY');
  const folderId = settingsSheet.getRange("C5").getValue();

  const rowsToProcess = [];

  // チェックされた行をリストに追加
  for (let i = 0; i < checkValues.length; i++) {
    const isChecked = checkValues[i][0]; // チェックボックス
    const asin = asinValues[i][0]; // ASIN

    if (isChecked && asin) {
      rowsToProcess.push(i + 2); // 行番号を保存 (スプレッドシート上では2行目から)
    }
  }

  // 処理対象が無ければ終了
  if (rowsToProcess.length === 0) {
    Logger.log("処理対象の行がありません。");
    return;
  }

  // 各行に対して1回だけ実行する処理
  rowsToProcess.forEach(function (row) {
    if (!sheet || !apiKey || !folderId) {
      Logger.log("出品一覧シート、APIキー、またはフォルダIDが見つかりません。");
      return;
    }

    const asin = sheet.getRange(row, COL_ASIN).getValue(); // メインASIN

    if (!asin) {
      Logger.log(`行${row}にASINがありません。`);
      return;
    }

    // 処理開始前にC列～AD列をクリア
    sheet.getRange(row, COL_DATA_DATE, 1, VARIATION_DATAS).clearContent(); // C列からAD列をクリア

    let productData = postKeepaProducts(apiKey, asin);
    if (!productData) {
      Logger.log("APIトークンが無くなりました。");
      return;
    }

    // productDataがJSON文字列なので、オブジェクトに変換
    productData = JSON.parse(productData);

    // データが正しく取得できたか確認
    Logger.log(JSON.stringify(productData)); // ログでデータ確認用

    // 画像をダウンロードしてGoogle Driveに保存
    const imageIds = [];
    if (productData.image) {
      const images = productData.image.split(",");
      Logger.log(images);
      for (let j = 0; j < images.length && j < 8; j++) {
        const fileId = downloadImageAndSave(images[j], folderId);
        imageIds.push(fileId || "N/A");
      }
    }

    // シートにプロット
    const currentDate = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy/MM/dd"
    );
    sheet.getRange(row, 3).setValue(currentDate); // C列：データ取得日
    sheet.getRange(row, 4).setValue(productData.price || "N/A"); // D列：新品最安値
    sheet
      .getRange(row, 5)
      .setValue(productData.weight ? productData.weight / 1000 : "N/A"); // E列：重量 (kg)
    sheet
      .getRange(row, 6)
      .setValue(productData.length ? productData.length / 10 : "N/A"); // F列：長さ (cm)
    sheet
      .getRange(row, 7)
      .setValue(productData.width ? productData.width / 10 : "N/A"); // G列：幅 (cm)
    sheet
      .getRange(row, 8)
      .setValue(productData.height ? productData.height / 10 : "N/A"); // H列：高さ (cm)
    sheet.getRange(row, 9).setValue(productData.title || "N/A"); // I列：商品名

    // featuresとdescriptionの結合（改行を減らす）
    const features = productData.features
      ? productData.features.join("\n").replace(/\n+/g, "\n")
      : "";
    const description = productData.description
      ? productData.description.replace(/\n+/g, "\n")
      : "";
    sheet.getRange(row, 10).setValue(features + "\n" + description || "N/A"); // J列：商品説明

    // 翻訳用の数式をセット
    sheet.getRange(row, 11).setFormula(`=GOOGLETRANSLATE(I${row}, "ja", "en")`); // K列：商品名 (英)
    sheet.getRange(row, 12).setFormula(`=GOOGLETRANSLATE(J${row}, "ja", "en")`); // L列：商品説明 (英)
    sheet
      .getRange(row, 13)
      .setFormula(`=GOOGLETRANSLATE(I${row}, "ja", "zh-CN")`); // M列：商品名 (中)
    sheet
      .getRange(row, 14)
      .setFormula(`=GOOGLETRANSLATE(J${row}, "ja", "zh-CN")`); // N列：商品説明 (中)

    // 画像ファイルIDをプロット、N/Aの場合はイメージ欄もN/Aにする
    for (let j = 0; j < 8; j++) {
      const imageId = imageIds[j] || "N/A";
      sheet.getRange(row, 15 + j).setValue(imageId); // O～V列：画像①～⑧
      var imageFormula = "N/A";
      if (j === 0) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&O${row})`;
      } else if (j === 1) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&P${row})`;
      } else if (j === 2) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&Q${row})`;
      } else if (j === 3) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&R${row})`;
      } else if (j === 4) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&S${row})`;
      } else if (j === 5) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&T${row})`;
      } else if (j === 6) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&U${row})`;
      } else if (j === 7) {
        imageFormula = `=IMAGE("https://drive.google.com/uc?id="&V${row})`;
      }
      sheet.getRange(row, 23 + j).setValue(imageFormula); // W～AD列：イメージ①～⑧
    }

    // 各シートから必要な情報を取得
    const initWeight = parseFloat(settingsSheet.getRange("C7").getValue()); // 初期重量（未取得時）
    const addWeight = parseFloat(settingsSheet.getRange("C8").getValue()); // 重量加算
    const initStock = parseFloat(settingsSheet.getRange("C9").getValue()); // 初期在庫

    const rateMap = {}; // 各国レート
    rateMap["SG"] = parseFloat(settingsSheet.getRange("F22").getValue()); // SGレート
    rateMap["MY"] = parseFloat(settingsSheet.getRange("F23").getValue()); // MYレート
    rateMap["PH"] = parseFloat(settingsSheet.getRange("F24").getValue()); // PHレート
    rateMap["TH"] = parseFloat(settingsSheet.getRange("F25").getValue()); // THレート
    rateMap["TW"] = parseFloat(settingsSheet.getRange("F26").getValue()); // TWレート
    rateMap["VN"] = parseFloat(settingsSheet.getRange("F27").getValue()); // VNレート
    rateMap["BR"] = parseFloat(settingsSheet.getRange("F28").getValue()); // VNレート

    const productWeight =
      sheet.getRange(row, 5).getValue() === "N/A"
        ? initWeight
        : parseFloat(sheet.getRange(row, 5).getValue()) + addWeight; // 重量 (kg)

    const priceMap = priceCalc(row);
    // プロット前にクリア
    sheet.getRange(row, 31, 1, 34).clearContent();

    // シートにプロット
    // sheet.getRange(row, 31).setValue(`${asin}-${currentDate}`); // AE列：SKU
    sheet.getRange(row, 31).setValue(`${asin}`); // AE列：SKU
    sheet.getRange(row, 32).setValue(priceMap["SG"]);
    sheet.getRange(row, 33).setValue(priceMap["MY"]);
    sheet.getRange(row, 34).setValue(priceMap["PH"]);
    sheet.getRange(row, 35).setValue(priceMap["TH"]);
    sheet.getRange(row, 36).setValue(priceMap["TW"]);
    sheet.getRange(row, 37).setValue(priceMap["VN"]);
    sheet.getRange(row, 38).setValue(priceMap["BR"]);
    sheet
      .getRange(row, 39)
      .setValue(Math.round(priceMap["SG"] * rateMap["SG"]));
    sheet
      .getRange(row, 40)
      .setValue(Math.round(priceMap["MY"] * rateMap["MY"]));
    sheet
      .getRange(row, 41)
      .setValue(Math.round(priceMap["PH"] * rateMap["PH"]));
    sheet
      .getRange(row, 42)
      .setValue(Math.round(priceMap["TH"] * rateMap["TH"]));
    sheet
      .getRange(row, 43)
      .setValue(Math.round(priceMap["TW"] * rateMap["TW"]));
    sheet
      .getRange(row, 44)
      .setValue(Math.round(priceMap["VN"] * rateMap["VN"]));
    sheet.getRange(row, 50).setValue(productWeight);
    sheet
      .getRange(row, 51)
      .setValue(
        sheet.getRange(row, 6).getValue() === "N/A"
          ? settingsSheet.getRange("I7").getValue()
          : sheet.getRange(row, 6).getValue()
      );
    sheet
      .getRange(row, 52)
      .setValue(
        sheet.getRange(row, 7).getValue() === "N/A"
          ? settingsSheet.getRange("I8").getValue()
          : sheet.getRange(row, 7).getValue()
      );
    sheet
      .getRange(row, 53)
      .setValue(
        sheet.getRange(row, 8).getValue() === "N/A"
          ? settingsSheet.getRange("I9").getValue()
          : sheet.getRange(row, 8).getValue()
      );
    sheet.getRange(row, 54).setValue(initStock);

    //////////////////////////////////////////////
    const keepaBrand = productData.brand || "N/A";
    sheet.getRange(row, COL_BRAND_FROM_KEEPA).setValue(keepaBrand); // Keepaから取得したブランドを格納

    let colStart = COL_VARIATION_OUTPUT_START; // CK列 (バリエーションASIN開始列)
    const variationDataList = [];

    while (sheet.getRange(row, colStart + VARIATION_ASIN).getValue()) {
      const variationAsin = sheet
        .getRange(row, colStart + VARIATION_ASIN)
        .getValue();
      if (variationAsin != "") {
        // プロット前にクリア
        sheet
          .getRange(row, colStart + 2, 1, VARIATION_DATAS - 2)
          .clearContent();
        const variationData = fetchProductDataFromKeepa(variationAsin); // バリエーションデータ取得

        if (variationData) {
          writeVariationProductDataToSheet(sheet, row, colStart, variationData);

          // 各国の出品価格を計算
          const priceMap = priceCalcVariation(row, colStart);

          // シートにプロット
          sheet
            .getRange(row, colStart + VARIATION_SKU)
            .setValue(`${variationAsin}`); // AE列：SKU
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG)
            .setValue(priceMap["SG"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 1)
            .setValue(priceMap["MY"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 2)
            .setValue(priceMap["PH"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 3)
            .setValue(priceMap["TH"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 4)
            .setValue(priceMap["TW"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 5)
            .setValue(priceMap["VN"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 6)
            .setValue(priceMap["BR"]);
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN)
            .setValue(Math.round(priceMap["SG"] * rateMap["SG"]));
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN + 1)
            .setValue(Math.round(priceMap["MY"] * rateMap["MY"]));
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN + 2)
            .setValue(Math.round(priceMap["PH"] * rateMap["PH"]));
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN + 3)
            .setValue(Math.round(priceMap["TH"] * rateMap["TH"]));
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN + 4)
            .setValue(Math.round(priceMap["TW"] * rateMap["TW"]));
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN + 5)
            .setValue(Math.round(priceMap["VN"] * rateMap["VN"]));
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG_YEN + 6)
            .setValue(Math.round(priceMap["BR"] * rateMap["BR"]));

          sheet
            .getRange(row, colStart + VARIATION_WEIGHT_CON)
            .setValue(productWeight);
          sheet
            .getRange(row, colStart + VARIATION_LENGTH_CON)
            .setValue(
              sheet.getRange(row, 6).getValue() === "N/A"
                ? settingsSheet.getRange("I7").getValue()
                : sheet.getRange(row, 6).getValue()
            );
          sheet
            .getRange(row, colStart + VARIATION_WIDTH_CON)
            .setValue(
              sheet.getRange(row, 7).getValue() === "N/A"
                ? settingsSheet.getRange("I8").getValue()
                : sheet.getRange(row, 7).getValue()
            );
          sheet
            .getRange(row, colStart + VARIATION_HEIGHT_CON)
            .setValue(
              sheet.getRange(row, 8).getValue() === "N/A"
                ? settingsSheet.getRange("I9").getValue()
                : sheet.getRange(row, 8).getValue()
            );
          sheet
            .getRange(row, colStart + VARIATION_ZAIKO_CON)
            .setValue(initStock);

          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG)
            .setValue(priceMap["SG"]); // SG価格
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 1)
            .setValue(priceMap["MY"]); // MY価格
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 2)
            .setValue(priceMap["PH"]); // PH価格
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 3)
            .setValue(priceMap["TH"]); // TH価格
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 4)
            .setValue(priceMap["TW"]); // TW価格
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 5)
            .setValue(priceMap["VN"]); // VN価格
          sheet
            .getRange(row, colStart + VARIATION_PRICE_SG + 6)
            .setValue(priceMap["BR"]); // VN価格

          // バリエーションデータを配列に追加
          variationDataList.push(variationData);
        }
      }

      colStart += VARIATION_DATAS; // 28列間隔で次のバリエーションASIN列に進む
    }

    // ChatGPTを使用してメイン商品名とメイン商品説明を生成
    const productName = generateMainProductNameWithChatGPT(variationDataList);
    const productDescription =
      generateMainProductDescriptionWithChatGPT(variationDataList);

    // 日本語、英語、簡体語でメイン商品名を出力
    sheet.getRange(row, COL_MAIN_PRODUCT_NAME_JP).setValue(productName.ja);
    // 英語の商品名を255文字以内にする
    const item_name_en_bef = additionalSheet.getRange("D6").getValue();
    const item_name_en_aft = additionalSheet.getRange("D7").getValue();
    var item_name_en = productName.en;
    if (item_name_en_bef) {
      item_name_en = item_name_en_bef + " " + item_name_en;
    }
    if (item_name_en_aft) {
      item_name_en = item_name_en + " " + item_name_en_aft;
    }
    const item_name_en_255 =
      item_name_en.length > 255 ? item_name_en.substring(0, 255) : item_name_en;

    // 英語の商品説明を3000文字以内にする
    const description_en_bef = additionalSheet.getRange("D9").getValue();
    const description_en_aft = additionalSheet.getRange("D10").getValue();
    var description_en = productDescription.en;
    if (description_en_bef) {
      description_en = description_en_bef + "\n" + description_en;
    }
    if (description_en_aft) {
      description_en = description_en + "\n" + description_en_aft;
    }
    const description_en_3000 =
      description_en.length > 3000
        ? description_en.substring(0, 3000)
        : description_en;

    // 中国語の商品名を60文字以内にする
    const item_name_ch_bef = additionalSheet.getRange("E6").getValue();
    const item_name_ch_aft = additionalSheet.getRange("E7").getValue();
    var item_name_ch = productName.zh;
    if (item_name_ch_bef) {
      item_name_ch = item_name_ch_bef + " " + item_name_ch;
    }
    if (item_name_ch_aft) {
      item_name_ch = item_name_ch + " " + item_name_ch_aft;
    }
    const item_name_ch_60 =
      item_name_ch.length > 60 ? item_name_ch.substring(0, 60) : item_name_ch;

    // 中国語の商品説明を3000文字以内にする
    const description_ch_bef = additionalSheet.getRange("E9").getValue();
    const description_ch_aft = additionalSheet.getRange("E10").getValue();
    var description_ch = productDescription.zh;
    if (description_ch_bef) {
      description_ch = description_ch_bef + "\n" + description_ch;
    }
    if (description_ch_aft) {
      description_ch = description_ch + "\n" + description_ch_aft;
    }
    const description_ch_3000 =
      description_ch.length > 3000
        ? description_ch.substring(0, 3000)
        : description_ch;

    sheet.getRange(row, COL_MAIN_PRODUCT_NAME_EN).setValue(item_name_en_255);
    sheet.getRange(row, COL_MAIN_PRODUCT_NAME_ZH).setValue(item_name_ch_60);

    // 日本語、英語、簡体語で商品説明を出力
    sheet
      .getRange(row, COL_MAIN_DESCRIPTION_JP)
      .setValue(productDescription.ja);
    sheet.getRange(row, COL_MAIN_DESCRIPTION_EN).setValue(description_en_3000);
    sheet.getRange(row, COL_MAIN_DESCRIPTION_ZH).setValue(description_ch_3000);

    sheet
      .getRange(row, COL_MAIN_PRODUCT_NAME_EN_CON)
      .setValue(item_name_en_255);
    sheet.getRange(row, COL_MAIN_PRODUCT_NAME_ZH_CON).setValue(item_name_ch_60);

    // 日本語、英語、簡体語で商品説明を出力
    sheet
      .getRange(row, COL_MAIN_DESCRIPTION_EN_CON)
      .setValue(description_en_3000);
    sheet
      .getRange(row, COL_MAIN_DESCRIPTION_ZH_CON)
      .setValue(description_ch_3000);
  });
}

function processMakeDataStep1to2_Step2ChatGPT() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODUCTS);
  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("基本設定");
  const additionalSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("追加設定");
  const lastRow = sheet.getLastRow();
  const checkRange = sheet.getRange(2, COL_CHECKBOX, lastRow - 1); // チェックボックス列の範囲
  const asinRange = sheet.getRange(2, COL_ASIN, lastRow - 1); // メインASIN列の範囲
  const checkValues = checkRange.getValues(); // チェックボックス状態
  const asinValues = asinRange.getValues(); // ASIN

  // 基本設定シートからAPIキーとフォルダIDを取得
  const apiKey = settingsSheet.getRange("C4").getValue();
  // const apiKey = PropertiesService.getScriptProperties().getProperty('KEEPA_API_KEY');
  const folderId = settingsSheet.getRange("C5").getValue();

  const rowsToProcess = [];

  // チェックされた行をリストに追加
  for (let i = 0; i < checkValues.length; i++) {
    const isChecked = checkValues[i][0]; // チェックボックス
    const asin = asinValues[i][0]; // ASIN

    if (isChecked && asin) {
      rowsToProcess.push(i + 2); // 行番号を保存 (スプレッドシート上では2行目から)
    }
  }

  // 処理対象が無ければ終了
  if (rowsToProcess.length === 0) {
    Logger.log("処理対象の行がありません。");
    return;
  }

  // 各行に対して1回だけ実行する処理
  rowsToProcess.forEach(function (row) {
    if (!sheet || !apiKey || !folderId) {
      Logger.log("出品一覧シート、APIキー、またはフォルダIDが見つかりません。");
      return;
    }

    const asin = sheet.getRange(row, COL_ASIN).getValue(); // メインASIN

    if (!asin) {
      Logger.log(`行${row}にASINがありません。`);
      return;
    }

    let colStart = COL_VARIATION_OUTPUT_START; // CK列 (バリエーションASIN開始列)
    const variationDataList = [];

    while (sheet.getRange(row, colStart + VARIATION_ASIN).getValue()) {
      const variationAsin = sheet
        .getRange(row, colStart + VARIATION_ASIN)
        .getValue();
      if (variationAsin != "") {
        // // プロット前にクリア
        const variationData = fetchProductDataFromKeepa(variationAsin); // バリエーションデータ取得

        if (variationData) {
          // バリエーションデータを配列に追加
          variationDataList.push(variationData);
        }
      }

      colStart += VARIATION_DATAS; // 28列間隔で次のバリエーションASIN列に進む
    }

    // ChatGPTでバリエーション1名、バリエーション2名を生成
    const variations = generateVariationsWithChatGPT(variationDataList);

    saveVariationsAndAssignmentsToSheet(row, variations, sheet);
    processMakeDataStep1();
    Logger.log(`行${row}のバリエーション名を生成しました。`);
  });
}
