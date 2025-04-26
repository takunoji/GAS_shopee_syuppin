// バックエンド側のmainFunction.gs
// チェックされた行の処理
function processCheckedShops(selectedShopsIndices, rowsToProcess) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const shopsData = JSON.parse(
    scriptProperties.getProperty("SHOPS_DATA") || "[]"
  );

  // 処理結果の保存用
  const processedRows = [];

  // 各行の処理
  for (const row of rowsToProcess) {
    try {
      const rowNum = row.rowNum;
      const asin = row.asin;
      const rowData = row.data;

      // 結果データを初期化
      const resultData = {};

      // 現在の日付
      const currentDate = Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyyy/MM/dd"
      );
      resultData[3] = currentDate; // C列: データ収集日

      // KeepaからのAmazon商品情報取得
      const keepaData = processKeepaProduct(asin);

      // 商品データをresultDataに追加
      Object.assign(resultData, keepaData);

      // Shopee出品データの生成
      const listingData = processShopeeListingData(keepaData, asin);
      Object.assign(resultData, listingData);

      // 画像のアップロード
      const imageIds = uploadImagesToShopee(rowData.imageIds);
      for (let i = 0; i < imageIds.length; i++) {
        resultData[56 + i] = imageIds[i]; // 画像ID列
      }

      // 各ショップへの出品処理
      for (const index of selectedShopsIndices) {
        const shop = shopsData[index];
        const shopResultData = processAddShopeeItem(
          asin,
          listingData,
          imageIds,
          shop.shop_id,
          shop.access_token,
          shop.region,
          shop.shop_name
        );

        // ショップ固有のデータをresultDataに追加
        Object.assign(resultData, shopResultData);
      }

      // チェックボックスをオフに設定
      resultData[1] = false;

      // 処理日時を設定
      resultData[62] = currentDate;

      // 処理結果を追加
      processedRows.push({
        rowNum: rowNum,
        data: resultData,
      });
    } catch (error) {
      console.error(`行 ${row.rowNum} の処理中にエラー: ${error.message}`);
    }
  }

  return { processedRows: processedRows };
}

// Keepa API呼び出し
function processKeepaProduct(asin) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty("KEEPA_API_KEY");
  const folderId = scriptProperties.getProperty("FOLDER_ID");

  if (!apiKey || !folderId) {
    throw new Error("Keepa APIキーまたはフォルダIDが設定されていません。");
  }

  let productData = postKeepaProducts(apiKey, asin);
  if (!productData) {
    throw new Error("Keepa APIトークンが使い切られました。");
  }

  // 文字列からオブジェクトに変換
  productData = JSON.parse(productData);

  // 画像のダウンロードとGoogle Driveへの保存
  const imageIds = [];
  if (productData.image) {
    const images = productData.image.split(",");
    for (let j = 0; j < images.length && j < 8; j++) {
      const fileId = downloadImageAndSave(images[j], folderId);
      imageIds.push(fileId || "N/A");
    }
  }

  // 結果データの生成
  const resultData = {
    4: productData.price || "N/A", // 新品最安値
    5: productData.weight ? productData.weight / 1000 : "N/A", // 重量 (kg)
    6: productData.length ? productData.length / 10 : "N/A", // 長さ (cm)
    7: productData.width ? productData.width / 10 : "N/A", // 幅 (cm)
    8: productData.height ? productData.height / 10 : "N/A", // 高さ (cm)
    9: productData.title || "N/A", // 商品名
    10: productData.features
      ? productData.features.join("\n").replace(/\n+/g, "\n")
      : "" +
        "\n" +
        (productData.description
          ? productData.description.replace(/\n+/g, "\n")
          : ""), // 商品説明
    65: productData.brand || "N/A", // ブランド名
  };

  // 画像IDを追加
  for (let j = 0; j < 8; j++) {
    resultData[15 + j] = imageIds[j] || "N/A"; // 画像ID列
  }

  return resultData;
}

// 出品データの作成
function processShopeeListingData(keepaData, asin) {
  // ここでは、mainFunction.gsの既存のprocessShopeeListingRow関数の内容を
  // スプレッドシート直接操作ではなくデータを返すように変更

  const currentDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd"
  );

  // スクリプトプロパティから設定を取得
  const scriptProperties = PropertiesService.getScriptProperties();
  const initWeight = parseFloat(
    scriptProperties.getProperty("INIT_WEIGHT") || "0.5"
  );
  const addWeight = parseFloat(
    scriptProperties.getProperty("ADD_WEIGHT") || "0.1"
  );
  const initStock = parseFloat(
    scriptProperties.getProperty("INIT_STOCK") || "10"
  );

  // 価格計算
  const priceMap = calculatePrices(keepaData, asin);

  // 結果データ
  const resultData = {
    31: `${asin}-${currentDate}`, // SKU
    32: priceMap.SG,
    33: priceMap.MY,
    34: priceMap.PH,
    35: priceMap.TH,
    36: priceMap.TW,
    37: priceMap.VN,
    38: priceMap.BR,
    39: Math.round(priceMap.SG * scriptProperties.getProperty("RATE_SG")),
    40: Math.round(priceMap.MY * scriptProperties.getProperty("RATE_MY")),
    41: Math.round(priceMap.PH * scriptProperties.getProperty("RATE_PH")),
    42: Math.round(priceMap.TH * scriptProperties.getProperty("RATE_TH")),
    43: Math.round(priceMap.TW * scriptProperties.getProperty("RATE_TW")),
    44: Math.round(priceMap.VN * scriptProperties.getProperty("RATE_VN")),
    45: Math.round(priceMap.BR * scriptProperties.getProperty("RATE_BR")),
    // 重量と寸法
    49: parseFloat(keepaData[5]) + addWeight, // 重量 (kg)
    50: keepaData[6] === "N/A" ? initWeight : parseFloat(keepaData[6]), // 長さ (cm)
    51: keepaData[7] === "N/A" ? initWeight : parseFloat(keepaData[7]), // 幅 (cm)
    52: keepaData[8] === "N/A" ? initWeight : parseFloat(keepaData[8]), // 高さ (cm)
    // 在庫数
    54: initStock,
  };

  return resultData;
}

// 価格計算
function calculatePrices(keepaData, asin) {
  // priceCalcと同様の計算ロジックを実装
  // スプレッドシート参照の代わりにパラメータから計算

  const scriptProperties = PropertiesService.getScriptProperties();
  const initWeight = parseFloat(
    scriptProperties.getProperty("INIT_WEIGHT") || "0.5"
  );
  const addWeight = parseFloat(
    scriptProperties.getProperty("ADD_WEIGHT") || "0.1"
  );
  const profitRatio = parseFloat(
    scriptProperties.getProperty("PROFIT_RATIO") || "0.2"
  );

  // レート情報
  const rateMap = {
    SG: parseFloat(scriptProperties.getProperty("RATE_SG") || "84"),
    MY: parseFloat(scriptProperties.getProperty("RATE_MY") || "26"),
    PH: parseFloat(scriptProperties.getProperty("RATE_PH") || "2.2"),
    TH: parseFloat(scriptProperties.getProperty("RATE_TH") || "3.2"),
    TW: parseFloat(scriptProperties.getProperty("RATE_TW") || "4"),
    VN: parseFloat(scriptProperties.getProperty("RATE_VN") || "0.0049"),
    BR: parseFloat(scriptProperties.getProperty("RATE_BR") || "21"),
  };

  // 実際の価格計算ロジック（簡易版）
  const amazonPrice = parseFloat(keepaData[4]) || 0;
  const weight =
    keepaData[5] === "N/A" ? initWeight : parseFloat(keepaData[5]) + addWeight;

  // 価格計算（実際はさらに複雑な計算を行うべき）
  const basePrice = amazonPrice * 1.1 + weight * 1000; // 10%マージン + 重量加算

  return {
    SG: Math.round((basePrice / rateMap.SG) * 100) / 100,
    MY: Math.round((basePrice / rateMap.MY) * 100) / 100,
    PH: Math.round(basePrice / rateMap.PH),
    TH: Math.round(basePrice / rateMap.TH),
    TW: Math.round((basePrice / rateMap.TW) * 100) / 100,
    VN: Math.round(basePrice / rateMap.VN),
    BR: Math.round((basePrice / rateMap.BR) * 100) / 100,
  };
}

// 処理行なしで済む関数
function processMakeData(rowsToProcess) {
  // 各行の処理結果を保存
  const processedRows = [];

  for (const row of rowsToProcess) {
    try {
      const asin = row.asin;

      // Keepaデータ取得
      const keepaData = processKeepaProduct(asin);

      // 出品データ生成
      const listingData = processShopeeListingData(keepaData, asin);

      // 結果をマージ
      const resultData = { ...keepaData, ...listingData };

      processedRows.push({
        rowNum: row.rowNum,
        data: resultData,
      });
    } catch (error) {
      console.error(`行 ${row.rowNum} の処理中にエラー: ${error.message}`);
    }
  }

  return { processedRows: processedRows };
}

// 出品データのみ生成
function processCalcData(rowsToProcess) {
  // 各行の処理結果を保存
  const processedRows = [];

  for (const row of rowsToProcess) {
    try {
      const asin = row.asin;
      const data = row.data;

      // 既存データを使用して出品データのみ生成
      const keepaData = {
        4: data.price,
        5: data.weight,
        6: data.length,
        7: data.width,
        8: data.height,
      };

      // 出品データ生成
      const listingData = processShopeeListingData(keepaData, asin);

      processedRows.push({
        rowNum: row.rowNum,
        data: listingData,
      });
    } catch (error) {
      console.error(`行 ${row.rowNum} の処理中にエラー: ${error.message}`);
    }
  }

  return { processedRows: processedRows };
}

// バリエーション生成ステップ1
function processMakeDataStep1(rowsToProcess) {
  // 各行の処理結果を保存
  const processedRows = [];

  for (const row of rowsToProcess) {
    try {
      const variationType = row.variationType;

      if (variationType === "1つ" || variationType === "2つ") {
        // バリエーション組み合わせ生成ロジック
        const combinations = generateVariationCombinations(row);

        // 結果データ
        const resultData = {};

        // バリエーション組み合わせをCK列以降に設定
        combinations.forEach((combination, index) => {
          resultData[134 + index * 42] = combination;
        });

        processedRows.push({
          rowNum: row.rowNum,
          data: resultData,
        });
      }
    } catch (error) {
      console.error(`行 ${row.rowNum} の処理中にエラー: ${error.message}`);
    }
  }

  return { processedRows: processedRows };
}

// バリエーション組み合わせを生成
function generateVariationCombinations(rowData) {
  const combinations = [];

  if (rowData.variationType === "1つ") {
    // バリエーション1の値を直接使用
    return rowData.variation1Values;
  } else if (rowData.variationType === "2つ") {
    // バリエーション1と2の組み合わせを生成
    rowData.variation1Values.forEach((value1) => {
      rowData.variation2Values.forEach((value2) => {
        combinations.push(`${value1}, ${value2}`);
      });
    });
    return combinations;
  }

  return combinations;
}
