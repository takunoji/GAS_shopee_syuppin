// バックエンド側のapi.gs
function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  // CORS ヘッダーを設定
  output.setHeader("Access-Control-Allow-Origin", "*");
  output.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  output.setHeader("Access-Control-Allow-Headers", "Content-Type");
  try {
    // リクエストデータのパース
    const requestData = JSON.parse(e.postData.contents);
    const functionName = requestData.function;
    const params = requestData.parameters;
    const apiKey = requestData.apiKey;

    // // APIキー検証
    // if (!validateApiKey(apiKey)) {
    //   return createJsonResponse({
    //     status: "error",
    //     error: "不正なAPIキー"
    //   });
    // }

    let result;
    // 関数呼び出し
    switch (functionName) {
      case "getAuthorizationUrl":
        result = getAuthorizationUrl();
        break;
      case "getShopsData":
        result = getShopsData();
        break;
      case "saveApiSettings":
        result = saveApiSettings(
          params.partnerId,
          params.partnerKey,
          params.redirectUrl
        );
        break;
      case "processKeepaProduct":
        // Keepa APIを呼び出して商品データを取得するだけ
        result = processKeepaProduct(params.asin);
        break;
      case "uploadImagesToShopee":
        // 画像をShopeeにアップロードするだけ
        result = uploadImagesToShopee(params.fileIds);
        break;
      case "getShopeeCategories":
        // Shopeeカテゴリを取得するだけ
        result = getCategoryRecommend(
          params.shop_id,
          params.access_token,
          params.item_name,
          params.cover_image
        );
        break;
      case "addShopeeItem":
        // 商品を登録するだけ
        result = processAddShopeeItem(
          params.asin,
          params.listingData,
          params.imageIds,
          params.shop_id,
          params.access_token,
          params.region,
          params.shop_name
        );
        break;
      case "processCheckedShops":
        result = processCheckedShops(
          params.selectedShopsIndices,
          params.rowsToProcess
        );
        break;
      case "processMakeData":
        result = processMakeData(params.rowsToProcess);
        break;
      case "processCalcData":
        result = processCalcData(params.rowsToProcess);
        break;
      case "processMakeDataStep1":
        result = processMakeDataStep1(params.rowsToProcess);
        break;
      // 他の関数も同様に追加
      case "processKeepaProducts":
        result = processKeepaProducts(params);
        break;
      default:
        throw new Error(`未知の関数: ${functionName}`);
    }

    return createJsonResponse({
      status: "success",
      result: result,
    });
  } catch (error) {
    console.error("バックエンドエラー:", error);
    return createJsonResponse({
      status: "error",
      error: error.message,
    });
  }
}

// CORS対応のためのヘッダー設定
function doOptions(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  output.setContent(JSON.stringify({ status: "success" }));

  output.setHeader("Access-Control-Allow-Origin", "*");
  output.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  output.setHeader("Access-Control-Allow-Headers", "Content-Type");

  lock.releaseLock();

  return output;
}

// APIキー検証
function validateApiKey(apiKey) {
  // 実際の実装では適切な検証を行う
  // 例: const validApiKey = PropertiesService.getScriptProperties().getProperty("API_KEY");
  // return apiKey === validApiKey;
  return true; // 開発中は常にtrueを返す
}
/**
 * ウェブアプリケーションのトップページ
 */
function doGet() {
  return HtmlService.createHtmlOutput(`
    <h1>Shopee Integration Backend</h1>
    <p>このサービスはフロントエンドアプリケーションからのAPI呼び出しを受け付けます。</p>
    <p>現在の時刻: ${new Date().toLocaleString()}</p>
  `);
}
// JSON応答の作成
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// デプロイURLを取得するユーティリティ関数
function getDeployedUrl() {
  const url = ScriptApp.getService().getUrl();
  return url;
}

// Keepa APIを使ってAmazonデータを収集
function processKeepaProducts(params) {
  try {
    const asinList = params.asinList;
    if (!asinList || !Array.isArray(asinList) || asinList.length === 0) {
      return { error: "ASINリストが空か無効です" };
    }

    // APIキーの取得
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiKey = scriptProperties.getProperty("KEEPA_API_KEY");
    const folderId = scriptProperties.getProperty("FOLDER_ID");

    if (!apiKey || !folderId) {
      return { error: "Keepa APIキーまたはフォルダIDが設定されていません" };
    }

    // 各ASINのデータを収集
    const asinData = {};

    for (const asin of asinList) {
      try {
        // Keepa APIを呼び出し
        const productData = postKeepaProducts(apiKey, asin);
        if (!productData) {
          asinData[asin] = { error: "APIトークンが不足しています" };
          continue;
        }

        // JSON文字列をオブジェクトに変換
        const parsedData = JSON.parse(productData);

        // 結果データを生成
        const currentDate = Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          "yyyy/MM/dd"
        );

        const resultData = {
          3: currentDate, // C列：データ取得日
          4: parsedData.price || "N/A", // D列：新品最安値
          5: parsedData.weight ? parsedData.weight / 1000 : "N/A", // E列：重量 (kg)
          6: parsedData.length ? parsedData.length / 10 : "N/A", // F列：長さ (cm)
          7: parsedData.width ? parsedData.width / 10 : "N/A", // G列：幅 (cm)
          8: parsedData.height ? parsedData.height / 10 : "N/A", // H列：高さ (cm)
          9: parsedData.title || "N/A", // I列：商品名
          10: parsedData.features
            ? parsedData.features.join("\n").replace(/\n+/g, "\n") +
              "\n" +
              (parsedData.description
                ? parsedData.description.replace(/\n+/g, "\n")
                : "")
            : parsedData.description
            ? parsedData.description.replace(/\n+/g, "\n")
            : "N/A", // J列：商品説明
          65: parsedData.brand || "N/A", // ブランド名
        };

        // 画像をダウンロードしてGoogle Driveに保存
        if (parsedData.image) {
          const images = parsedData.image.split(",");

          for (let j = 0; j < images.length && j < 8; j++) {
            try {
              const fileId = downloadImageAndSave(images[j], folderId);
              resultData[15 + j] = fileId || "N/A"; // O～V列：画像①～⑧
            } catch (imgError) {
              console.error(
                `画像ダウンロードエラー (${asin}, ${j}): ${imgError.message}`
              );
              resultData[15 + j] = "N/A";
            }
          }
        }

        // ASINごとのデータを保存
        asinData[asin] = resultData;
      } catch (asinError) {
        console.error(`ASIN処理エラー (${asin}): ${asinError.message}`);
        asinData[asin] = { error: asinError.message };
      }
    }

    return { asinData: asinData };
  } catch (error) {
    console.error("Keepaデータ収集エラー:", error);
    return { error: error.message };
  }
}

function testEcho(params) {
  return { message: "テスト成功", params: params };
}
