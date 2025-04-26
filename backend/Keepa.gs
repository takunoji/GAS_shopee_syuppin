/**
 * Keepa APIを使ってAmazonデータを収集し、商品画像をGoogle Driveに保存。

 */
function testKeepa() {
  const key =
    "fkfe1gchqar95amvecs7jdsg18tr3rd66rqh57n2tjtlgtqjo3a0vr6lq2j5uakc";
  const asin = "B0CGR1DS3M";
  const response = postKeepaProducts(key, asin);
  Logger.log(response);
}

function testDownload() {
  const fileName = "71YudOwUvbL.jpg";
  const folderId = "1L0pTXdCssPU6Bva7e1lxiu6YHzrRfHWS";
  const response = downloadImageAndSave(fileName, folderId);
  Logger.log(response);
}

/**
 * Keepa APIを呼び出してAmazonの商品情報（価格、サイズ、重量、説明文、画像URL）を取得。
 */
function postKeepaProducts(key, asin) {
  const KeepaProductURL = "https://api.keepa.com/product";
  const payload = {
    key,
    asin,
    domain: "5",
  };

  const tokensLeft = getKeepaTokensLeft(key);

  if (tokensLeft) {
    let options = {
      method: "post",
      muteHttpExceptions: true,
      payload: payload,
    };

    const response = UrlFetchApp.fetch(KeepaProductURL, options);
    const keepaProducts = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 200) {
      throw new Error(response.getContentText());
    } else {
      const price = keepaProducts.products[0].csv[1].slice(-1)[0];
      const height = keepaProducts.products[0].packageHeight;
      const length = keepaProducts.products[0].packageLength;
      const width = keepaProducts.products[0].packageWidth;
      const weight = keepaProducts.products[0].packageWeight;
      const title = keepaProducts.products[0].title;
      const features = keepaProducts.products[0].features;
      const description = keepaProducts.products[0].description;
      const image = keepaProducts.products[0].imagesCSV;
      const brand = keepaProducts.products[0].brand; // ブランド情報を追加

      const keepData = {
        price,
        height,
        length,
        width,
        weight,
        title,
        features,
        description,
        image,
        brand, // ブランド情報を含む
      };

      return JSON.stringify(keepData, null, "\t");
    }
  }

  return null;
}

/**
 * Keepa APIの残りトークン数を取得
 * @param {string} key - Keepa APIキー
 * @return {number} - 残りトークン数
 */
function getKeepaTokensLeft(key) {
  const response = checkKeepaApiKey(key);
  const tokensLeft = JSON.parse(response.getContentText()).tokensLeft;
  //Logger.log("getKeepaTokensLeft response: " + response.getContentText());
  return tokensLeft;
}

/**
 * Keepa APIキーをチェック
 * @param {string} key - Keepa APIキー
 * @return {HTTPResponse} - Keepa APIのレスポンス
 */
function checkKeepaApiKey(key) {
  const url = "https://api.keepa.com/token?key=" + key; // パラメータをURLに含める
  const options = {
    method: "get",
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    Logger.log("checkKeepaApiKey response: " + response.getContentText());
    throw new Error(response.getContentText());
  }
  //Logger.log("checkKeepaApiKey response: " + response.getContentText());
  return response;
}

/**
 * 商品画像をAmazonからダウンロードし、Google Driveに保存。
 */
function downloadImageAndSave(fileName, folderId) {
  var response = UrlFetchApp.fetch(
    `https://m.media-amazon.com/images/I/${fileName}`
  );
  var fileBlob = response.getBlob().setName(fileName);
  var folder = DriveApp.getFolderById(folderId);
  var file = folder.createFile(fileBlob);

  return file.getId(); // ファイルID 1DJ1umRH9T8xAcAYCT7vyCtCKq1YIUW2s
}
