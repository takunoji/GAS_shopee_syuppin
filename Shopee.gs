/**
 * Shopee APIを操作する関数群（出品処理、画像アップロード、ブランド取得など）。
 */

/**
 * Google Driveから画像を取得し、Shopee APIでアップロードする。
 */
function postUploadImage(fileId) {
  const timestamp = Math.floor(Date.now() / 1000);
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();

  // 正しい MIME タイプを使用する (image/jpeg や image/png)
  const mimeType = blob.getContentType(); // MIME タイプを取得
  if (mimeType !== "image/jpeg" && mimeType !== "image/png") {
    Logger.log("Unsupported image type: " + mimeType);
    return;
  }

  const boundary = "----WebKitFormBoundary" + Utilities.getUuid(); // ランダムな boundary を生成

  const baseString = `${PARTNER_ID}${UPLOAD_IMAGE_API_PATH}${timestamp}`;
  const signature = Utilities.computeHmacSha256Signature(
    baseString,
    PARTNER_KEY
  );
  const encodedSignature = signature
    .map((e) => (e < 0 ? e + 256 : e).toString(16).padStart(2, "0"))
    .join("");

  const url =
    `${API_URL}${UPLOAD_IMAGE_API_PATH}` +
    `?partner_id=${PARTNER_ID}` +
    `&timestamp=${timestamp}` +
    `&sign=${encodedSignature}`;

  // Multipart/form-data のリクエストボディを厳密に組み立てる
  const formData = Utilities.newBlob("")
    .setDataFromString(
      `--${boundary}\r\n` +
        `Content-Disposition: form-data; name="image"; filename="${file.getName()}"\r\n` +
        `Content-Type: ${mimeType}\r\n\r\n`
    )
    .getBytes()
    .concat(blob.getBytes()) // ファイルのバイナリデータを追加
    .concat(Utilities.newBlob(`\r\n--${boundary}--\r\n`).getBytes()); // リクエストを終了

  const options = {
    method: "post",
    contentType: `multipart/form-data; boundary=${boundary}`,
    payload: formData,
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());

    if (responseData.error) {
      Logger.log("Error uploading image: " + responseData.message);
      return null;
    }

    const imageId = responseData.response.image_info.image_id;
    return imageId; // sg-11134201-7rdxy-m01fhjzynjgp83
  } catch (e) {
    Logger.log("Error occurred: " + e.message);
    return null;
  }
}

function testCatRec() {
  const shop_id = "1184051917";
  const access_token = "4b5868514553446667494e427455786e";
  const item_name = `Weiss Schwarz Booster Pack Hololive Production Vol.2 BOX`;
  const cover_image = `sg-11134201-7rdvy-m0hzrjpezgn8c7`;

  const response = getCategoryRecommend(
    shop_id,
    access_token,
    item_name,
    cover_image
  );

  Logger.log(parseInt(response[0]));
}

/**
 * Shopeeに商品名を送信し、カテゴリIDの候補を取得する。
 */
function getCategoryRecommend(shop_id, access_token, item_name, cover_image) {
  try {
    const timestamp = Math.floor(Date.now() / 1000);

    const baseString = `${PARTNER_ID}${CATEGORY_RECOMMEND_API_PATH}${timestamp}${access_token}${shop_id}`;
    const signature = Utilities.computeHmacSha256Signature(
      baseString,
      PARTNER_KEY
    );
    const encodedSignature = signature
      .map((e) => (e < 0 ? e + 256 : e).toString(16).padStart(2, "0"))
      .join("");

    const options = {
      method: "get",
      contentType: "application/json",
      muteHttpExceptions: true,
    };

    const url =
      `${API_URL}${CATEGORY_RECOMMEND_API_PATH}` +
      `?partner_id=${PARTNER_ID}` +
      `&shop_id=${shop_id}` +
      `&timestamp=${timestamp}` +
      `&access_token=${access_token}` +
      `&sign=${encodedSignature}` +
      `&item_name=${item_name}` +
      `&product_cover_image=${cover_image}`;
    Logger.log(url);

    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());

    return responseData.response.category_id; // [100000,100001,100002]
  } catch (e) {
    Logger.log("An error occurred: " + e.message);
    return null;
  }
}

function testAtt() {
  const shop_id = "1184058570";
  const access_token = "517a524f624368794b4754584142544d";
  const category_id = "101642";

  const response = getMandatoryAttributes(shop_id, access_token, category_id);

  Logger.log(response);
}

/**
 * 指定したカテゴリIDに対する必須属性リストを取得する。
 */
function getMandatoryAttributes(shop_id, access_token, category_id) {
  try {
    const timestamp = Math.floor(Date.now() / 1000);

    const baseString = `${PARTNER_ID}${ATTRIBUTES_API_PATH}${timestamp}${access_token}${shop_id}`;
    const signature = Utilities.computeHmacSha256Signature(
      baseString,
      PARTNER_KEY
    );
    const encodedSignature = signature
      .map((e) => (e < 0 ? e + 256 : e).toString(16).padStart(2, "0"))
      .join("");

    const options = {
      method: "get",
      contentType: "application/json",
      muteHttpExceptions: true,
    };

    const url =
      `${API_URL}${ATTRIBUTES_API_PATH}` +
      `?partner_id=${PARTNER_ID}` +
      `&shop_id=${shop_id}` +
      `&timestamp=${timestamp}` +
      `&access_token=${access_token}` +
      `&sign=${encodedSignature}` +
      `&language=en` +
      `&category_id=${category_id}`;
    Logger.log(url);

    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());

    // レスポンスにattribute_listが存在しない場合は終了
    if (!responseData.response || !responseData.response.attribute_list) {
      Logger.log("Response text is empty or null.");
      return null;
    }

    // 必須のAttributeのみ取得する
    const attributeList = responseData.response.attribute_list
      .filter((attribute) => attribute.is_mandatory) // is_mandatoryがtrueのものだけをフィルター
      .map((attribute) => {
        const firstValue =
          attribute.attribute_value_list &&
          attribute.attribute_value_list.length > 0
            ? {
                value_id: attribute.attribute_value_list[0].value_id,
                value_name:
                  attribute.attribute_value_list[0].original_value_name,
              }
            : {
                value_id: 0,
                value_name: "-",
              };

        return {
          attribute_id: attribute.attribute_id,
          attribute_name: attribute.original_attribute_name,
          first_value: firstValue,
        };
      });

    return attributeList;
  } catch (e) {
    Logger.log("An error occurred: " + e.message);
    return null;
  }
}

/**
 * 商品情報をShopee APIに送信し、商品を出品する。
 */
function postAddItem(shop_id, access_token, payload) {
  const timestamp = Math.floor(Date.now() / 1000);

  const baseString = `${PARTNER_ID}${ADD_ITEM_API_PATH}${timestamp}${access_token}${shop_id}`;
  const signature = Utilities.computeHmacSha256Signature(
    baseString,
    PARTNER_KEY
  );
  const encodedSignature = signature
    .map((e) => (e < 0 ? e + 256 : e).toString(16).padStart(2, "0"))
    .join("");

  const url =
    `${API_URL}${ADD_ITEM_API_PATH}` +
    `?partner_id=${PARTNER_ID}` +
    `&shop_id=${shop_id}` +
    `&timestamp=${timestamp}` +
    `&access_token=${access_token}` +
    `&sign=${encodedSignature}`;

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());

    if (responseData.error) {
      Logger.log("Error adding item: " + responseData.message);
      return response;
    } else {
      return response;
    }
  } catch (e) {
    Logger.log("Error occurred: " + e.message);
    return e;
  }
}

/**
 * Shopee APIからブランド名のリストを取得する（カテゴリIDが必要）。
 */
function getBrandList(shopId, accessToken, categoryId) {
  const timestamp = Math.floor(Date.now() / 1000);
  const baseString = `${PARTNER_ID}${BRAND_LIST_API_PATH}${timestamp}${accessToken}${shopId}`;
  const signature = Utilities.computeHmacSha256Signature(
    baseString,
    PARTNER_KEY
  );
  const encodedSignature = signature
    .map((e) => (e < 0 ? e + 256 : e).toString(16).padStart(2, "0"))
    .join("");

  const url =
    `${API_URL}${BRAND_LIST_API_PATH}` +
    `?partner_id=${PARTNER_ID}` +
    `&shop_id=${shopId}` +
    `&timestamp=${timestamp}` +
    `&access_token=${accessToken}` +
    `&category_id=${categoryId}` +
    `&sign=${encodedSignature}`;

  const options = {
    method: "get",
    contentType: "application/json",
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());

    // ブランドリストを取得
    if (
      responseData &&
      responseData.response &&
      responseData.response.brand_list
    ) {
      const brandList = responseData.response.brand_list;
      return brandList.map((brand) => ({
        brand_id: brand.brand_id,
        brand_name: brand.original_brand_name,
      }));
    }
  } catch (error) {
    Logger.log("ブランドリスト取得エラー: " + error.message);
  }

  return [];
}
////////////////////////////
/**
 * Shopee APIからブランド名のリストを取得する（カテゴリIDが必要）。
 */
/**
 * Shopee APIからブランド名のリストを取得する（カテゴリIDが必要）。
 */
function fetchBrandListFromShopee(shopId, accessToken, categoryId) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ブランドリスト");
  const offsetSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("オフセット管理") ||
    SpreadsheetApp.getActiveSpreadsheet().insertSheet("オフセット管理");

  // オフセット管理シートのヘッダー確認または追加
  if (offsetSheet.getLastRow() === 0) {
    offsetSheet.appendRow(["カテゴリID", "オフセット", "完了ステータス"]); // 3列目に完了ステータス
  }

  // 現在のオフセットと完了ステータスを取得または初期化
  const offsetData = offsetSheet.getDataRange().getValues();
  let currentOffset = 0;
  let offsetRowIndex = offsetData.findIndex(
    (row) => String(row[0]) === String(categoryId)
  );
  let isCompleted = false;

  if (offsetRowIndex !== -1) {
    currentOffset = parseInt(offsetData[offsetRowIndex][1] || "0", 10);
    isCompleted = offsetData[offsetRowIndex][2] === "完了";
  } else {
    offsetRowIndex = offsetSheet.getLastRow() + 1;
    offsetSheet.appendRow([categoryId, currentOffset, "未完了"]);
  }

  // 完了しているカテゴリはスキップ
  if (isCompleted) {
    Logger.log(
      `カテゴリID: ${categoryId} は既に完了しています。スキップします。`
    );
    return;
  }

  const pageSize = 100; // 一度に取得するブランド数
  let hasMore = true;

  while (hasMore) {
    const timestamp = Math.floor(Date.now() / 1000);
    const path = "/api/v2/product/get_brand_list";
    const sign = generateSign(
      path,
      PARTNER_ID,
      shopId,
      accessToken,
      timestamp,
      PARTNER_KEY
    );

    const url =
      `${API_URL}${path}` +
      `?partner_id=${PARTNER_ID}&shop_id=${shopId}&timestamp=${timestamp}` +
      `&access_token=${accessToken}&category_id=${categoryId}` +
      `&offset=${currentOffset}&page_size=${pageSize}&status=1&sign=${sign}`;

    const options = {
      method: "get",
      contentType: "application/json",
      muteHttpExceptions: true,
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseData = JSON.parse(response.getContentText());

      if (responseData.error) {
        throw new Error(`ブランドリスト取得エラー: ${responseData.message}`);
      }

      const brandList = responseData.response.brand_list || [];
      if (brandList.length > 0) {
        saveBrandListToSheet(sheet, categoryId, brandList); // ブランドリストを保存
      }

      hasMore = responseData.response.has_next_page;
      currentOffset =
        responseData.response.next_offset || currentOffset + pageSize;

      // オフセットをシートに保存 (既存の行を更新)
      offsetSheet.getRange(offsetRowIndex + 1, 2).setValue(currentOffset);

      // 最終ページの場合は完了ステータスを設定
      if (!hasMore) {
        offsetSheet.getRange(offsetRowIndex + 1, 3).setValue("完了");
      }
    } catch (error) {
      Logger.log(`ブランドリスト取得エラー: ${error.message}`);
      break;
    }
  }

  Logger.log(`カテゴリID: ${categoryId} のブランドリスト取得完了`);
}

/**
 * Shopee 全カテゴリを取得
 */
function fetchAllCategories(shopId, accessToken) {
  const timestamp = Math.floor(Date.now() / 1000);
  const path = "/api/v2/product/get_category";
  const sign = generateSign(
    path,
    PARTNER_ID,
    shopId,
    accessToken,
    timestamp,
    PARTNER_KEY
  );

  const url =
    `${API_URL}${path}` +
    `?shop_id=${shopId}&access_token=${accessToken}&timestamp=${timestamp}` +
    `&partner_id=${PARTNER_ID}&sign=${sign}`;

  const options = {
    method: "get",
    contentType: "application/json",
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());

    if (responseData.error) {
      throw new Error(`カテゴリ取得エラー: ${responseData.message}`);
    }

    const categories = responseData.response.category_list || [];
    const leafCategories = categories.filter(
      (category) => category.has_children === false
    );
    Logger.log(`葉カテゴリの数: ${leafCategories.length}`);
    return leafCategories;
  } catch (error) {
    Logger.log(`カテゴリ取得エラー: ${error.message}`);
    return [];
  }
}
