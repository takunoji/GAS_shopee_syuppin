// バックエンド側のShopeeAuth.gs
// Shopee APIの設定を保存
function saveApiSettings(partnerId, partnerKey, redirectUrl) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("PARTNER_ID", partnerId);
  scriptProperties.setProperty("PARTNER_KEY", partnerKey);
  scriptProperties.setProperty("REDIRECT_URL", redirectUrl);
  return { success: true, message: "設定が保存されました。" };
}

// Shopeeの認証URLを生成
function getAuthorizationUrl() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const partnerId = scriptProperties.getProperty("PARTNER_ID");
  const partnerKey = scriptProperties.getProperty("PARTNER_KEY");
  const redirectUrl = encodeURIComponent(
    scriptProperties.getProperty("REDIRECT_URL")
  );

  if (!partnerId || !partnerKey || !redirectUrl) {
    throw new Error("API認証情報が設定されていません");
  }

  var path = "/api/v2/shop/auth_partner";
  var timestamp = Math.floor(Date.now() / 1000);
  var baseString = `${partnerId}${path}${timestamp}`;
  var sign = Utilities.computeHmacSha256Signature(baseString, partnerKey);
  sign = sign
    .map(function (char) {
      return ("0" + (char & 0xff).toString(16)).slice(-2);
    })
    .join("");

  const baseUrl = "https://partner.shopeemobile.com/api/v2/shop/auth_partner";
  const authUrl = `${baseUrl}?partner_id=${partnerId}&redirect=${redirectUrl}&timestamp=${timestamp}&sign=${sign}`;

  return { authUrl: authUrl };
}

// ショップデータを取得
function getShopsData() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const shopsData = JSON.parse(
    scriptProperties.getProperty("SHOPS_DATA") || "[]"
  );

  // 機密情報を削除して返す
  return shopsData.map((shop) => ({
    shop_id: shop.shop_id,
    region: shop.region,
    shop_name: shop.shop_name,
  }));
}

// 認可コードからアクセストークンとリフレッシュトークンを取得し、保存する。
function fetchTokens(authCode, shopId) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const partnerId = scriptProperties.getProperty("PARTNER_ID");
  const partnerKey = scriptProperties.getProperty("PARTNER_KEY");

  var path = "/api/v2/auth/token/get";
  var timestamp = Math.floor(Date.now() / 1000);
  var baseString = `${partnerId}${path}${timestamp}`;
  var sign = Utilities.computeHmacSha256Signature(baseString, partnerKey);
  sign = sign
    .map(function (char) {
      return ("0" + (char & 0xff).toString(16)).slice(-2);
    })
    .join("");

  const url = `https://partner.shopeemobile.com/api/v2/auth/token/get?partner_id=${partnerId}&timestamp=${timestamp}&sign=${sign}`;

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      code: authCode,
      partner_id: parseInt(partnerId),
      shop_id: parseInt(shopId),
    }),
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  if (data && data.access_token && data.refresh_token) {
    let shopData = JSON.parse(
      scriptProperties.getProperty("SHOPS_DATA") || "[]"
    );
    shopData = shopData.filter((shop) => shop.shop_id !== shopId); // 既存の同じshop_idを削除
    shopData.push({
      shop_id: shopId,
      access_token: data.access_token,
      refresh_token: data.refresh_token,
      region: "",
      shop_name: "",
    });
    scriptProperties.setProperty("SHOPS_DATA", JSON.stringify(shopData));
    console.log("ACCESS_TOKEN: " + data.access_token);
    console.log("REFRESH_TOKEN: " + data.refresh_token);
    console.log("トークン情報が保存されました。");
    return { success: true };
  } else {
    console.log("トークンの取得に失敗しました。");
    return { success: false, error: "トークンの取得に失敗しました" };
  }
}

function fetchShopInfo(shopId) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const partnerId = scriptProperties.getProperty("PARTNER_ID");
  const partnerKey = scriptProperties.getProperty("PARTNER_KEY");
  const shopsData = JSON.parse(
    scriptProperties.getProperty("SHOPS_DATA") || "[]"
  );
  const shop = shopsData.find((shop) => shop.shop_id === shopId);
  if (!shop) {
    console.log("指定されたShopIDのデータが見つかりません。");
    return {
      success: false,
      error: "指定されたShopIDのデータが見つかりません",
    };
  }
  const accessToken = shop.access_token;

  var path = "/api/v2/shop/get_shop_info";
  var timestamp = Math.floor(Date.now() / 1000);
  var baseString = `${partnerId}${path}${timestamp}${accessToken}${shopId}`;
  var sign = Utilities.computeHmacSha256Signature(baseString, partnerKey);
  sign = sign
    .map(function (char) {
      return ("0" + (char & 0xff).toString(16)).slice(-2);
    })
    .join("");

  const url = `https://partner.shopeemobile.com/api/v2/shop/get_shop_info?partner_id=${partnerId}&timestamp=${timestamp}&shop_id=${shopId}&access_token=${accessToken}&sign=${sign}`;

  var options = {
    method: "get",
    contentType: "application/json",
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  if (data && data.shop_name && data.region) {
    shop.shop_name = data.shop_name;
    shop.region = data.region;
    scriptProperties.setProperty("SHOPS_DATA", JSON.stringify(shopsData));
    console.log("SHOP_NAME: " + data.shop_name);
    console.log("REGION: " + data.region);
    console.log("ショップ名とリージョンが保存されました。");
    return { success: true, shop_name: data.shop_name, region: data.region };
  } else {
    console.log("ショップ情報の取得に失敗しました。");
    console.log(data);
    return { success: false, error: "ショップ情報の取得に失敗しました" };
  }
}
function refreshAccessToken(shopId) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const partnerId = scriptProperties.getProperty("PARTNER_ID");
  const partnerKey = scriptProperties.getProperty("PARTNER_KEY");
  const shopsData = JSON.parse(
    scriptProperties.getProperty("SHOPS_DATA") || "[]"
  );
  const shop = shopsData.find((shop) => shop.shop_id === shopId);
  if (!shop) {
    console.log("指定されたShopIDのデータが見つかりません。");
    return {
      success: false,
      error: "指定されたShopIDのデータが見つかりません",
    };
  }
  const refreshToken = shop.refresh_token;

  var path = "/api/v2/auth/access_token/get";
  var timestamp = Math.floor(Date.now() / 1000);
  var baseString = `${partnerId}${path}${timestamp}`;
  var sign = Utilities.computeHmacSha256Signature(baseString, partnerKey);
  sign = sign
    .map(function (char) {
      return ("0" + (char & 0xff).toString(16)).slice(-2);
    })
    .join("");

  const url = `https://partner.shopeemobile.com/api/v2/auth/access_token/get?partner_id=${partnerId}&timestamp=${timestamp}&sign=${sign}`;

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      refresh_token: refreshToken,
      partner_id: parseInt(partnerId),
      shop_id: parseInt(shopId),
    }),
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  if (data && data.access_token && data.refresh_token) {
    shop.access_token = data.access_token;
    shop.refresh_token = data.refresh_token;
    scriptProperties.setProperty("SHOPS_DATA", JSON.stringify(shopsData));
    console.log("ACCESS_TOKEN: " + data.access_token);
    console.log("REFRESH_TOKEN: " + data.refresh_token);
    console.log("トークン情報が更新されました。");
    return { success: true };
  } else {
    console.log("トークンの更新に失敗しました。");
    return { success: false, error: "トークンの更新に失敗しました" };
  }
}

function refreshAllTokens() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const shopsData = JSON.parse(
    scriptProperties.getProperty("SHOPS_DATA") || "[]"
  );

  const results = [];
  shopsData.forEach((shop) => {
    const result = refreshAccessToken(shop.shop_id);
    results.push({
      shop_id: shop.shop_id,
      shop_name: shop.shop_name,
      success: result.success,
      error: result.error || null,
    });
  });

  return results;
}

// 定期的にアクセストークンを更新するトリガーを設定する。
function setRefreshTrigger() {
  // トリガーがすでに存在する場合は削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === "refreshAllTokens") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 2時間ごとに実行されるトリガーを設定
  ScriptApp.newTrigger("refreshAllTokens").timeBased().everyHours(2).create();
  return { success: true };
}

// Web Apps用のdoGet関数（認証コールバック用）
function doGet(e) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const authCode = e.parameter.code;
  const shopId = e.parameter.shop_id;

  if (authCode && shopId) {
    fetchTokens(authCode, shopId);
    fetchShopInfo(shopId);
    setRefreshTrigger();
    return HtmlService.createHtmlOutput(
      "認可が成功しました。タブを閉じてください。"
    );
  } else {
    return HtmlService.createHtmlOutput(
      "認可に失敗しました。codeまたはshop_idがありません。"
    );
  }
}
