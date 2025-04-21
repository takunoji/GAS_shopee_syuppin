/**
 *  Shopee APIの認証・設定機能。
 */

/**
 * Shopee APIの設定（パートナーID、キー、リダイレクトURL）を入力するフォームを表示。
 */
function showSettingsForm() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const template = HtmlService.createTemplateFromFile("SettingsForm");
  template.partnerId = scriptProperties.getProperty("PARTNER_ID") || "";
  template.partnerKey = scriptProperties.getProperty("PARTNER_KEY") || "";
  template.redirectUrl = scriptProperties.getProperty("REDIRECT_URL") || "";
  const htmlOutput = template.evaluate().setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Shopee API 設定");
}

/**
 * 入力されたAPI設定をスクリプトプロパティに保存。
 */
function setScriptPropertiesFromForm(data) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("PARTNER_ID", data.partnerId);
  scriptProperties.setProperty("PARTNER_KEY", data.partnerKey);
  scriptProperties.setProperty("REDIRECT_URL", data.redirectUrl);
  return "設定が保存されました。";
}

/**
 * Shopeeの認証URLを生成し、ユーザーがAPI認可を行うためのリンクを表示。
 */
function getAuthorizationUrl() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const partnerId = scriptProperties.getProperty("PARTNER_ID");
  const partnerKey = scriptProperties.getProperty("PARTNER_KEY");
  const redirectUrl = encodeURIComponent(
    scriptProperties.getProperty("REDIRECT_URL")
  );

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
  Logger.log("認可URL: " + authUrl);

  const ui = SpreadsheetApp.getUi();
  const htmlOutput = HtmlService.createHtmlOutput(
    `<p>アプリを認可するには<a href="${authUrl}" target="_blank">こちらをクリック</a>してください。</p>`
  )
    .setWidth(250)
    .setHeight(100);
  ui.showModalDialog(htmlOutput, "認可");
}

function doGet(e) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const authCode = e.parameter.code;
  const shopId = e.parameter.shop_id; // ShopeeのショップID

  if (authCode && shopId) {
    fetchTokens(authCode, shopId); // トークン取得関数を呼び出し
    fetchShopInfo(shopId);
    setRefreshTrigger(); // トリガーの設定を追加
    return ContentService.createTextOutput(
      "認可が成功しました。タブを閉じてください。"
    );
  } else {
    return ContentService.createTextOutput(
      "認可に失敗しました。codeまたはshop_idがありません。"
    );
  }
}

/**
 * 認可コードからアクセストークンとリフレッシュトークンを取得し、保存する。
 */
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
    Logger.log("ACCESS_TOKEN: " + data.access_token);
    Logger.log("REFRESH_TOKEN: " + data.refresh_token);
    Logger.log("トークン情報が保存されました。");
  } else {
    Logger.log("トークンの取得に失敗しました。");
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
    Logger.log("指定されたShopIDのデータが見つかりません。");
    return;
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
    Logger.log("SHOP_NAME: " + data.shop_name);
    Logger.log("REGION: " + data.region);
    Logger.log("ショップ名とリージョンが保存されました。");
  } else {
    Logger.log("ショップ情報の取得に失敗しました。");
    Logger.log(data);
  }
}

/**
 * リフレッシュトークンを利用して新しいアクセストークンを取得し、更新する。
 */
function refreshAccessToken(shopId) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const partnerId = scriptProperties.getProperty("PARTNER_ID");
  const partnerKey = scriptProperties.getProperty("PARTNER_KEY");
  const shopsData = JSON.parse(
    scriptProperties.getProperty("SHOPS_DATA") || "[]"
  );
  const shop = shopsData.find((shop) => shop.shop_id === shopId);
  if (!shop) {
    Logger.log("指定されたShopIDのデータが見つかりません。");
    return;
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
    Logger.log("ACCESS_TOKEN: " + data.access_token);
    Logger.log("REFRESH_TOKEN: " + data.refresh_token);
    Logger.log("トークン情報が更新されました。");
  } else {
    Logger.log("トークンの更新に失敗しました。");
  }
}

function refreshAllTokens() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const shopsData = JSON.parse(
    scriptProperties.getProperty("SHOPS_DATA") || "[]"
  );

  shopsData.forEach((shop) => {
    refreshAccessToken(shop.shop_id);
  });
}

/**
 * 定期的にアクセストークンを更新するトリガーを設定する。
 */
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
}
