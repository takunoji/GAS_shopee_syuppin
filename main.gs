function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Shopee設定")
    .addItem("事前設定", "showSettingsForm")
    .addItem("Shopee認証", "getAuthorizationUrl")
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

////////////////
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
////////////////

function getShopsData() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const shopsData = JSON.parse(scriptProperties.getProperty("SHOPS_DATA"));
  return shopsData;
}
