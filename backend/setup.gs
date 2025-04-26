// バックエンド側のsetup.gs
// 初期設定関数

// APIキーなどのデフォルト設定を行う
function setupDefaultProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();

  // すでに設定済みの場合は上書きしない（任意）
  if (!scriptProperties.getProperty("INIT_WEIGHT")) {
    scriptProperties.setProperty("INIT_WEIGHT", DEFAULT_INIT_WEIGHT.toString());
  }
  if (!scriptProperties.getProperty("ADD_WEIGHT")) {
    scriptProperties.setProperty("ADD_WEIGHT", DEFAULT_ADD_WEIGHT.toString());
  }
  if (!scriptProperties.getProperty("INIT_STOCK")) {
    scriptProperties.setProperty("INIT_STOCK", DEFAULT_INIT_STOCK.toString());
  }
  if (!scriptProperties.getProperty("PROFIT_RATIO")) {
    scriptProperties.setProperty(
      "PROFIT_RATIO",
      DEFAULT_PROFIT_RATIO.toString()
    );
  }

  // APIキー用（初期値は空）
  if (!scriptProperties.getProperty("API_KEY")) {
    scriptProperties.setProperty("API_KEY", "YOUR_API_KEY");
  }

  // 他の必要な初期設定があれば追加
}

// 初期設定関数
function initializeBackend() {
  setupDefaultProperties();

  // プロパティの確認と表示
  const properties = PropertiesService.getScriptProperties().getProperties();
  console.log("現在の設定:");
  for (const key in properties) {
    if (
      key.includes("KEY") ||
      key.includes("TOKEN") ||
      key.includes("SECRET")
    ) {
      console.log(`${key}: ******** (機密情報)`);
    } else {
      console.log(`${key}: ${properties[key]}`);
    }
  }

  // デプロイURLの表示
  console.log(`バックエンドURL: ${ScriptApp.getService().getUrl()}`);
  return "バックエンドの初期化が完了しました。";
}

// Web Apps用のトップページ
function doGet() {
  return HtmlService.createHtmlOutput(`
    <h1>Shopee Integration Backend</h1>
    <p>このサービスはフロントエンドアプリケーションからのAPI呼び出しを受け付けます。</p>
    <p>現在の時刻: ${new Date().toLocaleString()}</p>
  `);
}
