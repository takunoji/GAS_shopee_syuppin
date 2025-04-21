// シート名を定義
const SHEET_PRODUCTS = "出品一覧"; // 出品一覧シート
const SHEET_SETTINGS = "基本設定"; // 基本設定シート
const SHEET_ADDITIONAL = "追加設定"; // 追加設定シート
const SHEET_SLS_SHIPPING = "SLS送料表"; // SLS送料表シート
const SHEET_HISTORY = "出品履歴"; // 出品履歴シート
const BRAND_CACHE_SHEET_NAME = "ブランドリスト";
// const SHEET_PRODUCTS = "商品一覧";

// Shopee APIのエンドポイントや定数を定義。

const PROPERTIES = PropertiesService.getScriptProperties();
const PARTNER_KEY = PROPERTIES.getProperty("PARTNER_KEY");
const PARTNER_ID = PROPERTIES.getProperty("PARTNER_ID");
const API_URL = "https://partner.shopeemobile.com";
const API_URL_BR = "https://openplatform.shopee.com.br";
const CATEGORY_RECOMMEND_API_PATH = "/api/v2/product/category_recommend";
const ATTRIBUTES_API_PATH = "/api/v2/product/get_attributes";
const CHANNEL_LIST_API_PATH = "/api/v2/logistics/get_channel_list";
const ITEM_LIMIT_API_PATH = "/api/v2/product/get_item_limit";
const ITEM_LIST_API_PATH = "/api/v2/product/get_item_list";
const ITEM_BASE_API_PATH = "/api/v2/product/get_item_base_info";
const ITEM_EXTRA_API_PATH = "/api/v2/product/get_item_extra_info";
const MODEL_LIST_API_PATH = "/api/v2/product/get_model_list";
const UPLOAD_IMAGE_API_PATH = "/api/v2/media_space/upload_image";
const ADD_ITEM_API_PATH = "/api/v2/product/add_item";
const INIT_TIER_VARIATION_API_PATH = "/api/v2/product/init_tier_variation";
const BRAND_LIST_API_PATH = "/api/v2/product/get_brand_list";

// 出品一覧シートの列番号
const COL_CHECKBOX = 1; // A列: チェックボックス
const COL_ASIN = 2; // B列: ASIN
const COL_DATA_DATE = 3; // C列: データ登録日
const COL_MAIN_PRODUCT_NAME_JP = 9; // I列: 商品名 (日本語)
const COL_MAIN_DESCRIPTION_JP = 10; // J列: 商品説明 (日本語)
const COL_MAIN_PRODUCT_NAME_EN = 11; // I列: 商品名 (日本語)
const COL_MAIN_DESCRIPTION_EN = 12; // J列: 商品説明 (日本語)
const COL_MAIN_PRODUCT_NAME_ZH = 13; // I列: 商品名 (日本語)
const COL_MAIN_DESCRIPTION_ZH = 14; // J列: 商品説明 (日本語)
const COL_MAIN_PRICE = 4; // D列: 価格
const COL_MAIN_WEIGHT = 5; // E列: 重量
const COL_MAIN_LENGTH = 6; // F列: 長さ
const COL_MAIN_WIDTH = 7; // G列: 幅
const COL_MAIN_HEIGHT = 8; // H列: 高さ
const COL_MAIN_IMAGE_1 = 15; // O列: 画像1
const COL_MAIN_IMAGE_1_DISPLAY = 23; // W列: イメージ1表示

const COL_MAIN_PRODUCT_NAME_EN_CON = 46; // 商品名(決定)
const COL_MAIN_DESCRIPTION_EN_CON = 47; // 商品名(決定)
const COL_MAIN_PRODUCT_NAME_ZH_CON = 48; // 商品名(決定)
const COL_MAIN_DESCRIPTION_ZH_CON = 49; // 商品名(決定)

const COL_SKU = 31; // CU列: SKU
const COL_STOCK = 54; // DH列: 在庫数

const COL_SET_DAY = 64; // ブランド(keepa)
const COL_BRAND_FROM_KEEPA = 65; // ブランド(keepa)
const COL_MATCHED_BRAND_ID = 66; // ブランドID
const COL_VARIATION_TYPE = 67; // BK列: バリエーションタイプ
const COL_VARIATION1_NAME = 68; // BM列: バリエーション1の名前
const COL_VARIATION1_VALUES_START = 71; // BN列: バリエーション1の値の開始列
const COL_VARIATION1_VALUES_END = 100; // BN列: バリエーション1の値の開始列
const COL_VARIATION2_NAME = 101; // BX列: バリエーション2の名前2
const COL_VARIATION2_VALUES_START = 103; // BY列: バリエーション2の値の開始列
const COL_VARIATION2_VALUES_END = 133; // BY列: バリエーション2の値の開始列
const COL_VARIATION_OUTPUT_START = 134;

// バリエーションごとの商品数
const VARIATION_COLUMNS = 10; // 各バリエーションの列数 (ASIN、重量、長さ、幅、高さ、名前、画像など)
const VARIATION_DATAS = 42;
// 各列における位置
const VARIATION_COL = 0; // バリエーション組み合わせの相対位置
const VARIATION_ASIN = 1; // ASIN列の相対位置
const VARIATION_PRICE = 2; // 新品最安値列の相対位置
const VARIATION_WEIGHT = 3; // 重量列の相対位置
const VARIATION_LENGTH = 4; // 長さ列の相対位置
const VARIATION_WIDTH = 5; // 幅列の相対位置
const VARIATION_HEIGHT = 6; // 高さ列の相対位置
const VARIATION_NAME = 7; // バリエーション名列の相対位置
const VATIATION_DESCRIPTION_JP = 8; // 商品説明
const VARIATION_PRODUCT_NAME_EN = 9; // 商品名 (英) の相対位置
const VATIATION_DESCRIPTION_EN = 10; // 商品説明
const VARIATION_PRODUCT_NAME_CH = 11; // 商品名 (中) の相対位置
const VATIATION_DESCRIPTION_ZH = 12; // 商品説明
const VARIATION_IMAGE_ID = 13; // 画像ID列の相対位置
const VARIATION_IMAGE = 14; // 画像列の相対位置
const VARIATION_SKU = 15; // SKUの相対位置
const VARIATION_PRICE_SG = 16; // 商品価格(SG)の相対位置
const VARIATION_PRICE_SG_YEN = 23; // 商品価格(SG)YENの相対位置

const VARIATION_PRODUCT_NAME_EN_CON = 30; // 商品名 (英) の相対位置
const VATIATION_DESCRIPTION_EN_CON = 31; // 商品説明
const VARIATION_PRODUCT_NAME_CH_CON = 32; // 商品名 (中) の相対位置
const VATIATION_DESCRIPTION_ZH_CON = 33; // 商品説明
const VARIATION_WEIGHT_CON = 34; // 重量列の相対位置
const VARIATION_LENGTH_CON = 35; // 長さ列の相対位置
const VARIATION_WIDTH_CON = 36; // 幅列の相対位置
const VARIATION_HEIGHT_CON = 37; // 高さ列の相対位置
const VARIATION_ZAIKO_CON = 38; // 在庫数の相対位置

const VARIATION_IMAGE_CON = 39; // 画像ID(決定)の相対位置
const VARIATION_IMAGE_ID_CON = 40; // 画像ID(決定)の相対位置
const VARIATION_DATE_CON = 41; // 登録日の相対位置
const MAX_VARIATIONS = 20; // 最大バリエーション数 (例)

const COL_PRODUCT_DESCRIPTION = 3; // 商品説明列
const COL_VARIATION1_NAME_JP = 5; // バリエーション1名 (日本語)
const COL_VARIATION2_NAME_JP = 6; // バリエーション2名 (日本語)
const COL_VARIATION1_NAME_EN = 7; // バリエーション1名 (英語)
const COL_VARIATION2_NAME_EN = 8; // バリエーション2名 (英語)
const COL_VARIATION1_NAME_ZH = 9; // バリエーション1名 (簡体語)
const COL_VARIATION2_NAME_ZH = 10; // バリエーション2名 (簡体語)

// 基本設定シートの列番号
const COL_API_KEY = 3; // C4: APIキー
const COL_FOLDER_ID = 5; // C5: Google Drive フォルダID
const COL_INIT_WEIGHT = 7; // C7: 初期重量
const COL_ADD_WEIGHT = 8; // C8: 重量加算
const COL_INIT_STOCK = 9; // C9: 初期在庫
const COL_PROFIT_RATIO = 11; // C11: 利益率
const COL_PAYONEER_RATIO = 15; // C15: Payoneer手数料
const COL_PROMO_RATIO = 16; // C16: プロモーション手数料
const COL_VOUCHER_RATIO = 17; // C17: バウチャー手数料
