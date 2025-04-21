/**
 * スプレッドシートにブランドリストを保存します（重複を避ける）。
 */
function saveBrandListToSheet(sheet, categoryId, brandList) {
  // ブランドリストシートの既存データを取得
  const existingData = sheet.getDataRange().getValues();
  const brandCategoryMap = {}; // ブランドIDとカテゴリIDのマップ
  existingData.forEach((row, index) => {
    if (index === 0) return; // ヘッダーをスキップ
    const brandId = row[0]; // 1列目: ブランドID
    const originalBrand = row[1]; // 2列目: オリジナルブランド名
    const translatedBrand = row[2]; // 3列目: 和訳されたブランド名
    const categories = row
      .slice(3)
      .filter((v) => v)
      .join(",")
      .split(","); // 4列目以降: カテゴリIDリスト

    if (brandId && originalBrand) {
      const key = `${brandId}|${originalBrand}`;
      if (!brandCategoryMap[key]) {
        brandCategoryMap[key] = {
          categories: new Set(categories), // 重複排除のためSetを使用
          translatedBrand: translatedBrand || "", // 和訳がない場合は空文字
        };
      }
    }
  });

  // 新しいブランド情報をマップに追加または更新
  const untranslatedBrands = []; // 翻訳が必要なブランド
  brandList.forEach((brand) => {
    const key = `${brand.brand_id}|${brand.original_brand_name}`;
    if (!brandCategoryMap[key]) {
      untranslatedBrands.push(brand.original_brand_name); // 翻訳が必要なブランドを収集
      brandCategoryMap[key] = {
        categories: new Set(), // 初期値として空のSetを設定
        translatedBrand: "", // 初期値として空文字を設定
      };
    }
    brandCategoryMap[key].categories.add(String(categoryId)); // カテゴリIDを追加
  });

  // DeepLを使って未翻訳のブランド名を翻訳
  const translations = translateShopeeBrandsWithDeepL(untranslatedBrands);

  // 翻訳結果をブランドマップに反映
  brandList.forEach((brand) => {
    const key = `${brand.brand_id}|${brand.original_brand_name}`;
    if (!brandCategoryMap[key].translatedBrand) {
      brandCategoryMap[key].translatedBrand =
        translations[brand.original_brand_name] || ""; // 翻訳結果を設定
    }
  });

  // シート用データを整形
  const rowsToWrite = [];
  for (const key in brandCategoryMap) {
    const [brandId, originalBrandName] = key.split("|");
    const categories = Array.from(brandCategoryMap[key].categories); // Setを配列に変換
    const translatedBrand = brandCategoryMap[key].translatedBrand;

    // カテゴリIDを32000文字ごとに分割し、列ごとに保存
    const categoryColumns = [];
    let currentString = "";
    categories.forEach((category) => {
      const categoryString = String(category); // 数字を文字列に変換
      if (currentString.length + categoryString.length + 1 > 32000) {
        categoryColumns.push(currentString);
        currentString = categoryString;
      } else {
        currentString += (currentString ? "," : "") + categoryString;
      }
    });
    if (currentString) categoryColumns.push(currentString);

    // 1行データとして追加
    rowsToWrite.push([
      brandId,
      originalBrandName,
      translatedBrand,
      ...categoryColumns,
    ]);
  }

  // シートのヘッダーを確認または追加
  const header = [
    "ブランドID",
    "オリジナルブランド名",
    "和訳されたブランド名",
    "カテゴリID (分割保存)",
  ];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(header);
  }

  // データ書き込み
  rowsToWrite.forEach((row) => {
    const existingRowIndex = existingData.findIndex((existingRow) => {
      return (
        String(existingRow[0]) === String(row[0]) &&
        String(existingRow[1]) === String(row[1])
      );
    });
    if (existingRowIndex > 0) {
      // 既存の行を更新
      const updatedRow = [
        row[0], // ブランドID
        row[1], // オリジナルブランド名
        row[2], // 和訳されたブランド名
        ...row.slice(3), // カテゴリID
      ];
      sheet
        .getRange(existingRowIndex + 1, 1, 1, updatedRow.length)
        .setValues([updatedRow]);
    } else {
      // 新しい行を追加
      sheet.appendRow(row);
    }
  });

  Logger.log(`${rowsToWrite.length} 件のブランド情報を更新しました。`);
}

/**
 * スプレッドシートの行数が不足している場合に、必要な分だけ行を追加する関数を追加します。
 */
function ensureSheetHasEnoughRows(sheet, additionalRows) {
  const currentRows = sheet.getMaxRows();
  const requiredRows = sheet.getLastRow() + additionalRows;
  if (currentRows < requiredRows) {
    sheet.insertRowsAfter(currentRows, requiredRows - currentRows);
    Logger.log(`Added ${requiredRows - currentRows} rows to the sheet.`);
  }
}

/**
 * Levenshtein距離を計算する関数
 * @param {string} a - 比較する文字列1
 * @param {string} b - 比較する文字列2
 * @returns {number} - Levenshtein距離
 */
function levenshteinDistance(a, b) {
  // 入力が数値であれば文字列に変換
  a = a.toString();
  b = b.toString();

  // 空文字列の場合の特別処理
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;

  const matrix = [];
  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // 置換
          Math.min(
            matrix[i][j - 1] + 1, // 挿入
            matrix[i - 1][j] + 1
          ) // 削除
        );
      }
    }
  }
  return matrix[b.length][a.length];
}

// スプレッドシートからShopeeブランドリストを取得し、カテゴリーIDでフィルタリング
function getShopeeBrandsFromSheetByCategory(categoryId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    BRAND_CACHE_SHEET_NAME
  );
  const data = sheet.getDataRange().getValues(); // シートのすべてのデータを取得
  const matchingBrands = [];

  data.forEach((row) => {
    const brandId = row[0]; // 2列目 (brand_id)
    const originalBrandName = row[1]; // 3列目 (original_brand_name)
    const translatedBrandName = row[2]; // 4列目 (translated_brand_name)

    // 4列目以降のカテゴリID列を取得し、値を分割してリスト化
    const categoryIdColumns = row
      .slice(3)
      .filter((value) => value) // 空セルを除外
      .flatMap((value) =>
        String(value)
          .split(",")
          .map((id) => id.trim())
      ); // カンマ区切りの値を分割してリスト化

    // カテゴリIDが一致する場合にブランド情報を収集
    if (categoryIdColumns.includes(String(categoryId).trim())) {
      matchingBrands.push({
        brand_id: brandId,
        original_brand: originalBrandName,
        translated_brand: translatedBrandName || "",
      });
    }
  });

  Logger.log(
    `取得したブランド数 (カテゴリーID: ${categoryId}): ${matchingBrands.length}`
  );
  return matchingBrands;
}

/**
 * 全shopeeブランドを取得する
 */
function fetchAllShopeeBrands(shopId, accessToken, categoryId) {
  const pageSize = 100; // 一度に取得するブランド数
  let offset = 0;
  let hasMore = true;
  const allBrands = [];

  while (hasMore) {
    const brandList = fetchBrandListFromShopee(
      shopId,
      accessToken,
      categoryId,
      offset,
      pageSize
    );
    if (!brandList || brandList.length === 0) break;

    allBrands.push(...brandList);
    hasMore = brandList.has_next_page;
    offset += pageSize;
  }

  // 翻訳処理を実行
  translateShopeeBrandsWithDeepL(allBrands);

  return allBrands;
}

/**
 * Shopee 全カテゴリのブランドリストを取得してスプレッドシートに保存
 */
function fetchAllCategoriesAndBrands(selectedShopsIndices) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ブランドリスト");
  if (!sheet) {
    throw new Error("「ブランドリスト」という名前のシートを作成してください。");
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  const shopsData = JSON.parse(scriptProperties.getProperty("SHOPS_DATA"));

  selectedShopsIndices.forEach((index) => {
    try {
      const shop = shopsData[index];
      const shopId = shop.shop_id;
      const accessToken = shop.access_token;

      if (!shopId || !accessToken) {
        throw new Error("SHOP_IDまたはACCESS_TOKENが設定されていません。");
      }

      const categories = fetchAllCategories(shopId, accessToken); // 全カテゴリを取得
      categories.forEach((category) => {
        const categoryId = category.category_id;
        const categoryName = category.category_name;

        Logger.log(`カテゴリID: ${categoryId}, カテゴリ名: ${categoryName}`);
        fetchBrandListFromShopee(shopId, accessToken, categoryId); // ブランドリストを取得
      });
    } catch (error) {
      Logger.log(`エラー: ${error.message}`);
    }
  });

  Logger.log("全カテゴリのブランドIDの取得が完了しました。");
}

/**
 * Brandの一致度✅
 */
function matchBrandWithShopee(keepaBrand, shopeeBrands) {
  let nearestBrand = null;
  let minDistance = Infinity;

  // Keepaブランドを正規化
  const normalizedKeepaBrand = normalizeString(keepaBrand);

  shopeeBrands.forEach((brand) => {
    // Shopeeブランドの各フィールドを正規化
    const normalizedOriginalBrand = normalizeString(brand.original_brand || "");
    const normalizedTranslatedBrand = normalizeString(
      brand.translated_brand || ""
    );

    // Levenshtein距離を計算
    const originalDistance = levenshteinDistance(
      normalizedKeepaBrand,
      normalizedOriginalBrand
    );
    const translatedDistance = levenshteinDistance(
      normalizedKeepaBrand,
      normalizedTranslatedBrand
    );

    const distance = Math.min(originalDistance, translatedDistance);
    if (distance < minDistance) {
      minDistance = distance;
      nearestBrand = brand;
    }
  });

  return nearestBrand;
}

function normalizeString(str) {
  if (typeof str !== "string") {
    // strが文字列でない場合は空文字列に置き換え
    return "";
  }
  return str
    .toLowerCase() // 小文字化
    .replace(/\s+/g, "") // 空白削除
    .replace(/[^\w]/g, ""); // 非英数字削除
}
