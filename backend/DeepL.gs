// DeepL APIを使ってブランドリストを翻訳し、ブランドリストシートに書き込む
// function translateShopeeBrandsWithDeepL(brandList) {
//   const brandListSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ブランドリスト');
//   if (!brandListSheet) {
//     throw new Error('ブランドリストシートが存在しません。');
//   }

//   const lastRow = brandListSheet.getLastRow();

//   // 翻訳列（オリジナルブランドの右隣、E列を使用）
//   const translatedColumnIndex = 5; // 5列目（E列）を翻訳列とする
//   const originalBrandColumnIndex = 4; // オリジナルブランドは4列目（D列）

//   // D列（オリジナルブランド）とE列（翻訳済み列）を取得
//   const originalBrands = brandListSheet
//     .getRange(2, originalBrandColumnIndex, lastRow - 1, 1)
//     .getValues()
//     .map(row => row[0]);

//   const translatedBrands = brandListSheet
//     .getRange(2, translatedColumnIndex, lastRow - 1, 1)
//     .getValues()
//     .map(row => row[0]);

//   // DeepL APIを利用して翻訳（未翻訳のみ処理）
//   originalBrands.forEach((originalBrand, index) => {
//     if (originalBrand && !translatedBrands[index]) {
//       // E列が空の場合のみ翻訳を実行
//       const translatedText = translateTextWithDeepL(originalBrand, 'JA'); // 翻訳先を日本語に設定
//       translatedBrands[index] = translatedText;

//       // 翻訳結果をE列に書き込む
//       brandListSheet.getRange(index + 2, translatedColumnIndex).setValue(translatedText);
//     }
//   });

//   // ブランドリストを返す（ブランドID、オリジナルブランド名、翻訳されたブランド名）
//   return brandList.map((brand, index) => ({
//     brand_id: brand[1], // brand_id
//     original_brand: brand[0], // 元のブランド名
//     translated_brand: translatedBrands[index] || "" // 翻訳されたブランド名
//   }));
// }

/**
 * DeepLを使ってブランド名を翻訳する関数
 * @param {Array} originalBrands - 翻訳するブランド名の配列
 * @return {Object} - 元のブランド名をキーとし、翻訳結果を値とするオブジェクト
 */
function translateShopeeBrandsWithDeepL(originalBrands) {
  const translations = {};
  originalBrands.forEach((brand) => {
    try {
      const translatedText = translateTextWithDeepL(brand, "JA"); // 翻訳先を日本語に設定
      translations[brand] = translatedText;
    } catch (error) {
      Logger.log(
        `ブランド名 "${brand}" の翻訳に失敗しました: ${error.message}`
      );
      translations[brand] = ""; // 翻訳に失敗した場合は空文字を設定
    }
  });
  return translations;
}

/**
 * DeepL APIを使った翻訳処理
 * @param {string} text - 翻訳する文字列
 * @param {string} targetLang - 翻訳先の言語コード（例: 'JA'）
 * @return {string} - 翻訳された文字列
 */
function translateTextWithDeepL(text, targetLang) {
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("DEEPL_API_KEY");
  if (!apiKey) {
    throw new Error("DeepL APIキーが設定されていません。");
  }

  const url = `https://api-free.deepl.com/v2/translate`;
  const payload = {
    auth_key: apiKey,
    text: text,
    target_lang: targetLang,
  };

  const options = {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload: payload,
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseData = JSON.parse(response.getContentText());

  if (responseData.translations && responseData.translations.length > 0) {
    return responseData.translations[0].text;
  } else {
    Logger.log(`DeepL翻訳エラー: ${response.getContentText()}`);
    return ""; // 翻訳に失敗した場合は空文字を返す
  }
}

/**
 * 翻訳処理
 */
function translateTextWithDeepL(text, targetLang) {
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("DEEPL_API_KEY");
  if (!apiKey) {
    throw new Error("DeepL APIキーが設定されていません。");
  }

  const url = `https://api-free.deepl.com/v2/translate`;
  const options = {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload: {
      auth_key: apiKey,
      text: text,
      target_lang: targetLang,
    },
    muteHttpExceptions: true,
  };

  const maxRetries = 3; // 最大リトライ回数
  let attempts = 0;

  while (attempts < maxRetries) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseData = JSON.parse(response.getContentText());

      if (responseData.translations && responseData.translations.length > 0) {
        return responseData.translations[0].text;
      } else {
        Logger.log(`DeepL翻訳エラー: ${response.getContentText()}`);
        return ""; // 翻訳に失敗した場合は空文字を返す
      }
    } catch (error) {
      attempts++;
      Logger.log(
        `DeepL APIリクエストエラー (試行 ${attempts}/${maxRetries}): ${error.message}`
      );
      if (attempts >= maxRetries) {
        throw new Error("DeepL APIリクエストが最大試行回数に達しました。");
      }
      Utilities.sleep(2000); // 再試行の前に2秒待機
    }
  }
}
