/**
 * バリエーション品の価格Mapを計算
 */
function priceCalcVariation(row, startcol) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出品一覧");
  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("基本設定");
  const slsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SLS送料表");

  // 各シートから必要な情報を取得
  const initWeight = parseFloat(settingsSheet.getRange("C7").getValue()); // 初期重量（未取得時）
  const addWeight = parseFloat(settingsSheet.getRange("C8").getValue()); // 重量加算
  const initLength = parseFloat(settingsSheet.getRange("I7").getValue()); // 初期長さ（未取得時）
  const initWidth = parseFloat(settingsSheet.getRange("I8").getValue()); // 初期幅（未取得時）
  const initHeight = parseFloat(settingsSheet.getRange("I9").getValue()); // 初期高さ（未取得時）
  const profitRatio = settingsSheet.getRange("C11").getValue(); // 利益率
  const calcWeight =
    sheet.getRange(row, startcol + VARIATION_WEIGHT).getValue() === "N/A"
      ? initWeight
      : parseFloat(
          sheet.getRange(row, startcol + VARIATION_WEIGHT).getValue()
        ) + addWeight; // 計算用重量 (kg)
  const calcLength =
    sheet.getRange(row, startcol + VARIATION_LENGTH).getValue() === "N/A"
      ? initLength
      : parseFloat(sheet.getRange(row, startcol + VARIATION_LENGTH).getValue());
  const calcWitdh =
    sheet.getRange(row, startcol + VARIATION_WIDTH).getValue() === "N/A"
      ? initWidth
      : parseFloat(sheet.getRange(row, startcol + VARIATION_WIDTH).getValue());
  const calcHeight =
    sheet.getRange(row, startcol + VARIATION_HEIGHT).getValue() === "N/A"
      ? initHeight
      : parseFloat(sheet.getRange(row, startcol + VARIATION_HEIGHT).getValue());
  const calcSize = calcLength + calcWitdh + calcHeight;
  const domCost = getDomShippingCost(calcSize, calcWeight); // 国内送料を取得
  const fixedCost =
    parseInt(domCost ? domCost : settingsSheet.getRange("C13").getValue()) +
    parseInt(settingsSheet.getRange("C14").getValue()); // 固定費(国内送料+その他経費)
  const payoneerRatio = settingsSheet.getRange("C15").getValue(); // Payoneer手数料
  const promoRatio = settingsSheet.getRange("C16").getValue(); // プロモーション割引
  const voucherRatio = settingsSheet.getRange("C17").getValue(); // バウチャー割引

  const rateMap = {}; // 各国レート
  rateMap["SG"] = parseFloat(settingsSheet.getRange("F22").getValue()); // SGレート
  rateMap["MY"] = parseFloat(settingsSheet.getRange("F23").getValue()); // MYレート
  rateMap["PH"] = parseFloat(settingsSheet.getRange("F24").getValue()); // PHレート
  rateMap["TH"] = parseFloat(settingsSheet.getRange("F25").getValue()); // THレート
  rateMap["TW"] = parseFloat(settingsSheet.getRange("F26").getValue()); // TWレート
  rateMap["VN"] = parseFloat(settingsSheet.getRange("F27").getValue()); // VNレート
  rateMap["BR"] = parseFloat(settingsSheet.getRange("F28").getValue()); // VNレート

  const deliveryMap = {}; // 各国配送設定
  deliveryMap["SG"] = parseInt(settingsSheet.getRange("C32").getValue());
  deliveryMap["MY"] = parseInt(settingsSheet.getRange("C33").getValue());
  deliveryMap["PH"] = parseInt(settingsSheet.getRange("C34").getValue());
  deliveryMap["TH"] = parseInt(settingsSheet.getRange("C35").getValue());
  deliveryMap["TW"] = parseInt(settingsSheet.getRange("C36").getValue());
  deliveryMap["VN"] = parseInt(settingsSheet.getRange("C37").getValue());
  deliveryMap["BR"] = parseInt(settingsSheet.getRange("C38").getValue());

  const salesMap = {}; // 各国販売手数料
  salesMap["SG"] = settingsSheet.getRange("F32").getValue();
  salesMap["MY"] = settingsSheet.getRange("G32").getValue();
  salesMap["PH"] = settingsSheet.getRange("H32").getValue();
  salesMap["TH"] = settingsSheet.getRange("I32").getValue();
  salesMap["TW"] = settingsSheet.getRange("J32").getValue();
  salesMap["VN"] = settingsSheet.getRange("K32").getValue();
  salesMap["BR"] = settingsSheet.getRange("L32").getValue();

  const paymentMap = {}; // 各国決済手数料
  paymentMap["SG"] = settingsSheet.getRange("F33").getValue();
  paymentMap["MY"] = settingsSheet.getRange("G33").getValue();
  paymentMap["PH"] = settingsSheet.getRange("H33").getValue();
  paymentMap["TH"] = settingsSheet.getRange("I33").getValue();
  paymentMap["TW"] = settingsSheet.getRange("J33").getValue();
  paymentMap["VN"] = settingsSheet.getRange("K33").getValue();
  paymentMap["BR"] = settingsSheet.getRange("L33").getValue();

  const ccbfssMap = {}; // CCB/FSS
  ccbfssMap["SG"] = settingsSheet.getRange("F34").getValue();
  ccbfssMap["MY"] = settingsSheet.getRange("G34").getValue();
  ccbfssMap["PH"] = settingsSheet.getRange("H34").getValue();
  ccbfssMap["TH"] = settingsSheet.getRange("I34").getValue();
  ccbfssMap["TW"] = settingsSheet.getRange("J34").getValue();
  ccbfssMap["VN"] = settingsSheet.getRange("K34").getValue();
  ccbfssMap["BR"] = settingsSheet.getRange("L34").getValue();

  const mdvMap = {}; // MDV
  mdvMap["SG"] = settingsSheet.getRange("F35").getValue();
  mdvMap["MY"] = settingsSheet.getRange("G35").getValue();
  mdvMap["PH"] = settingsSheet.getRange("H35").getValue();
  mdvMap["TH"] = settingsSheet.getRange("I35").getValue();
  mdvMap["TW"] = settingsSheet.getRange("J35").getValue();
  mdvMap["VN"] = settingsSheet.getRange("K35").getValue();
  mdvMap["BR"] = settingsSheet.getRange("L35").getValue();

  const amazonPrice = parseInt(
    sheet.getRange(row, startcol + VARIATION_PRICE).getValue()
  ); // 新品最安値
  const calcWeightG = calcWeight * 1000; // 計算用重量 (g)
  const slsMap = getSlsShippingCost(calcWeightG); // 各国SLS送料

  const totalCommMap = {}; // 換算後合計手数料
  totalCommMap["SG"] =
    (1 - voucherRatio) * (1 - salesMap["SG"] - ccbfssMap["SG"] - mdvMap["SG"]) -
    paymentMap["SG"];
  totalCommMap["MY"] =
    (1 - voucherRatio) * (1 - salesMap["MY"] - ccbfssMap["MY"] - mdvMap["MY"]) -
    paymentMap["MY"];
  totalCommMap["PH"] =
    (1 - voucherRatio) * (1 - salesMap["PH"] - ccbfssMap["PH"] - mdvMap["PH"]) -
    paymentMap["PH"];
  totalCommMap["TH"] =
    (1 - voucherRatio) * (1 - salesMap["TH"] - ccbfssMap["TH"] - mdvMap["TH"]) -
    paymentMap["TH"];
  totalCommMap["TW"] =
    (1 - voucherRatio) * (1 - salesMap["TW"] - ccbfssMap["TW"] - mdvMap["TW"]) -
    paymentMap["TW"];
  totalCommMap["VN"] =
    (1 - voucherRatio) * (1 - salesMap["VN"] - ccbfssMap["VN"] - mdvMap["VN"]) -
    paymentMap["VN"];
  totalCommMap["BR"] =
    (1 - voucherRatio) * (1 - salesMap["BR"] - ccbfssMap["BR"] - mdvMap["BR"]) -
    paymentMap["BR"];

  const priceMap = {}; // 出品価格
  priceMap["SG"] =
    Math.round(
      ((slsMap["SG"] * (1 - payoneerRatio) +
        (amazonPrice + fixedCost) / rateMap["SG"]) /
        (totalCommMap["SG"] * (1 - payoneerRatio) -
          profitRatio * (1 - voucherRatio)) /
        (1 - promoRatio)) *
        100
    ) / 100;
  priceMap["MY"] =
    Math.round(
      ((slsMap["MY"] * (1 - payoneerRatio) +
        (amazonPrice + fixedCost) / rateMap["MY"]) /
        (totalCommMap["MY"] * (1 - payoneerRatio) -
          profitRatio * (1 - voucherRatio)) /
        (1 - promoRatio)) *
        100
    ) / 100;
  priceMap["PH"] = Math.round(
    (slsMap["PH"] * (1 - payoneerRatio) +
      (amazonPrice + fixedCost) / rateMap["PH"]) /
      (totalCommMap["PH"] * (1 - payoneerRatio) -
        profitRatio * (1 - voucherRatio)) /
      (1 - promoRatio)
  );
  priceMap["TH"] = Math.round(
    (slsMap["TH"] * (1 - payoneerRatio) +
      (amazonPrice + fixedCost) / rateMap["TH"]) /
      (totalCommMap["TH"] * (1 - payoneerRatio) -
        profitRatio * (1 - voucherRatio)) /
      (1 - promoRatio)
  );
  priceMap["TW"] =
    Math.round(
      ((slsMap["TW"] * (1 - payoneerRatio) +
        (amazonPrice + fixedCost) / rateMap["TW"]) /
        (totalCommMap["TW"] * (1 - payoneerRatio) -
          profitRatio * (1 - voucherRatio)) /
        (1 - promoRatio)) *
        100
    ) / 100;
  priceMap["VN"] = Math.round(
    (slsMap["VN"] * (1 - payoneerRatio) +
      (amazonPrice + fixedCost) / rateMap["VN"]) /
      (totalCommMap["VN"] * (1 - payoneerRatio) -
        profitRatio * (1 - voucherRatio)) /
      (1 - promoRatio)
  );
  priceMap["BR"] =
    Math.round(
      ((slsMap["BR"] * (1 - payoneerRatio) +
        (amazonPrice + fixedCost) / rateMap["BR"]) /
        (totalCommMap["BR"] * (1 - payoneerRatio) -
          profitRatio * (1 - voucherRatio)) /
        (1 - promoRatio)) *
        100
    ) / 100;

  return priceMap;
}

// /**
//  * バリエーション結果をシートに保存
//  */
// function saveVariationsToSheet(sheet, variationResult) {
//   const lastRow = sheet.getLastRow();
//   const variationStartRow = lastRow + 2;

//   sheet.getRange(variationStartRow, COL_VARIATION1_NAME).setValue(variationResult.variations.variation1);
//   sheet.getRange(variationStartRow, COL_VARIATION2_NAME).setValue(variationResult.variations.variation2);

//   const combinationStartRow = variationStartRow + 1;
//   variationResult.variations.combinations.forEach((combination, i) => {
//     sheet.getRange(combinationStartRow + i, COL_VARIATION_COMBINATIONS).setValue(combination);
//   });

//   const assignmentStartRow = combinationStartRow + variationResult.variations.combinations.length + 1;
//   variationResult.assignments.forEach((assignment, i) => {
//     sheet.getRange(assignmentStartRow + i, COL_VARIATION_ASSIGNMENTS).setValue(assignment);
//   });

//   Logger.log('バリエーション情報を保存しました。');
// }

/**
 * バリエーション品の価格Mapを取得
 */
function getPriceMapVariation(sheet, row, startcol) {
  return {
    SG: parseFloat(
      sheet.getRange(row, startcol + VARIATION_PRICE_SG).getValue()
    ),
    MY: parseFloat(
      sheet.getRange(row, startcol + VARIATION_PRICE_SG + 1).getValue()
    ),
    PH: parseInt(
      sheet.getRange(row, startcol + VARIATION_PRICE_SG + 2).getValue()
    ),
    TH: parseInt(
      sheet.getRange(row, startcol + VARIATION_PRICE_SG + 3).getValue()
    ),
    TW: parseFloat(
      sheet.getRange(row, startcol + VARIATION_PRICE_SG + 4).getValue()
    ),
    VN: parseInt(
      sheet.getRange(row, startcol + VARIATION_PRICE_SG + 5).getValue()
    ),
    BR: parseInt(
      sheet.getRange(row, startcol + VARIATION_PRICE_SG + 6).getValue()
    ),
  };
}
