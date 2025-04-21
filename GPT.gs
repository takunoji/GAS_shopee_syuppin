/**
 * ChatGPTを使ってメイン商品の商品名を生成
 * @param {Object} mainData - メイン商品のデータ
 * @param {Array} variationsData - バリエーション商品のデータ配列
 * @return {Object} 各言語のメイン商品の商品名
 */
function generateMainProductNameWithChatGPT(variationsData) {
  const variationNames = variationsData.map((data) => data.title).join(", ");
  const promptJa = `
以下の #手順 を忠実に守って、 #テーマ に関する文章を200字～255字で作成してください。

#テーマ：
以下のバリエーション商品名を参考にして日本語でメイン商品の商品名を生成してください：
${variationNames}

#要件:
1. **文字列内に以下の特殊文字を含めないでください**:
   - ダブルクォーテーション ("")
   - アンパサンド (&)
   - その他のURLエンコードが必要な特殊文字

2. **検索用キーワード（最後に「日本直送」を追加）**:
   - "Brand Name + Product Type + Key Features(Materials, colors, size, modelなど)"を1つの文字列にまとめること。
   - 商品特徴や長所をできるだけ多くキーワードに付与してください。

3. **「悩み事ワード」に対する「使用したらどうなるか」を追加してください**。

#手順:
- 出力する前に、生成された文字列に特殊文字が含まれていないか確認してください。
- 何文字になったかをカウントしてください。
- カウント結果が200～255字の範囲を満たしている場合にのみ出力してください。

#文字数:
- 下限: 200字
- 上限: 255字
`;

  const promptEn = `
Please generate the following in **English** only, strictly following the instructions below:

#Theme:
Generate a descriptive **keyword string** for the main product in English (200-255 characters) based on the following variation product names:
${variationNames}

#Requirements:
1. **Do not include the following special characters in the string**:
   - Double quotes ("")
   - Ampersand (&)
   - Other characters requiring URL encoding

2. **Include "Direct from Japan" at the end**:
   - Format: Brand Name + Product Type + Key Features (Materials, colors, size, model, etc.).
   - Include details such as performance, compatibility, materials, size, portability, and model variations.

3. **Describe the result of using the product** ("what happens if you use it") and common concerns or issues the product addresses.

4. **Verify the string before output**:
   - Ensure the string does not exceed 255 characters or fall below 200 characters.
   - Confirm the string contains no prohibited special characters.
   - If the string does not meet the conditions, revise it until all requirements are satisfied.

#Output Format:
- Provide only the keyword string in English.
- Do not add any labels, descriptions, or comments.
`;
  const promptZh = `
请严格遵守以下指令，以**简体中文**创作 50-60 字的内容：

#主题：
参考以下变体产品名称，以**简体中文**生成主产品的产品名称：
${variationNames}

#要求:
1. **输出的内容中不得包含以下特殊字符**:
   - 双引号 ("")
   - 符号 (&)
   - 其他需要URL编码的特殊字符

2. 在名称末尾添加“日本直送”:
   - 包括品牌名称 + 产品类型 + 主要特点（材质、颜色、尺寸、型号等）。
   - 突出产品特点和优势，使名称更吸引人。

3. **描述使用后的效果和解决的问题**。

4. **验证输出前的内容**:
   - 确保字数在 50-60 字之间。
   - 确认不包含任何禁止的特殊字符。
   - 如果条件未满足，请重新调整内容，直至符合所有要求。

#输出格式:
- 仅输出主产品的简体中文名称，不添加其他说明或标签。
`;

  return {
    ja: callChatGPT(promptJa),
    en: callChatGPT(promptEn),
    zh: callChatGPT(promptZh),
  };
}

/**
 * ChatGPTを使ってメイン商品の商品説明を生成
 * @param {Object} mainData - メイン商品のデータ
 * @param {Array} variationsData - バリエーション商品のデータ配列
 * @return {Object} 各言語のメイン商品の商品説明
 */
function generateMainProductDescriptionWithChatGPT(variationsData) {
  const variationDescriptions = variationsData
    .map((data) => data.description)
    .join("\n");

  const promptJa = `
    以下の **手順** を忠実に守り、**日本語**で #テーマ に関する文章を 1000 ～ 2000 字で作成してください：

    #テーマ：
    以下のバリエーション商品説明を参考にして、日本語でメイン商品の商品説明を生成してください：
    ${variationDescriptions}

    1. 作成する商品説明の要件：
      - **閲覧者が購入したくなるような響く説明文**を作成してください。
      - **箇条書きスタイル**を積極的に活用し、各商品の特徴や魅力を明確に伝えてください。
      - **「使用したらどうなるか」や「悩み事ワード」**を加え、使用後の効果や解決できる課題を具体的に示してください。
      - 電化製品の場合、必ず以下の注意文を加えてください：
        - 「電化製品は変圧器、変換プラグは別途用意してください」。
      - 商品名の前には**絵文字**を使い、視覚的なインパクトを与えてください。
      - 箇条書きの各項目の先頭には必ず「・」を付けてください。

    2. **文字数条件**：
      - 作成する文章の文字数は **1000 字以上 2000 字以内** に収めてください。
      - 出力前に文字数をカウントし、指定された条件を満たしていることを確認してください。
      - もし条件を満たしていない場合は、文字を追加または削除し、条件を満たすように調整してください。

    3. **出力形式**：
      - 完成した文章を **自然な日本語** で書いてください。
      - 箇条書き形式を多用し、説明を分かりやすく、読みやすくしてください。
      - 必要に応じて改行を挿入し、構成を整えてください。

    上記手順に従い、日本語で高品質な説明文を作成してください。

  `;
  const promptEn = `
    Please generate the following in **English** only, strictly following the instructions below:

    #Theme:
    Create a **resonant product description** for the main product based on the following variation product descriptions:
    ${variationDescriptions}

    1. Write a detailed **product description** (1000-2000 characters) in English:
      - Include the phrase "Please prepare transformers and conversion plugs separately for electrical appliances."
      - Use bullet points to highlight key product features.
      - Use pictograms (emojis) before the product name.
      - Add "What happens if you use" and "Troubleshooting words" in the description.
      - Create a description that resonates with viewers and motivates them to buy.

    2. Format the description as follows:
      - Use more bullet points to emphasize the features of each product.
      - Add a dot before each bullet point in the description for readability.

    3. Before outputting, **count the number of characters**:
      - Ensure the character count is between **1000 and 2000 characters**.
      - If the count does not meet this range, adjust the content by adding or removing text until the condition is satisfied.

    Output strictly in English only.

  `;

  const promptZh = `
      请严格按照以下**程序**，用**简体中文**就#主题撰写一篇 1000-2000 字的文章：

      #主题：
      参照以下变体产品描述，用简体中文为主要产品生成产品描述：
      ${variationDescriptions}

      1. 创建产品描述的要求：
        - **创建一个能引起共鸣的**描述，让浏览者产生强烈的购买欲望。
        - **积极使用项目符号**，明确传达每个产品的特点和吸引力。
        - 添加“使用后会发生什么”和“担心的问题”**，具体说明使用后的好处以及可以解决的问题。
        - 对于电器产品，一定要加上以下警示语：
          - “请为电器单独准备变压器和转换插头”。
        - 在产品名称前使用**图形符号**，以增强视觉吸引力。
        - 每个列表项目前必须加一个“.”。

      2. **字符数要求**：
        - 文本的字符数必须在**1000 到 2000**之间。
        - 在输出之前，请计算字符数，并确认是否满足指定条件。
        - 如果字符数不满足条件，请增加或删除字符，并调整文本以满足要求。

      3. **输出格式**：
        - 用**自然流畅的简体中文**撰写完整的文本。
        - 大量使用项目符号格式，使说明清晰易读。
        - 必要时插入换行符和结构调整，以优化可读性。

      请按照上述步骤，用简体中文完成一份高质量的产品说明。
    `;

  return {
    ja: callChatGPT(promptJa),
    en: callChatGPT(promptEn),
    zh: callChatGPT(promptZh),
  };
}

/**
 * ChatGPT APIを呼び出して応答を取得
 * @param {string} prompt - プロンプト
 * @return {string} ChatGPTからの応答
 */
function callChatGPT(prompt) {
  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);

  // 基本設定シートからAPIキーとフォルダIDを取得
  const apiKey = settingsSheet.getRange("C39").getValue();
  if (!apiKey) {
    throw new Error("APIキーが設定されていません。");
  }

  const url = "https://api.openai.com/v1/chat/completions";

  const payload = {
    model: "gpt-3.5-turbo", // ChatGPT用モデル
    messages: [
      {
        role: "system",
        content: "あなたはプロフェッショナルな商品説明作成者です。",
      },
      { role: "user", content: prompt },
    ],
    max_tokens: 500,
    temperature: 0.7,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${apiKey}`,
    },
    payload: JSON.stringify(payload),
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const jsonResponse = JSON.parse(response.getContentText());

    if (jsonResponse.error) {
      throw new Error(`OpenAI APIエラー: ${jsonResponse.error.message}`);
    }

    return jsonResponse.choices[0].message.content.trim();
  } catch (e) {
    Logger.log(`エラーが発生しました: ${e.message}`);
    throw new Error(`ChatGPT APIのリクエストに失敗しました: ${e.message}`);
  }
}

/**
 * ChatGPT API呼び出し
 */
function callChatGPTVariation(prompt) {
  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);

  // 基本設定シートからAPIキーとフォルダIDを取得
  const apiKey = settingsSheet.getRange("C39").getValue();
  if (!apiKey) {
    throw new Error("APIキーが設定されていません。");
  }
  const url = "https://api.openai.com/v1/chat/completions";

  const payload = {
    model: "gpt-3.5-turbo",
    messages: [
      { role: "system", content: "あなたはプロのバリエーション作成者です。" },
      { role: "user", content: prompt },
    ],
    max_tokens: 1000,
    temperature: 0.7,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${apiKey}`,
    },
    payload: JSON.stringify(payload),
  };

  const response = UrlFetchApp.fetch(url, options);
  const jsonResponse = JSON.parse(response.getContentText());

  if (jsonResponse.error) {
    throw new Error(`ChatGPT APIエラー: ${jsonResponse.error.message}`);
  }

  return jsonResponse.choices[0].message.content.trim();
}

/**
 * ChatGPTの応答を解析
 */
function parseChatGPTResponse(response) {
  const result = {
    variations: {},
    assignments: [],
  };

  // 応答を解析してバリエーションと商品割り当てを取得
  const lines = response.split("\n");
  let currentSection = "";
  lines.forEach((line) => {
    if (line.startsWith("バリエーション1名:")) {
      currentSection = "variation1";
      result.variations.variation1 = line.split(":")[1].trim();
    } else if (line.startsWith("バリエーション2名:")) {
      currentSection = "variation2";
      result.variations.variation2 = line.split(":")[1].trim();
    } else if (line.startsWith("組み合わせ:")) {
      currentSection = "combinations";
      result.variations.combinations = line.split(":")[1].trim().split(", ");
    } else if (line.startsWith("- 商品")) {
      currentSection = "assignments";
      result.assignments.push(line.replace("- ", "").trim());
    }
  });

  return result;
}

// function saveVariationNamesToSheet() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('バリエーションASIN一覧'); // シート名を変更してください
//   const lastRow = sheet.getLastRow(); // データの最終行を取得
//   const asinColumn = 1; // ASINが入力されている列（例: A列）
//   const variationNameColumn = 2; // バリエーション名を保存する列（例: B列）

//   // ASINリストを取得
//   const asinList = sheet.getRange(2, asinColumn, lastRow - 1, 1).getValues().flat(); // A列からASINリストを取得

//   // ChatGPT APIを使用してバリエーション名を作成
//   asinList.forEach((asin, index) => {
//     if (!asin) return; // 空のASINをスキップ

//     const productName = getProductNameFromASIN(asin); // ASINから商品名を取得する関数
//     const variationName = generateVariationName(asin, productName); // 商品名からバリエーション名を生成する関数

//     // バリエーション名をシートに保存
//     sheet.getRange(index + 2, variationNameColumn).setValue(variationName);
//   });

//   Logger.log('バリエーション名が保存されました！');
// }

// // ASINから商品名を取得する関数（仮定: 外部APIやデータベースを利用）
// function getProductNameFromASIN(asin) {
//   // 実装例: 商品名を取得（外部APIにリクエストを送信するなど）
//   const productName = `Dummy Product Name for ${asin}`; // 仮の実装
//   return productName;
// }

// // ChatGPTを使用してバリエーション名を生成する関数
// function generateVariationName(asin, productName) {
//   const prompt = `ASIN: ${asin}\nProduct Name: ${productName}\nGenerate a variation name within 30 characters that highlights the unique features of this variation.`;

//   const variationName = callChatGPT(prompt); // ChatGPT API呼び出し
//   return variationName;
// }

// function generateVariationName(variationsData) {
//   const variationDescriptions = variationsData.map(data => data.description).join("\n");

//   const promptJa = `

// 以下のASINリストに基づいて、それぞれの商品名を取得し、バリエーション名を作成してください。バリエーション名は30文字以内で、各バリエーションがどのように特徴的であるかを明確に伝えるようにしてください。

// ASINリスト:
// ${ASIN_LIST}

// 出力形式:
// 1. 商品名: [ASINの商品名]
// 2. バリエーション名: [30文字以内の名前]

// バリエーション名を作成する際は、商品のサイズ、色、用途、特長などを考慮してください。

//   `;
//   const promptEn = `
//     Based on the following ASIN list, retrieve each product name and create a variation name for each. The variation name must be within 30 characters and clearly reflect the distinguishing feature of each variation.

// ASIN List:
// ${ASIN_LIST}

// Output Format:
// 1. Product Name: [Product name for the ASIN]
// 2. Variation Name: [Name within 30 characters]

// When creating variation names, consider factors such as size, color, purpose, and key features of the product.

//   `;

//   const promptZh = `

// 根据以下ASIN列表，获取每个商品的名称，并为每个商品创建一个变体名称。变体名称必须在30个字符以内，并且能够清晰地反映每个变体的独特特征。

// ASIN列表：
// ${ASIN_LIST}

// 输出格式：
// 1. 商品名称：[ASIN的商品名称]
// 2. 变体名称：[30字符以内的名称]

// 在创建变体名称时，请考虑商品的尺寸、颜色、用途和主要特征。
//     `;

//   return {
//     ja: callChatGPT(promptJa),
//     en: callChatGPT(promptEn),
//     zh: callChatGPT(promptZh)
//   };
// }

/**
 * ChatGPTを使用してバリエーションを生成
 * @param {Array} variationDataList - 各商品のASIN、商品名、商品説明を含むオブジェクトのリスト
 * @returns {Object} - バリエーション名とその組み合わせ
 */
function generateVariationsWithChatGPT(variationDataList) {
  const asinCount = variationDataList.length; // ASINの数
  const promptTemplate = `
  以下の商品リストを基にして、日本語、英語、簡体語でのバリエーション1名とバリエーション2名を提案してください。
  さらに、ASIN の総数が素数の場合は **バリエーション 2 を作成しない** ようにし、それに基づいて各商品の割り当てを生成してください。

  # 必ず以下の条件を満たしてください:
  1. **ASIN の数を受け取り、素数かどうかを判定すること**:
    - ASIN の総数 (${asinCount}) が素数である場合、バリエーション2を作成せず、バリエーション1のみを生成してください。
    - 素数でない場合、バリエーション1とバリエーション2を両方作成してください。

  2. **バリエーション1Names とバリエーション2Names の文字数を 14文字以内 にすること**:
    - 日本語、英語、簡体語のいずれの言語でも、それぞれの名前が **14文字以内** に収まるようにしてください。

  3. **バリエーション1とバリエーション2の組み合わせがASINの数と一致すること**:
    - バリエーション2が存在する場合、"バリエーション1の値の数 × バリエーション2の値の数 = ASINの総数 (${asinCount})" としてください。
    - バリエーション2が存在しない場合、"バリエーション1の値の数 = ASINの総数 (${asinCount})" としてください。

  4. **assignments の各値は30文字以内にすること**:
    - 日本語、英語、簡体語の各値がすべて 30文字以内 に収まることを保証してください。

  5. **バリエーション1とバリエーション2の値が重複しないこと**:
    - バリエーション1の値およびバリエーション2の値に重複がないことを保証してください。

  6. **異なるASINに同じバリエーション1とバリエーション2の組み合わせを割り当てないこと**:
    - 各ASINに対して、バリエーション1とバリエーション2の組み合わせはユニークである必要があります。

  7. **出力前に条件を確認し、それを記録すること**:
    - バリエーション1の数、バリエーション2の数、およびASINの数を出力してください。
    - 条件が成立している場合のみ結果を出力してください。
    - 成立しない場合は、バリエーション1またはバリエーション2の値を調整し、条件を満たす結果を生成してください。

  # 出力形式:
  以下のフォーマットで、必ず **厳密なJSON形式** で返してください。
  余計な記号やコメント（例: "json" や "説明文"）を含めないでください。

  {
    "variation1Names": {
      "ja": "14文字以内の日本語",
      "en": "14文字以内の英語",
      "zh": "14文字以内の簡体語"
    },
    "variation2Names": {
      "ja": "14文字以内の日本語（素数の場合は空文字列）",
      "en": "14文字以内の英語（素数の場合は空文字列）",
      "zh": "14文字以内の簡体語（素数の場合は空文字列）"
    },
    "variation1": [
      {
          "ja": "日本語のバリエーション1値",
          "en": "英語のバリエーション1値",
          "zh": "簡体語のバリエーション1値"
      }
    ],
    "variation2": [
      {
          "ja": "日本語のバリエーション2値",
          "en": "英語のバリエーション2値",
          "zh": "簡体語のバリエーション2値"
      }
    ],
    "assignments": [
      {
        "variation1": "日本語のバリエーション1値",
        "variation2": "日本語のバリエーション2値（素数の場合は空文字列）",
        "asin": "ASIN1"
      }
    ],
    "validation": {
      "asinCount": ${asinCount},
      "isPrime": ${isPrime(asinCount)}, // 素数判定を直接渡す
      "variation1Count": バリエーション1の値の数,
      "variation2Count": バリエーション2の値の数（素数の場合は 0）,
      "totalCombinations": バリエーション1の値の数 × (バリエーション2の値の数 || 1),
      "isValid": true または false
    }
  }

  # 商品リスト:
  {{product_list}}

  注意:
  - **ASIN の数が素数の場合は、バリエーション2を作成しないこと**を確実に保証してください。
  - **バリエーション1Names とバリエーション2Names は 14文字以内 に収めること**。
  - **バリエーション1 × バリエーション2 の総数と ASIN の数が一致していること**を確認してください。
  - 条件を満たさない場合は再計算を行い、結果を生成してください。
  - **純粋な JSON 形式** で出力してください。
  - JSON 以外の記号や装飾、説明文を含めないでください。
  `;

  // 商品リストを整形
  const productListText = variationDataList
    .map((item) => {
      return `ASIN: ${item.asin}, 商品名: ${item.title}, 商品説明: ${item.features}`;
    })
    .join("\n");

  const finalPrompt = promptTemplate.replace(
    "{{product_list}}",
    productListText
  );

  Logger.log(`finalPrompt: ${finalPrompt}`);
  const maxRetries = 3; // 最大リトライ回数
  let retryCount = 0;
  let isValid = false;
  let parsedResponse = null;

  while (retryCount < maxRetries && !isValid) {
    try {
      const chatResponse = callChatGPTVariation(finalPrompt);
      Logger.log(`chatResponse: ${chatResponse}`);
      parsedResponse = parseVariationAssignmentResponse(chatResponse);

      const variation1Count = parsedResponse.variation1Values?.length || 0;
      const variation2Count = parsedResponse.variation2Values?.length || 0;
      const totalCombinations = variation1Count * (variation2Count || 1);

      if (totalCombinations === variationDataList.length) {
        isValid = true;
      } else {
        Logger.log(
          `Validation failed. Total combinations: ${totalCombinations}, ASIN count: ${variationDataList.length}`
        );
        retryCount++;
      }
    } catch (error) {
      Logger.log(`エラーが発生しました: ${error.message}`);
      retryCount++;
    }
  }

  if (!isValid) {
    Logger.log("条件を満たす結果を生成できませんでした。");
  }

  return parsedResponse;
}

/**
 * 素数判定関数
 */
function isPrime(num) {
  if (num <= 1) return false;
  if (num <= 3) return true;
  if (num % 2 === 0 || num % 3 === 0) return false;
  for (let i = 5; i * i <= num; i += 6) {
    if (num % i === 0 || num % (i + 2) === 0) return false;
  }
  return true;
}
/**
 * ChatGPTのレスポンスを解析してバリエーションデータと商品割り当てを取得
 * @param {string} chatResponse - ChatGPTからのレスポンス
 * @returns {Object} - バリエーション名と割り当て結果
 */
function parseVariationAssignmentResponse(chatResponse) {
  try {
    const parsed = JSON.parse(chatResponse);

    // バリエーション1名とバリエーション2名を初期化
    const variation1Names = parsed.variation1Names || {
      ja: "",
      en: "",
      zh: "",
    };
    const variation2Names = parsed.variation2Names || {
      ja: "",
      en: "",
      zh: "",
    };

    // バリエーション1とバリエーション2の値を取得
    const variation1Values = (parsed.variation1 || []).map((variation) => ({
      ja: variation.ja || "",
      en: variation.en || "",
      zh: variation.zh || "",
    }));

    const variation2Values = (parsed.variation2 || []).map((variation) => ({
      ja: variation.ja || "",
      en: variation.en || "",
      zh: variation.zh || "",
    }));

    // assignments を解析
    const assignments = (parsed.assignments || []).map((assignment) => {
      return {
        asin: assignment.asin || "",
        variation1: {
          ja: assignment.variation1 || "",
          en: assignment.variation1 || "",
          zh: assignment.variation1 || "",
        },
        variation2: {
          ja: assignment.variation2 || "",
          en: assignment.variation2 || "",
          zh: assignment.variation2 || "",
        },
      };
    });

    return {
      variation1Names,
      variation2Names,
      variation1Values,
      variation2Values,
      assignments,
    };
  } catch (error) {
    Logger.log(`レスポンス解析エラー: ${error.message}`);
    Logger.log(`解析対象レスポンス: ${chatResponse}`);
    throw new Error("ChatGPTレスポンスの解析に失敗しました");
  }
}

/**
 * 割り当て結果をスプレッドシートに保存
 * @param {Object} result - バリエーション名と割り当て結果
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 保存先のスプレッドシート
 */
function saveVariationsAndAssignmentsToSheet(row, result, sheet) {
  const variation1Column = COL_VARIATION1_NAME; // バリエーション1列 (例: O列)
  const variation2Column = COL_VARIATION2_NAME; // バリエーション2列 (例: P列)
  const variationTypeColumn = COL_VARIATION_TYPE; // バリエーションタイプ列 (例: Q列)
  const assignment1StartColumn = COL_VARIATION1_VALUES_START; // 割り当て列開始 (例: R列)
  const assignment2StartColumn = COL_VARIATION2_VALUES_START; // 割り当て列開始 (例: R列)

  // バリエーション名を保存
  sheet.getRange(row, variation1Column).setValue(result.variation1Names.ja);
  sheet.getRange(row, variation1Column + 1).setValue(result.variation1Names.en);
  sheet.getRange(row, variation1Column + 2).setValue(result.variation1Names.zh);
  sheet
    .getRange(row, variation2Column)
    .setValue(result.variation2Names.ja || "N/A");
  sheet
    .getRange(row, variation2Column + 1)
    .setValue(result.variation2Names.en || "N/A");
  sheet
    .getRange(row, variation2Column + 2)
    .setValue(result.variation2Names.zh || "N/A");

  // バリエーション2が存在するかどうかを確認して結果を保存
  const variationCount = result.variation2Names.ja ? "2つ" : "1つ";
  sheet.getRange(row, variationTypeColumn).setValue(variationCount);

  // 商品の割り当てを保存
  result.variation1Values.forEach((value, index) => {
    // バリエーション1を保存
    sheet
      .getRange(row, assignment1StartColumn + index * 3)
      .setValue(value.ja || "N/A");
    sheet
      .getRange(row, assignment1StartColumn + index * 3 + 1)
      .setValue(value.en || "N/A");
    sheet
      .getRange(row, assignment1StartColumn + index * 3 + 2)
      .setValue(value.zh || "N/A");
  });
  // 商品の割り当てを保存
  result.variation2Values.forEach((value, index) => {
    // バリエーション2を保存
    sheet
      .getRange(row, assignment2StartColumn + index * 3)
      .setValue(value.ja || "N/A");
    sheet
      .getRange(row, assignment2StartColumn + index * 3 + 1)
      .setValue(value.en || "N/A");
    sheet
      .getRange(row, assignment2StartColumn + index * 3 + 2)
      .setValue(value.zh || "N/A");
  });
}

// /**
//  * バリエーションの組み合わせに商品を割り当てる
//  */
// function assignProductsToVariations(variationResults) {
//   const assignments = [];

//   variationResults.forEach(result => {
//     if (result.variationType === 'バリエーション1+2') {
//       assignments.push({
//         variation1: result.variation1,
//         variation2: result.variation2,
//         asin: result.asin,
//       });
//     } else {
//       assignments.push({
//         variation1: result.variation1,
//         asin: result.asin,
//       });
//     }
//   });

//   return assignments;
// }

// /**
//  * 割り当て結果をスプレッドシートに保存
//  */
// function saveAssignmentsToSheet(assignments, sheet) {
//   const assignmentRange = sheet.getRange(2, 7, assignments.length, 3); // G列以降
//   const assignmentData = assignments.map(assignment => [
//     assignment.variation1 || '',
//     assignment.variation2 || '',
//     assignment.asin,
//   ]);
//   assignmentRange.setValues(assignmentData);
// }
