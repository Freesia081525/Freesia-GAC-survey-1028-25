// 全域變數，設定您的工作表名稱
const SHEET_NAME = "工作表1"; // 請確認這是您要寫入資料的工作表名稱
const GEMINI_API_KEY = "在這裡貼上您的 Gemini API 金鑰"; // 警告：直接貼上金鑰有安全風險，建議使用 PropertiesService 儲存

/**
 * 當使用者透過瀏覽器訪問 Web App URL 時，執行此函數。
 * 它會載入並顯示我們的問卷網頁 (index.html)。
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('AI 代理人工具應用於 TFDA 醫療器材查驗登記審查電子問卷')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 引入外部 CSS 檔案，讓我們的網頁更好看。
 * @param {string} filename CSS 檔案的名稱。
 * @returns {string} 包含在 <style> 標籤中的 CSS 內容。
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 接收從前端網頁提交的表單資料。
 * @param {object} formData 包含所有問卷回答的物件。
 * @returns {object} 一個包含成功訊息和 AI 分析結果的物件。
 */
function submitSurvey(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    // 如果找不到工作表，就建立一個新的
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      // 建立表頭
      const headers = [
        "提交時間", "A1_單位", "A2_使用時間", "A3_每週時數", "B4_效率同意度", "B5_介面滿意度", 
        "B6_突出功能", "C7_優化部分", "C8_系統遲緩", "D9_耗時部分", "D10_電子化優點", 
        "E11_期待AI協助", "E12_最有用AI工具", "E13_系統整合", "E14_準確度vs速度", "E15_偏好UI", 
        "E16_UI客製化", "E17_AI自動化程度", "E18_期待程度", "E19_參加試用", "F20_其他想法",
        "姓名", "Email", "AI分析_摘要", "AI分析_情緒"
      ];
      sheet.appendRow(headers);
    }
    
    // 準備要寫入工作表的資料行
    const newRow = [
      new Date(),
      formData.A1_unit.join(', '),
      formData.A2_experience,
      formData.A3_hours,
      formData.B4_efficiency,
      formData.B5_satisfaction,
      formData.B6_features.join(', '),
      formData.C7_optimization,
      formData.C8_lag,
      formData.D9_time_consuming.join(', '),
      formData.D10_advantages.join(', '),
      formData.E11_ai_help.join(', '),
      formData.E12_ai_tools.join(', '),
      formData.E13_integration,
      formData.E14_priority,
      formData.E15_ui_preference.join(', '),
      formData.E16_customization,
      formData.E17_automation,
      formData.E18_excitement,
      formData.E19_workshop,
      formData.F20_ideas,
      formData.name,
      formData.email
    ];

    // **AI 代理人功能整合**
    let summary = "N/A";
    let sentiment = "N/A";
    if (formData.F20_ideas && formData.F20_ideas.trim() !== "" && GEMINI_API_KEY !== "在這裡貼上您的 Gemini API 金鑰") {
      const aiResult = callGeminiAPI(formData.F20_ideas);
      summary = aiResult.summary;
      sentiment = aiResult.sentiment;
    }
    
    newRow.push(summary, sentiment);
    
    // 將新資料行附加到工作表
    sheet.appendRow(newRow);
    
    return { status: "success", message: "感謝您完成問卷！您的回饋已成功提交。", summary: summary, sentiment: sentiment };
  } catch (e) {
    Logger.log(e.toString());
    return { status: "error", message: "提交失敗，請稍後再試。錯誤：" + e.message };
  }
}

/**
 * 呼叫 Gemini API 進行文本分析。
 * @param {string} text 要分析的開放式問題回覆。
 * @returns {object} 包含摘要和情緒的物件。
 */
function callGeminiAPI(text) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${GEMINI_API_KEY}`;
  
  const prompt = `
    你是一個專業的問卷資料分析師。請分析以下這段來自台灣食品藥物管理署審查人員的建議，並用繁體中文回答。
    請提供兩個結果：
    1.  **摘要 (summary)**：用不超過50個字，精簡地總結這段文字的核心建議或想法。
    2.  **情緒 (sentiment)**：判斷這段文字表達的情緒，只能從「正面」、「中性」、「負面」中選擇一個。

    分析的文字如下：
    "${text}"

    請嚴格按照以下 JSON 格式回傳，不要包含任何其他說明文字或 markdown 符號：
    {
      "summary": "你的摘要內容",
      "sentiment": "正面"
    }
  `;

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const jsonResponse = JSON.parse(responseText);
    
    // 從回傳的文字中解析出 JSON 物件
    const candidateText = jsonResponse.candidates[0].content.parts[0].text;
    const result = JSON.parse(candidateText);
    
    return {
        summary: result.summary || "分析失敗",
        sentiment: result.sentiment || "分析失敗"
    };

  } catch (error) {
    Logger.log("Gemini API Error: " + error.toString());
    return {
      summary: "AI 分析時發生錯誤",
      sentiment: "錯誤"
    };
  }
}
