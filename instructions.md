好的，這是一整套的解決方案，包含 Google Apps Script 程式碼、一個整合 Gemini AI 功能的代理人概念、詳細的安裝教學，以及20個綜合性問題，旨在將您先前的 Google Sheet 轉換為一個線上的問卷調查網頁。

### **第一部分：Google Apps Script 程式碼**

您需要在您的 Google Sheet 專案中建立三個檔案：`Code.gs`（主程式碼）、`index.html`（網頁結構）和 `styles.css`（網頁樣式）。

---

#### **檔案 1：Code.gs (Google Apps Script)**
*   **功能**：處理網頁請求、接收表單提交的資料，並將其寫入 Google Sheet。
*   **AI 代理人功能**：在成功提交後，此程式碼會呼叫 Gemini API，對開放式問題（Q20）的回答進行分析，自動提取「摘要」和「情緒分析」，並將結果寫入工作表的相應欄位。

```javascript
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
```

---

#### **檔案 2：index.html (HTML 網頁)**
*   **功能**：建立整個問卷的網頁介面，包含所有問題和選項，並透過 JavaScript 處理表單提交邏輯。

```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <!-- 引入 CSS 樣式 -->
  <?!= include('styles'); ?>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>AI 代理人工具應用於 TFDA 醫療器材查驗登記審查電子問卷 (草案)</h1>
      <p class="version">Ver.6 2025-10-28</p>
    </div>
    
    <div class="intro">
      <p>您好，這份問卷想了解您使用醫療器材查驗登記電子化送件系統情況，並研究如何透過 AI 代理人工具幫助您更愉快的進行醫療器材查驗登記審查：讓您的審查更快、更聰明，甚至更加有趣。我們希望透過此問卷能探求您的專業意見，以優化系統並提供您最新AI代理人輔助審查工具。</p>
      <p>本問卷共有 20 道題目，預計需要您10-15 分鐘。讓我們開始吧！🚀</p>
    </div>

    <!-- 表單開始 -->
    <form id="surveyForm">
      
      <!-- 所有問卷題目放在這裡 -->
      <!-- 範例：問題 1 -->
      <div class="question-block">
        <label>1. 您所屬單位（可複選）：</label>
        <div class="options">
          <input type="checkbox" name="A1_unit" value="TFDA（案件承辦人員）"> TFDA（案件承辦人員）<br>
          <input type="checkbox" name="A1_unit" value="TFDA（具覆核權限人員）"> TFDA（具覆核權限人員）<br>
          <input type="checkbox" name="A1_unit" value="TFDA（具決行權限人員）"> TFDA（具決行權限人員）<br>
          <input type="checkbox" name="A1_unit" value="醫藥品查驗中心（CDE）"> 醫藥品查驗中心（CDE）<br>
          <input type="checkbox" name="A1_unit" value="工研院量測中心"> 工研院量測中心<br>
          <input type="checkbox" name="A1_unit_other" value="其他"> 其他 (請說明: <input type="text" name="A1_unit_other_text">)
        </div>
      </div>
      
      <!-- 範例：問題 2 -->
      <div class="question-block">
        <label>2. 您使用醫療器材查驗登記電子化送件系統的時間：</label>
        <div class="options">
            <input type="radio" name="A2_experience" value="1年以下" required> 1年以下<br>
            <input type="radio" name="A2_experience" value="1-3年"> 1-3年<br>
            <input type="radio" name="A2_experience" value="3-5年"> 3-5年<br>
            <input type="radio" name="A2_experience" value="5年以上"> 5年以上<br>
        </div>
      </div>

      <!-- ... 此處省略其他 18 個問題的 HTML 結構，您可以按照上面的格式自行添加 ... -->
      <!-- 確保每個 input 都有一個獨特的 name 屬性 -->
      
      <!-- 問題 20 -->
      <div class="question-block">
        <label for="F20_ideas">20. 其他想法或 AI 功能的瘋狂點子？</label>
        <textarea id="F20_ideas" name="F20_ideas" rows="4"></textarea>
      </div>

      <!-- 聯絡資訊 -->
      <div class="question-block">
        <h3>聯絡資訊 (選擇性填寫)</h3>
        <label for="name">姓名：</label>
        <input type="text" id="name" name="name">
        <label for="email">聯絡Email：</label>
        <input type="email" id="email" name="email">
      </div>

      <div class="button-container">
        <button type="submit" id="submitBtn">提交問卷</button>
      </div>
    </form>
    
    <div id="loading" style="display:none;">
      <p>正在提交並進行 AI 分析，請稍候...</p>
    </div>
    
    <div id="thankYouMessage" style="display:none;">
      <h2>感謝您的參與！</h2>
      <p id="responseMessage"></p>
      <div id="aiAnalysisResult" style="display:none;">
          <h3>AI 洞察分析:</h3>
          <p><strong>摘要:</strong> <span id="aiSummary"></span></p>
          <p><strong>情緒:</strong> <span id="aiSentiment"></span></p>
      </div>
    </div>
  </div>

  <script>
    document.getElementById('surveyForm').addEventListener('submit', function(e) {
      e.preventDefault(); // 防止表單直接提交

      document.getElementById('submitBtn').disabled = true;
      document.getElementById('loading').style.display = 'block';

      // 收集表單資料
      const formData = new FormData(this);
      const data = {};
      
      // 處理複選框
      data.A1_unit = formData.getAll('A1_unit');
      if (formData.get('A1_unit_other_text')) {
          data.A1_unit.push('其他: ' + formData.get('A1_unit_other_text'));
      }
      
      // ... 您需要為每一個問題的 name 屬性在這裡添加處理邏輯 ...
      // 範例：處理單選和文字輸入
      data.A2_experience = formData.get('A2_experience');
      data.F20_ideas = formData.get('F20_ideas');
      data.name = formData.get('name');
      data.email = formData.get('email');
      
      // 處理所有複選題（假設您已命名）
      data.B6_features = formData.getAll('B6_features');
      data.D9_time_consuming = formData.getAll('D9_time_consuming');
      data.D10_advantages = formData.getAll('D10_advantages');
      data.E11_ai_help = formData.getAll('E11_ai_help');
      data.E12_ai_tools = formData.getAll('E12_ai_tools');
      data.E15_ui_preference = formData.getAll('E15_ui_preference');

      // 處理所有單選題
      ['A3_hours', 'B4_efficiency', 'B5_satisfaction', 'C7_optimization', 'C8_lag', 
       'E13_integration', 'E14_priority', 'E16_customization', 'E17_automation', 
       'E18_excitement', 'E19_workshop'].forEach(name => {
          data[name] = formData.get(name);
      });

      // 呼叫後端 Apps Script 函數
      google.script.run.withSuccessHandler(function(response) {
        document.getElementById('loading').style.display = 'none';
        document.getElementById('surveyForm').style.display = 'none';
        document.getElementById('thankYouMessage').style.display = 'block';
        document.getElementById('responseMessage').innerText = response.message;
        
        // 顯示 AI 分析結果
        if(response.summary && response.summary !== "N/A"){
            document.getElementById('aiAnalysisResult').style.display = 'block';
            document.getElementById('aiSummary').innerText = response.summary;
            document.getElementById('aiSentiment').innerText = response.sentiment;
        }

      }).withFailureHandler(function(error) {
        document.getElementById('loading').style.display = 'none';
        alert('提交失敗: ' + error.message);
        document.getElementById('submitBtn').disabled = false;
      }).submitSurvey(data);
    });
  </script>
</body>
</html>
```

---

#### **檔案 3：styles.css (CSS 樣式表)**
*   **功能**：美化網頁，提供更好的視覺體驗和排版。

```css
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  background-color: #f4f7f9;
  color: #333;
  line-height: 1.6;
  margin: 0;
  padding: 20px;
}

.container {
  max-width: 800px;
  margin: 0 auto;
  background-color: #fff;
  padding: 30px;
  border-radius: 8px;
  box-shadow: 0 4px 10px rgba(0,0,0,0.1);
}

.header {
  border-bottom: 2px solid #e0e0e0;
  padding-bottom: 15px;
  margin-bottom: 20px;
  text-align: center;
}

h1 {
  color: #1a237e;
  font-size: 24px;
}

.version {
  color: #666;
  font-size: 14px;
}

.intro {
  margin-bottom: 30px;
  font-size: 16px;
}

.question-block {
  margin-bottom: 25px;
  padding: 20px;
  border: 1px solid #ddd;
  border-radius: 5px;
  background-color: #fafafa;
}

label {
  font-weight: bold;
  display: block;
  margin-bottom: 10px;
  color: #3f51b5;
}

.options {
  margin-left: 10px;
}

input[type="text"],
input[type="email"],
textarea {
  width: 95%;
  padding: 10px;
  border: 1px solid #ccc;
  border-radius: 4px;
  font-size: 15px;
  margin-top: 5px;
}

textarea {
    resize: vertical;
}

.button-container {
  text-align: center;
  margin-top: 30px;
}

button[type="submit"] {
  background-color: #3f51b5;
  color: white;
  padding: 12px 30px;
  border: none;
  border-radius: 5px;
  cursor: pointer;
  font-size: 18px;
  font-weight: bold;
  transition: background-color 0.3s;
}

button[type="submit"]:disabled {
  background-color: #9fa8da;
  cursor: not-allowed;
}

button[type="submit"]:hover:enabled {
  background-color: #303f9f;
}

#loading, #thankYouMessage {
  text-align: center;
  padding: 40px;
  font-size: 18px;
}
#aiAnalysisResult {
    margin-top: 20px;
    padding: 15px;
    border: 1px dashed #3f51b5;
    background-color: #e8eaf6;
    border-radius: 5px;
}
```

---

### **第二部分：新手教學：如何將程式碼加入 Google Sheet**

請按照以下步驟操作，即使您是初學者也能輕鬆完成。

**步驟 1：取得 Gemini API 金鑰**
1.  前往 [Google AI for Developers](https://aistudio.google.com/app/apikey) 網站。
2.  點擊「**Create API key in new project**」。
3.  複製產生的 API 金鑰。這是一長串的亂數字元。
4.  將這串金鑰貼到 `Code.gs` 檔案中 `GEMINI_API_KEY` 的位置。

**步驟 2：打開 Google Sheet 並進入 Apps Script 編輯器**
1.  打開您先前建立的那個含有問卷欄位的 Google Sheet 檔案。
2.  點擊頂部選單的「**擴充功能 (Extensions)**」。
3.  在下拉選單中選擇「**Apps Script**」。這會在新的瀏覽器分頁中打開程式碼編輯器。

**步驟 3：貼上程式碼**
1.  **貼上 Code.gs**：
    *   在左側的檔案列表中，您會看到一個名為 `Code.gs` 的檔案。
    *   刪除裡面所有的預設程式碼。
    *   將上面提供的 **檔案 1：Code.gs** 的全部內容複製並貼上。
2.  **建立並貼上 index.html**：
    *   在左側「檔案」標題旁邊，點擊「**+**」號，選擇「**HTML**」。
    *   在跳出的視窗中，將檔案命名為 `index`，然後按 Enter。
    *   刪除新檔案中的預設 HTML 程式碼。
    *   將上面提供的 **檔案 2：index.html** 的全部內容複製並貼上。
3.  **建立並貼上 styles.css**：
    *   再次點擊「**+**」號，選擇「**HTML**」。
    *   將檔案命名為 `styles.css`。 **注意：雖然副檔名是 .css，但在這裡我們仍然選擇 HTML 類型，Apps Script 會正確處理它。**
    *   刪除新檔案中的預設程式碼。
    *   將上面提供的 **檔案 3：styles.css** 的全部內容複製並貼上。
4.  **儲存專案**：點擊編輯器上方的**磁碟片圖示** 💾 來儲存所有變更。

**步驟 4：部署為網路應用程式**
1.  在 Apps Script 編輯器的右上角，點擊藍色的「**部署 (Deploy)**」按鈕。
2.  選擇「**新增部署作業 (New deployment)**」。
3.  在「選取類型」旁邊，點擊**齒輪圖示**⚙️，然後選擇「**網頁應用程式 (Web app)**」。
4.  在「說明」欄位中，可以輸入一個描述，例如「TFDA 問卷調查 V1」。
5.  在「執行身分」欄位，選擇「**我 (您的 Google 帳戶)**」。
6.  在「誰可以存取」欄位，這是最重要的步驟，請務必選擇「**任何人 (Anyone)**」。這才能讓其他人填寫問卷。
7.  點擊「**部署**」。

**步驟 5：授權指令碼**
1.  點擊部署後，Google 會要求您授權此指令碼存取您的 Google Sheet。點擊「**授權存取 (Authorize access)**」。
2.  選擇您的 Google 帳戶。
3.  您可能會看到一個「Google 尚未驗證這個應用程式」的警告畫面。這是正常的。請點擊「**進階 (Advanced)**」。
4.  接著點擊「**前往 [您的專案名稱] (不安全) (Go to [Your Project Name] (unsafe))**」。
5.  在下一個畫面中，捲動到底部，然後點擊「**允許 (Allow)**」，授予權限。

**步驟 6：取得並分享您的問卷網址**
1.  授權成功後，您會看到一個「部署成功」的視窗，裡面有一個**網頁應用程式的網址 (URL)**。
2.  **複製這個網址**。這就是您公開的線上問卷連結！
3.  您可以將這個網址分享給所有問卷對象，他們點開後就能直接填寫。所有提交的資料都會即時、自動地寫入您的 Google Sheet 中。

---

### **第三部分：20個綜合性問題**

這些問題旨在深入探討AI在問卷調查、數據分析及改善審查流程中的潛力與挑戰。

1.  **即時洞察與趨勢分析**：除了對單一回饋進行摘要和情緒分析，我們如何利用 Gemini 進一步對所有已提交的問卷進行即時的趨勢分析？例如，自動識別出「最被頻繁提及的系統痛點」或「最受期待的AI功能」。
2.  **AI在數據清理中的角色**：Gemini 能否被訓練來自動識別並標記無效或矛盾的問卷回答？（例如，使用者選擇「從未使用系統」，但卻填寫了詳細的使用時數）。
3.  **語意化搜尋與知識庫建構**：我們能否將所有開放式問題的回答，透過 Gemini 轉換為一個語意化的知識庫？未來，開發者可以直接用自然語言提問（如：「使用者對於歷史紀錄查詢功能有什麼具體建議？」），AI 就能從問卷數據中找出相關答案。
4.  **使用者畫像自動生成**：基於受訪者的單位、資歷和回答模式，AI 能否自動生成幾種典型的「使用者畫像」(Personas)？例如，「資深TFDA審查員，重視準確性與歷史數據」、「年輕CDE審查員，期待高效率與創新UI」。
5.  **AI輔助的問卷設計優化**：在問卷發布前，能否讓 Gemini 模擬審查人員的角色來「預填寫」問卷，並根據其理解的困難度或模糊性，對問題的措辭提出改進建議？
6.  **個人化後續追蹤**：AI 能否根據特定使用者的回答，自動生成個人化的感謝信或後續追蹤問題？例如，對提出「瘋狂點子」的使用者，自動回信邀請他們參加更深入的焦點小組訪談。
7.  **情緒波動監控**：如果此問卷定期發放（例如每半年一次），AI 能否跨時間地追蹤特定主題（如「系統速度」）的情緒變化趨勢，並在偵測到負面情緒顯著增加時，自動向專案經理發出警報？
8.  **多語言與專業術語的挑戰**：在處理夾雜中英文及醫療器材專業術語的回饋時，Gemini 的分析準確度如何？我們需要提供哪些額外的上下文或進行「提示工程」(Prompt Engineering) 來優化其表現？
9.  **數據隱私與倫理邊界**：當 AI 分析使用者回饋時，如何確保個人隱私？特別是在分析可能帶有抱怨或批評的負面情緒時，如何避免數據被誤用，並保護填寫者的匿名性？
10. **從「描述性分析」到「預測性分析」**：AI 能否基於目前的問卷結果，預測哪些即將開發的 AI 功能可能會有最高的使用率或對審查效率帶來最大的提升？
11. **AI生成視覺化報告**：能否讓 AI 自動將統計結果和文字分析洞察，轉化為一份易於理解的 PowerPoint 簡報或互動式儀表板的草稿，供管理層快速決策？
12. **跨問卷關聯性分析**：如果 TFDA 還有其他針對不同系統的滿意度問卷，AI 能否整合多個問卷的數據，找出跨系統的、更根本性的問題或使用者期望？
13. **偏差識別與公平性**：AI 在分析數據時，是否存在放大某些群體聲音（例如，發言較多的資深人員）而忽略其他群體（例如，較少發言的新進人員）的風險？如何校準以確保公平性？
14. **「瘋狂點子」的可行性評分**：對於使用者提出的各種創新想法，AI 能否根據現有技術、預算限制和法規框架，進行初步的「可行性評分」和「潛在影響力評估」，幫助團隊排序開發優先級？
15. **整合工作流程的自動化**：當 AI 分析出一個高價值的建議時（例如，「希望系統能自動比對新舊版本文件的差異」），能否自動在專案管理工具（如 Jira、Trello）中建立一個新的功能需求單？
16. **資源消耗與成本效益**：對每一份問卷都進行 API 呼叫的成本是多少？與 AI 帶來的即時洞察價值相比，其成本效益如何？在何種規模下，這種即時分析最具價值？
17. **模型的選擇與微調**：為何選擇 Gemini 2.5 Flash 而非更強大的模型？在準確度、速度和成本之間，這個選擇代表了什麼樣的權衡？未來是否有可能使用 TFDA 的內部文件對模型進行微調，以提高其對專業術語的理解？
18. **提升使用者參與度的潛力**：在問卷的感謝頁面即時顯示 AI 對自己填寫內容的分析（摘要與情緒），這種互動性能否提升使用者的參與感和價值感，從而鼓勵更詳盡的回饋？
19. **失敗處理與可靠性**：如果 AI 分析失敗或回傳不合理的結果，系統應如何應對？是否有備用機制或人工審核流程來確保數據的完整性和可靠性？
20. **終極願景：從「被動收集」到「主動對話」**：未來的問卷是否能演變成一個由 AI 驅動的「對話式調查」？AI 可以根據使用者前面的回答，動態地生成追問的問題，像一場真實的訪談一樣，挖掘更深層次的需求。
