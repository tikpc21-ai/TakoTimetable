// 1. นำโค้ดนี้ไปวางใน Google Apps Script (script.google.com)
// 2. กดปุ่ม การทำให้ใช้งานได้ (Deploy) -> การทำให้ใช้งานได้รายการใหม่ (New Deployment)
// 3. เลือกประเภท "เว็บแอป" (Web App)
// 4. การเข้าถึง: เลือก "ทุกคน" (Anyone)
// 5. คัดลอก URL ของเว็บแอป (Web App URL) มาใส่ในช่อง "Google Apps Script URL" ในหน้าตั้งค่าของระบบ TakoTimetable

const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    // Save to script properties (as a simple DB)
    // Note: ScriptProperties has a quota limit. For large data, saving to Google Sheets is recommended.
    SCRIPT_PROPERTIES.setProperty('schoolData', JSON.stringify(data));
    
    return ContentService.createTextOutput(JSON.stringify({"status": "success", "message": "Data saved successfully"}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const savedData = SCRIPT_PROPERTIES.getProperty('schoolData');
    if (savedData) {
      return ContentService.createTextOutput(savedData)
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({"status": "empty", "message": "No data found"}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Enable CORS for cross-origin requests
function doOptions(e) {
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Max-Age": "86400"
  };
  return ContentService.createTextOutput("").setHeaders(headers);
}
