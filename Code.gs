/**
 * ================================================================================================
 * 🗳️ राजस्थान राज्य निर्वाचन आयोग - पंचायत आम चुनाव 2026
 * कार्यालय उपखंड मजिस्ट्रेट (SDM), भिनाय (अजमेर)
 * सचित्र मतदाता पर्ची एवं खोज पोर्टल (voterslip/slip)
 * Master Google Apps Script Web App Engine (Code.gs)
 * ================================================================================================
 */

var SHEET_USERS = "Users";
var SHEET_SETTINGS = "Settings";
var SHEET_SUMMARY = "Summary";

function onOpen() {
  SpreadsheetApp.getUi().createMenu("🗳️ चुनाव पोर्टल मेन्यू")
    .addItem("🚀 1-क्लिक शीट ऑटो-सेटअप (Initial Setup)", "initialSetup")
    .addItem("🔄 चेकबॉक्स व अनुमतियाँ रीफ्रेश करें", "refreshCheckboxes")
    .addItem("➕ नया प्रत्याशी/यूजर जोड़ें", "promptAddUser")
    .addSeparator()
    .addItem("ℹ️ डिप्लॉयमेंट व API सहायता", "showApiInfo")
    .addToUi();
}

/**
 * 1-Click Complete Google Sheet Setup with Checkboxes and Default Users
 */
function initialSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Users Sheet Setup
  var userSheet = ss.getSheetByName(SHEET_USERS) || ss.insertSheet(SHEET_USERS);
  userSheet.clear();

  var headers = [
    "Username",              // 1
    "Password",              // 2
    "Full_Name",             // 3
    "Mobile",                // 4
    "Role",                  // 5
    "Status",                // 6 (Checkbox: TRUE/FALSE)
    "Allowed_Panchayats",    // 7
    "Allowed_Wards",         // 8
    "Can_Search",            // 9 (Checkbox -> मतदाता खोज)
    "Can_Alpha",             // 10 (Checkbox -> वर्णमाला सूची)
    "Can_Bulk_Slips",        // 11 (Checkbox -> मत पर्ची डाउनलोड)
    "Can_With_Pic_Slips",    // 12 (Checkbox -> सचित्र पर्ची With Pic)
    "Can_Candidate_Profile", // 13 (Checkbox -> प्रत्याशी प्रोफाइल)
    "Can_Print",             // 14 (Checkbox -> प्रिंट अधिकार)
    "Can_Download",          // 15 (Checkbox -> डाउनलोड अधिकार)
    "Created_At"             // 16
  ];

  userSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var headerRange = userSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center");
  userSheet.setRowHeight(1, 38).setFrozenRows(1);

  var nowStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
  var sampleUsers = [
    // 1. Super Admin: Full Control
    ["admin", "admin2026", "मुख्य व्यवस्थापक (Super Admin)", "7023293283", "SUPER_ADMIN", true, "ALL", "ALL", true, true, true, true, true, true, true, nowStr],
    // 2. Candidate Kamala: Sarpanch Kanaikalan (Full with-pic slip)
    ["kamala", "bhinai2026", "कमला देवी (सरपंच प्रत्याशी - कनईकला)", "9876543210", "CANDIDATE", true, "कनईकला", "ALL", true, true, true, true, true, true, true, nowStr],
    ["kamla", "bhinai2026", "कमला देवी (सरपंच प्रत्याशी - कनईकला)", "9876543210", "CANDIDATE", true, "कनईकला", "ALL", true, true, true, true, true, true, true, nowStr],
    // 3. Candidate Ward Panch
    ["ward_panch", "ward2026", "रामप्रसाद (वार्ड पंच - बड़गांव)", "9876543211", "CANDIDATE", true, "बड़गांव", "1", true, true, true, true, true, true, true, nowStr]
  ];

  userSheet.getRange(2, 1, sampleUsers.length, headers.length).setValues(sampleUsers);

  // Apply Real Checkbox Validations on Columns: 6, 9, 10, 11, 12, 13, 14, 15
  var checkboxCols = [6, 9, 10, 11, 12, 13, 14, 15];
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  for (var c = 0; c < checkboxCols.length; c++) {
    userSheet.getRange(2, checkboxCols[c], 100, 1).setDataValidation(checkRule);
  }

  // Adjust Column Widths
  userSheet.setColumnWidth(1, 130);
  userSheet.setColumnWidth(2, 120);
  userSheet.setColumnWidth(3, 230);
  userSheet.setColumnWidth(4, 120);
  userSheet.setColumnWidth(5, 130);
  userSheet.setColumnWidth(6, 80);
  userSheet.setColumnWidth(7, 160);
  userSheet.setColumnWidth(8, 110);
  userSheet.setColumnWidth(9, 100);
  userSheet.setColumnWidth(10, 100);
  userSheet.setColumnWidth(11, 120);
  userSheet.setColumnWidth(12, 145);
  userSheet.setColumnWidth(13, 150);
  userSheet.setColumnWidth(14, 90);
  userSheet.setColumnWidth(15, 110);
  userSheet.setColumnWidth(16, 150);

  // 2. Settings Sheet Setup
  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS) || ss.insertSheet(SHEET_SETTINGS);
  settingsSheet.clear();
  var settingsData = [
    ["सेटिंग कुंजी (Setting Key)", "मान (Value)", "विवरण (Description)"],
    ["PORTAL_NAME", "पंचायती राज आम चुनाव - 2026 सचित्र मतदाता पर्ची पोर्टल", "पोर्टल का शीर्षक"],
    ["BLOCK_NAME", "भिनाय (अजमेर)", "पंचायत समिति"],
    ["ELECTION_DATE", "15 अक्टूबर 2026", "मतदान तिथि"],
    ["ELECTION_TIME", "प्रातः 7:00 बजे से सायं 5:00 बजे तक", "मतदान समय"],
    ["GITHUB_REPO", "https://github.com/voterslip/slip", "GitHub रिपॉजिटरी लिंक"],
    ["ALLOW_WITH_PIC_GLOBAL", "TRUE", "सचित्र पर्ची मास्टर स्विच"],
    ["DEFAULT_THERMAL_BATCH", "18", "थर्मल पर्ची बैच (15, 18, 21)"],
    ["DEFAULT_SORT_ORDER", "alpha", "पर्ची क्रम: alpha (A-Z अंग्रेजी)"]
  ];
  settingsSheet.getRange(1, 1, settingsData.length, 3).setValues(settingsData);
  settingsSheet.getRange(1, 1, 1, 3).setBackground("#0f766e").setFontColor("#ffffff").setFontWeight("bold");

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("✅ बधाई! चुनाव पोर्टल Google Sheet ऑटो-सेटअप पूर्ण हो चुका है।\n\n- Users शीट: सभी चेकबॉक्स व कॉलम तैयार (कमला देवी प्रत्याशी सहित)\n- Settings शीट: पोर्टल सेटिंग्स तैयार\n\nअब 'Deploy' > 'New deployment' करें।");
}

function refreshCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return;
  var checkboxCols = [6, 9, 10, 11, 12, 13, 14, 15];
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  for (var c = 0; c < checkboxCols.length; c++) {
    sheet.getRange(2, checkboxCols[c], 100, 1).setDataValidation(checkRule);
  }
  SpreadsheetApp.getUi().alert("✅ चेकबॉक्स रीफ्रेश पूर्ण!");
}

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || "ping";
  var res = {};

  try {
    if (action === "ping") {
      res = { success: true, message: "Voter Slip API Live!", timestamp: new Date().toISOString() };
    } else if (action === "login") {
      res = handleLogin(params.username, params.password);
    } else if (action === "getUsers") {
      res = handleGetUsers(params.adminUsername, params.adminPassword);
    } else if (action === "getSettings") {
      res = handleGetSettings();
    } else {
      res = { success: false, error: "अमान्य action" };
    }
  } catch(err) {
    res = { success: false, error: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function handleLogin(username, password) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, error: "Users शीट नहीं मिली!" };

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var uCol = headers.indexOf("Username");
  var pCol = headers.indexOf("Password");
  var sCol = headers.indexOf("Status");
  var cWithPic = headers.indexOf("Can_With_Pic_Slips");

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[uCol]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
      if (String(row[pCol]).trim() !== String(password).trim()) {
        return { success: false, error: "गलत पासवर्ड!" };
      }
      if (row[sCol] !== true && String(row[sCol]).toUpperCase() !== "ACTIVE" && String(row[sCol]).toUpperCase() !== "TRUE") {
        return { success: false, error: "खाता निष्क्रिय (Inactive) है!" };
      }
      var isSuper = (String(row[headers.indexOf("Role")]).toUpperCase() === "SUPER_ADMIN");
      
      // Dynamic Tab Visibility from Sheet Checkboxes
      var allowedTabs = [];
      if (isSuper) {
        allowedTabs = ["dashboardTab", "searchTab", "alphaTab", "bulkSlipTab", "withPicSlipTab", "candidateProfileTab", "settingsTab"];
      } else {
        if (row[headers.indexOf("Can_Search")] === true) allowedTabs.push("searchTab");
        if (row[headers.indexOf("Can_Alpha")] === true) allowedTabs.push("alphaTab");
        if (row[headers.indexOf("Can_Bulk_Slips")] === true) allowedTabs.push("bulkSlipTab");
        if (cWithPic !== -1 && row[cWithPic] === true) allowedTabs.push("withPicSlipTab");
        if (row[headers.indexOf("Can_Candidate_Profile")] === true) allowedTabs.push("candidateProfileTab");

        if (allowedTabs.length === 0) allowedTabs = ["searchTab"];
      }

      return {
        success: true,
        user: {
          username: String(row[uCol]),
          fullName: String(row[headers.indexOf("Full_Name")] || row[uCol]),
          role: String(row[headers.indexOf("Role")]),
          allowed_panchayats: String(row[headers.indexOf("Allowed_Panchayats")] || "ALL"),
          allowed_wards: String(row[headers.indexOf("Allowed_Wards")] || "ALL"),
          can_with_pic_slips: (isSuper || (cWithPic !== -1 && row[cWithPic] === true)),
          can_print: (isSuper || row[headers.indexOf("Can_Print")] === true),
          can_download: (isSuper || row[headers.indexOf("Can_Download")] === true),
          allowed_tabs: allowedTabs
        }
      };
    }
  }
  return { success: false, error: "उपयोगकर्ता नहीं मिला!" };
}

function handleGetUsers(adminUsername, adminPassword) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { success: true, users: [] };
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var users = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    users.push({
      username: String(row[0]),
      password: String(row[1]),
      fullName: String(row[2]),
      role: String(row[4]),
      status: (row[5] === true) ? "ACTIVE" : "INACTIVE",
      allowed_panchayats: String(row[6]),
      allowed_wards: String(row[7]),
      can_with_pic_slips: (row[11] === true)
    });
  }
  return { success: true, users: users };
}

function handleGetSettings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sheet) return { success: true, settings: {} };
  var data = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][0] || "").trim();
    var v = String(data[i][1] || "").trim();
    if (k) settings[k] = v;
  }
  return { success: true, settings: settings };
}

function promptAddUser() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt("➕ नया यूजर जोड़ें", "यूजरनेम दर्ज करें (उदा. cand_kanaikalan):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() === ui.Button.OK) {
    var uname = resp.getResponseText().trim();
    if (!uname) return;
    var pResp = ui.prompt("पासवर्ड", "पासवर्ड दर्ज करें:", ui.ButtonSet.OK_CANCEL);
    if (pResp.getSelectedButton() === ui.Button.OK) {
      var pass = pResp.getResponseText().trim();
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
      var nowStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
      sheet.appendRow([uname, pass, uname, "", "CANDIDATE", true, "ALL", "ALL", true, true, true, true, true, true, true, nowStr]);
      ui.alert("✅ यूजर " + uname + " जोड़ दिया गया है!");
    }
  }
}

function showApiInfo() {
  SpreadsheetApp.getUi().alert("ℹ️ डिप्लॉयमेंट:\n\n1. ऊपर 'Deploy' > 'New deployment' चुनें।\n2. Type: Web app, Execute as: Me, Access: Anyone.\n3. URL को पोर्टल की Google Sheet सिंक सेटिंग्स में जोड़ें।");
}
