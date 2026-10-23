/**
 * ================================================================================================
 * 🗳️ राजस्थान राज्य निर्वाचन आयोग - पंचायत आम चुनाव 2026
 * कार्यालय उपखंड मजिस्ट्रेट (SDM), भिनाय (अजमेर)
 * मास्टर Google Apps Script बैकएंड व सेटिंग्स कंट्रोलर (Code.gs)
 * ================================================================================================
 */

var SHEET_SETTINGS = "Settings";
var SHEET_USERS = "Users";
var SHEET_CANDIDATES = "Candidates";

function onOpen() {
  SpreadsheetApp.getUi().createMenu("🗳️ चुनाव पोर्टल कंट्रोल")
    .addItem("🚀 1-क्लिक सम्पूर्ण शीट ऑटो-सेटअप (Master Setup)", "setupAllSheets")
    .addItem("⚙️ केवल Settings शीट रीसेट/अपडेट करें", "setupSettingsSheet")
    .addItem("🔄 चेकबॉक्स व परमिशन रीफ्रेश करें", "refreshCheckboxes")
    .addSeparator()
    .addItem("ℹ️ वेब ऐप URL व सहायता", "showAppInfo")
    .addToUi();
}

function setupAllSheets() {
  setupSettingsSheet();
  setupUsersSheet();
  setupCandidatesSheet();
  
  SpreadsheetApp.flush();
  try { SpreadsheetApp.getUi().alert(
    "✅ सम्पूर्ण सेटअप सफल!\n\n" +
    "1. 'Settings' शीट: पोर्टल नाम, मतदान तिथि, समय और मॉड्यूल कंट्रोल तैयार हैं।\n" +
    "2. 'Users' शीट: पंच, सरपंच, पंचायत समिति व जिला परिषद यूजर्स चेकबॉक्स अनुमतियों सहित तैयार हैं।\n" +
    "3. 'Candidates' शीट: प्रत्याशी विवरण (कमला देवी - वार्ड 2) सहित तैयार है।"
  ); } catch(e) {}
}

function setupSettingsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SETTINGS) || ss.insertSheet(SHEET_SETTINGS);
  sheet.clear();

  var headers = ["सेटिंग कुंजी (Key)", "मान / स्थिति (Value)", "प्रकार (Type)", "विवरण (Description)"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#0f766e")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center");

  var settingsRows = [
    ["PORTAL_NAME", "पंचायती राज आम चुनाव - 2026 सचित्र मतदाता पर्ची पोर्टल", "TEXT", "वेब पोर्टल का मुख्य हेडर शीर्षक"],
    ["BLOCK_NAME", "भिनाय (अजमेर)", "TEXT", "पंचायत समिति / उपखंड का नाम"],
    ["ELECTION_DATE", "15 अक्टूबर 2026", "TEXT", "मतदान की दिनांक (पर्ची व बैनर पर)"],
    ["ELECTION_TIME", "प्रातः 7:00 बजे से सायं 5:00 बजे तक", "TEXT", "मतदान का समय"],
    ["ENABLE_SEARCH", true, "CHECKBOX", "मतदाता खोज टैब चालू रखें (TRUE/FALSE)"],
    ["ENABLE_ALPHA", true, "CHECKBOX", "वर्णमाला सूची टैब चालू रखें (TRUE/FALSE)"],
    ["ENABLE_BULK_SLIPS", true, "CHECKBOX", "मत पर्ची डाउनलोड टैब चालू रखें (TRUE/FALSE)"],
    ["ENABLE_WITH_PIC_SLIPS", true, "CHECKBOX", "सचित्र पर्ची (With Pic) मास्टर स्विच"],
    ["ENABLE_CANDIDATE_PROFILE", true, "CHECKBOX", "प्रत्याशी प्रोफाइल टैब चालू रखें"],
    ["ALLOW_CANDIDATE_BANNER", true, "CHECKBOX", "पर्ची पर प्रत्याशी अपील बैनर दिखाएं"],
    ["DEFAULT_THERMAL_BATCH", "18", "TEXT", "थर्मल पर्ची डिफ़ॉल्ट बैच साइज़ (15, 18, 21)"],
    ["DEFAULT_SORT_ORDER", "alpha", "TEXT", "पर्ची सॉर्ट: alpha (A-Z) या serial (सरल क्र.)"],
    ["ADMIN_CONTACT", "7023293283", "TEXT", "हेल्पलाइन / तकनीकी सहायता मोबाइल नंबर"]
  ];

  sheet.getRange(2, 1, settingsRows.length, headers.length).setValues(settingsRows);
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  for (var r = 0; r < settingsRows.length; r++) {
    if (settingsRows[r][2] === "CHECKBOX") {
      sheet.getRange(r + 2, 2).setDataValidation(checkRule);
    }
  }

  sheet.setColumnWidth(1, 230);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 380);
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
}

function setupUsersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS) || ss.insertSheet(SHEET_USERS);
  sheet.clear();

  var headers = [
    "Username", "Password", "Full_Name", "Mobile", "Role",
    "Status", "Allowed_Panchayats", "Allowed_Wards",
    "Can_Search", "Can_Alpha", "Can_Bulk_Slips", "Can_With_Pic_Slips", "Can_Candidate_Profile",
    "Can_Print", "Can_Download", "Created_At"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#1e3a8a")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var nowStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
  var defaultUsers = [
    // 1. Super Admin: समस्त 30 पंचायतें, सभी वार्ड
    ["admin", "admin2026", "मुख्य व्यवस्थापक (Super Admin)", "7023293283", "SUPER_ADMIN", true, "ALL", "ALL", true, true, true, true, true, true, true, nowStr],
    // 2. पंचायत समिति सदस्य: कमला देवी (बांदनवाड़ा - वार्ड 2)
    ["kamala", "bhinai2026", "कमला देवी (पंचायत समिति सदस्य - बांदनवाड़ा)", "9876543210", "CANDIDATE", true, "बांदनवाड़ा", "2", true, true, true, true, true, true, true, nowStr],
    ["kamla", "bhinai2026", "कमला देवी (पंचायत समिति सदस्य - बांदनवाड़ा)", "9876543210", "CANDIDATE", true, "बांदनवाड़ा", "2", true, true, true, true, true, true, true, nowStr],
    // 3. सरपंच यूजर: पूरी ग्राम पंचायत, सारे वार्ड (ALL)
    ["sarpanch_kanai", "kanai2026", "सरपंच प्रत्याशी (कनईकलां)", "9829000001", "CANDIDATE", true, "कनईकलां", "ALL", true, true, true, true, true, true, true, nowStr],
    // 4. वार्ड पंच यूजर: केवल उसका अधिकृत वार्ड (वार्ड 4) - केवल वही वार्ड खुलेगा
    ["panch_w4", "panch2026", "वार्ड पंच प्रत्याशी (कनईकलां - वार्ड 4)", "9829000004", "CANDIDATE", true, "कनईकलां", "4", true, true, true, true, true, true, true, nowStr],
    // 5. जिला परिषद सदस्य: मल्टीपल पंचायतें या ज़ोन
    ["zila_parishad", "zila2026", "जिला परिषद प्रत्याशी (वार्ड 1)", "9829000005", "CANDIDATE", true, "बांदनवाड़ा, कनईकलां", "1", true, true, true, true, true, true, true, nowStr]
  ];

  sheet.getRange(2, 1, defaultUsers.length, headers.length).setValues(defaultUsers);

  var checkCols = [6, 9, 10, 11, 12, 13, 14, 15];
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  for (var i = 0; i < checkCols.length; i++) {
    sheet.getRange(2, checkCols[i], 100, 1).setDataValidation(checkRule);
  }

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 260);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 180);
  sheet.setColumnWidth(8, 120);
  sheet.setFrozenRows(1);
}

function setupCandidatesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CANDIDATES) || ss.insertSheet(SHEET_CANDIDATES);
  sheet.clear();

  var headers = [
    "Username", "Password", "Candidate_Name", "Post", "Panchayat",
    "Ward", "Symbol_Name", "Slogan", "Election_Date", "Election_Time",
    "Mobile", "Status", "Can_Print", "Can_Download"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#4338ca")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  var candData = [
    [
      "kamala", "bhinai2026", "कमला देवी", "पंचायत समिति सदस्य", "बांदनवाड़ा",
      "2", "हाथ", "क्षेत्र के सर्वांगीण विकास एवं खुशहाली हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।",
      "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "9876543210", true, true, true
    ],
    [
      "sarpanch_kanai", "kanai2026", "सरपंच प्रत्याशी", "सरपंच", "कनईकलां",
      "ALL", "उगता सूरज", "ग्राम पंचायत के सर्वांगीण विकास हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।",
      "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "9829000001", true, true, true
    ],
    [
      "panch_w4", "panch2026", "वार्ड पंच प्रत्याशी", "वार्ड पंच", "कनईकलां",
      "4", "पतंग", "वार्ड के विकास हेतु भारी मतों से विजयी बनावें।",
      "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "9829000004", true, true, true
    ]
  ];
  sheet.getRange(2, 1, candData.length, headers.length).setValues(candData);
  
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(2, 12, 50, 3).setDataValidation(checkRule);
  sheet.setFrozenRows(1);
}

function refreshCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  var sSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (sSheet) {
    var sData = sSheet.getDataRange().getValues();
    for (var r = 1; r < sData.length; r++) {
      if (sData[r][2] === "CHECKBOX") {
        sSheet.getRange(r + 1, 2).setDataValidation(checkRule);
      }
    }
  }

  var uSheet = ss.getSheetByName(SHEET_USERS);
  if (uSheet) {
    var checkCols = [6, 9, 10, 11, 12, 13, 14, 15];
    for (var c = 0; c < checkCols.length; c++) {
      uSheet.getRange(2, checkCols[c], 100, 1).setDataValidation(checkRule);
    }
  }

  try { SpreadsheetApp.getUi().alert("✅ समस्त चेकबॉक्स सफलतापूर्वक रीफ्रेश हो गए हैं!"); } catch(e) {}
}

function showAppInfo() {
  var msg = "ℹ️ वेब ऐप सेटअप निर्देश:\n\n" +
    "1. ऊपर नीले बटन 'Deploy' > 'New deployment' पर क्लिक करें।\n" +
    "2. 'Select type' में 'Web app' चुनें।\n" +
    "3. 'Execute as' में 'Me' चुनें।\n" +
    "4. 'Who has access' में 'Anyone' (सभी) चुनें।\n" +
    "5. 'Deploy' पर क्लिक करके Web App URL कॉपी कर लें।";
  SpreadsheetApp.getUi().alert(msg);
}

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || "ping";
  var callback = params.callback;
  var res = {};

  try {
    switch (action) {
      case "ping":
        res = { success: true, message: "Panchayat Election API Live!", timestamp: new Date().toISOString() };
        break;

      case "init":
      case "setup":
      case "setupAllSheets":
        setupSettingsSheet();
        setupUsersSheet();
        setupCandidatesSheet();
        res = { success: true, message: "समस्त शीट्स (Settings, Users, Candidates) सफलतापूर्वक तैयार हो चुकी हैं!" };
        break;

      case "getSettings":
        res = handleGetSettings();
        break;

      case "getUsers":
        res = handleGetUsers(params.adminUsername, params.adminPassword);
        break;

      case "getCandidates":
        res = handleGetCandidates();
        break;

      case "login":
        res = handleLogin(params.username, params.password);
        break;

      default:
        res = { success: false, error: "अमान्य action: " + action };
    }
  } catch (err) {
    res = { success: false, error: err.toString() };
  }

  var jsonOutput = JSON.stringify(res);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + jsonOutput + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonOutput)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var res = {};
  try {
    var payload = {};
    if (e.postData && e.postData.contents) {
      try { payload = JSON.parse(e.postData.contents); } catch(ex) { payload = e.parameter || {}; }
    } else {
      payload = e.parameter || {};
    }

    var action = payload.action || "";
    if (action === "saveSettings") {
      res = handleSaveSettings(payload.settings || payload);
    } else if (action === "saveCandidate") {
      res = handleSaveCandidate(payload.candidateData || payload);
    } else {
      res = { success: false, error: "अमान्य POST action: " + action };
    }
  } catch (err) {
    res = { success: false, error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleGetSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet || sheet.getDataRange().getLastRow() <= 1) {
    setupSettingsSheet();
    sheet = ss.getSheetByName(SHEET_SETTINGS);
  }
  if (!sheet) return { success: false, error: "Settings शीट नहीं मिली!" };

  var data = sheet.getDataRange().getValues();
  var settingsObj = {};

  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    var val = data[i][1];
    if (key) {
      if (val === true || String(val).toUpperCase() === "TRUE") {
        settingsObj[key] = true;
      } else if (val === false || String(val).toUpperCase() === "FALSE") {
        settingsObj[key] = false;
      } else {
        settingsObj[key] = val;
      }
    }
  }

  return {
    success: true,
    settings: settingsObj,
    timestamp: new Date().toISOString()
  };
}

function handleSaveSettings(newSettings) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sheet) return { success: false, error: "Settings शीट नहीं मिली!" };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    if (key && newSettings[key] !== undefined) {
      sheet.getRange(i + 1, 2).setValue(newSettings[key]);
    }
  }

  return { success: true, message: "सेटिंग्स सफलतापूर्वक अपडेट हुईं!" };
}

function handleGetUsers(adminUsername, adminPassword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet || sheet.getDataRange().getLastRow() <= 1) {
    setupUsersSheet();
    sheet = ss.getSheetByName(SHEET_USERS);
  }
  if (!sheet) return { success: true, users: [] };

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var users = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    var tabs = [];
    if (row[headers.indexOf("Can_Search")] === true) tabs.push("searchTab");
    if (row[headers.indexOf("Can_Alpha")] === true) tabs.push("alphaTab");
    if (row[headers.indexOf("Can_Bulk_Slips")] === true) tabs.push("bulkSlipTab");
    if (row[headers.indexOf("Can_With_Pic_Slips")] === true) tabs.push("withPicSlipTab");
    if (row[headers.indexOf("Can_Candidate_Profile")] === true) tabs.push("candidateProfileTab");

    users.push({
      username: String(row[0]),
      password: String(row[1]),
      fullName: String(row[2]),
      mobile: String(row[3]),
      role: String(row[4]),
      status: (row[5] === true || String(row[5]).toUpperCase() === "TRUE") ? "ACTIVE" : "INACTIVE",
      assigned_panchayats: String(row[6] || "ALL"),
      assigned_wards: String(row[7] || "ALL"),
      allowed_tabs: tabs,
      can_print: (row[13] === true),
      can_download: (row[14] === true)
    });
  }

  return { success: true, users: users };
}

function handleGetCandidates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CANDIDATES);
  if (!sheet || sheet.getDataRange().getLastRow() <= 1) {
    setupCandidatesSheet();
    sheet = ss.getSheetByName(SHEET_CANDIDATES);
  }
  if (!sheet) return { success: true, candidates: [] };

  var data = sheet.getDataRange().getValues();
  var candidates = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    candidates.push({
      username: String(row[0]),
      password: String(row[1]),
      candidate_name: String(row[2]),
      post: String(row[3]),
      panchayat: String(row[4]),
      ward: String(row[5]),
      symbol_name: String(row[6]),
      slogan: String(row[7]),
      election_date: String(row[8]),
      election_time: String(row[9]),
      mobile: String(row[10]),
      status: (row[11] === true || String(row[11]).toUpperCase() === "TRUE") ? "ACTIVE" : "INACTIVE",
      can_print: (row[12] === true),
      can_download: (row[13] === true),
      allowed_tabs: ["searchTab", "alphaTab", "bulkSlipTab", "withPicSlipTab", "candidateProfileTab"]
    });
  }

  return { success: true, candidates: candidates };
}

function handleLogin(username, password) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, error: "Users शीट नहीं मिली!" };

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var uCol = headers.indexOf("Username");
  var pCol = headers.indexOf("Password");
  var sCol = headers.indexOf("Status");

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[uCol]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
      if (String(row[pCol]).trim() !== String(password).trim()) {
        return { success: false, error: "गलत पासवर्ड!" };
      }
      if (row[sCol] !== true && String(row[sCol]).toUpperCase() !== "TRUE" && String(row[sCol]).toUpperCase() !== "ACTIVE") {
        return { success: false, error: "उपयोगकर्ता खाता निष्क्रिय (Inactive) है!" };
      }

      var isSuper = (String(row[headers.indexOf("Role")]).toUpperCase() === "SUPER_ADMIN");
      var allowedTabs = [];

      if (isSuper) {
        allowedTabs = ["dashboardTab", "searchTab", "alphaTab", "bulkSlipTab", "withPicSlipTab", "candidateProfileTab", "settingsTab"];
      } else {
        if (row[headers.indexOf("Can_Search")] === true) allowedTabs.push("searchTab");
        if (row[headers.indexOf("Can_Alpha")] === true) allowedTabs.push("alphaTab");
        if (row[headers.indexOf("Can_Bulk_Slips")] === true) allowedTabs.push("bulkSlipTab");
        if (row[headers.indexOf("Can_With_Pic_Slips")] === true) allowedTabs.push("withPicSlipTab");
        if (row[headers.indexOf("Can_Candidate_Profile")] === true) allowedTabs.push("candidateProfileTab");
        if (allowedTabs.length === 0) allowedTabs = ["searchTab"];
      }

      return {
        success: true,
        user: {
          username: String(row[uCol]),
          fullName: String(row[headers.indexOf("Full_Name")]),
          role: String(row[headers.indexOf("Role")]),
          allowed_panchayats: String(row[headers.indexOf("Allowed_Panchayats")]),
          allowed_wards: String(row[headers.indexOf("Allowed_Wards")]),
          allowed_tabs: allowedTabs,
          can_print: (isSuper || row[headers.indexOf("Can_Print")] === true),
          can_download: (isSuper || row[headers.indexOf("Can_Download")] === true)
        }
      };
    }
  }

  return { success: false, error: "उपयोगकर्ता नहीं मिला!" };
}

function handleSaveCandidate(cand) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CANDIDATES);
  if (!sheet) return { success: false, error: "Candidates शीट नहीं मिली!" };

  var data = sheet.getDataRange().getValues();
  var uName = String(cand.username || "").trim();
  var rowIdx = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === uName.toLowerCase()) {
      rowIdx = i + 1;
      break;
    }
  }

  var rowVals = [
    uName,
    cand.password || "cand2026",
    cand.candidate_name || "",
    cand.post || "सरपंच",
    cand.panchayat || "",
    cand.ward || "ALL",
    cand.symbol_name || "उगता सूरज",
    cand.slogan || "",
    cand.election_date || "15 अक्टूबर 2026",
    cand.election_time || "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
    cand.mobile || "",
    (cand.status !== false),
    (cand.can_print !== false),
    (cand.can_download !== false)
  ];

  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, rowVals.length).setValues([rowVals]);
  } else {
    sheet.appendRow(rowVals);
  }

  return { success: true, message: "प्रत्याशी विवरण सुरक्षित हो गया!" };
}
