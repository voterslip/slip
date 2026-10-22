/**
 * ================================================================================================
 * 🗳️ राजस्थान राज्य निर्वाचन आयोग - पंचायत आम चुनाव 2026
 * कार्यालय उपखंड मजिस्ट्रेट (SDM), भिनाय (अजमेर)
 * सचित्र मतदाता पर्ची एवं खोज पोर्टल (voterslip/slip)
 * Master Google Apps Script Web App Engine (Code.gs)
 * ================================================================================================
 * 
 * 🚀 1-क्लिक ऑटो-सेटअप:
 *   - Google Sheet में 'Extensions' > 'Apps Script' में यह पूरा कोड पेस्ट करें।
 *   - ऊपर फ़ंक्शन ड्रॉपडाउन से 'initialSetup' चुनें और 'Run' (चलाएं) पर क्लिक करें।
 *   - यह अपने आप Users, Settings और Summary शीट्स को सभी आवश्यक हेडर, चेकबॉक्स और
 *     डिफ़ॉल्ट एडमिन व प्रत्याशी खातों के साथ तैयार कर देगा।
 * 
 * 🌐 वेब ऐप डिप्लॉयमेंट:
 *   - 'Deploy' > 'New deployment' पर क्लिक करें।
 *   - 'Select type' में 'Web app' चुनें।
 *   - Execute as: 'Me' (आपका ईमेल)
 *   - Who has access: 'Anyone' (कोई भी)
 *   - 'Deploy' पर क्लिक करके Web App URL कॉपी करें और पोर्टल सेटिंग्स में पेस्ट करें।
 * ================================================================================================
 */

// Global Sheet Tab Names
var SHEET_USERS = "Users";
var SHEET_SETTINGS = "Settings";
var SHEET_SUMMARY = "Summary";

/**
 * Custom Menu on Spreadsheet Open
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🗳️ चुनाव पोर्टल मेन्यू")
    .addItem("🚀 1-क्लिक शीट ऑटो-सेटअप (Initial Setup)", "initialSetup")
    .addItem("🔄 अनुमतियाँ व चेकबॉक्स रीफ्रेश करें", "refreshCheckboxes")
    .addItem("➕ नया यूजर जोड़ें", "promptAddUser")
    .addSeparator()
    .addItem("ℹ️ डिप्लॉयमेंट व API सहायता", "showApiInfo")
    .addToUi();
}

/**
 * 1-Click Complete Sheet Setup with Real Checkboxes and Formatting
 */
function initialSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // -------------------------------------------------------------------------
  // 1. SETUP USERS SHEET
  // -------------------------------------------------------------------------
  var userSheet = ss.getSheetByName(SHEET_USERS);
  if (!userSheet) {
    userSheet = ss.insertSheet(SHEET_USERS);
  } else {
    userSheet.clear();
  }

  var headers = [
    "Username",              // 1
    "Password",              // 2
    "Full_Name",             // 3
    "Mobile",                // 4
    "Role",                  // 5
    "Status",                // 6 (Checkbox: TRUE/FALSE)
    "Allowed_Panchayats",    // 7
    "Allowed_Wards",         // 8
    "Can_Search",            // 9 (Checkbox)
    "Can_Alpha",             // 10 (Checkbox)
    "Can_Bulk_Slips",        // 11 (Checkbox)
    "Can_With_Pic_Slips",    // 12 (Checkbox: सचित्र पर्ची With Pic अनुमति)
    "Can_Directory",         // 13 (Checkbox)
    "Can_Candidate_Profile", // 14 (Checkbox)
    "Can_Print",             // 15 (Checkbox)
    "Can_Download",          // 16 (Checkbox)
    "Created_At"             // 17
  ];

  userSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Style Header Row
  var headerRange = userSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1e3a8a"); // Navy Blue
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment("center");
  userSheet.setRowHeight(1, 38);
  userSheet.setFrozenRows(1);

  // Default Users Data
  var nowStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
  var sampleUsers = [
    // Super Admin: All permissions TRUE
    [
      "admin", "admin2026", "मुख्य व्यवस्थापक (Super Admin)", "7023293283",
      "SUPER_ADMIN", true, "ALL", "ALL",
      true, true, true, true, true, true, true, true,
      nowStr
    ],
    // Candidate Sarpanch: With Pic enabled
    [
      "sarpanch", "bhinai2026", "नरेंद्र जी (सरपंच प्रत्याशी)", "9876543210",
      "CANDIDATE", true, "कनईकला", "ALL",
      true, true, true, true, true, true, true, true,
      nowStr
    ],
    // Candidate Ward Panch
    [
      "ward_panch", "ward2026", "रामप्रसाद (वार्ड पंच)", "9876543211",
      "CANDIDATE", true, "बड़गांव", "1",
      true, true, true, true, false, true, true, true,
      nowStr
    ],
    // BLO User: Limited to search & directory
    [
      "blo_user", "blo2026", "बी.एल.ओ. प्रगणक", "9876543212",
      "BLO", true, "कनईकला", "1",
      true, false, false, false, true, false, true, false,
      nowStr
    ]
  ];

  userSheet.getRange(2, 1, sampleUsers.length, headers.length).setValues(sampleUsers);

  // Apply Real Checkbox Validations on Columns: 6, 9, 10, 11, 12, 13, 14, 15, 16
  var checkboxCols = [6, 9, 10, 11, 12, 13, 14, 15, 16];
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  var maxRows = Math.max(100, sampleUsers.length + 20);
  for (var c = 0; c < checkboxCols.length; c++) {
    var colIdx = checkboxCols[c];
    userSheet.getRange(2, colIdx, maxRows - 1, 1).setDataValidation(checkRule);
  }

  // Adjust Column Widths
  userSheet.setColumnWidth(1, 130); // Username
  userSheet.setColumnWidth(2, 120); // Password
  userSheet.setColumnWidth(3, 220); // Full_Name
  userSheet.setColumnWidth(4, 120); // Mobile
  userSheet.setColumnWidth(5, 130); // Role
  userSheet.setColumnWidth(6, 80);  // Status
  userSheet.setColumnWidth(7, 160); // Allowed_Panchayats
  userSheet.setColumnWidth(8, 110); // Allowed_Wards
  userSheet.setColumnWidth(9, 95);  // Can_Search
  userSheet.setColumnWidth(10, 95); // Can_Alpha
  userSheet.setColumnWidth(11, 115);// Can_Bulk_Slips
  userSheet.setColumnWidth(12, 140);// Can_With_Pic_Slips
  userSheet.setColumnWidth(13, 105);// Can_Directory
  userSheet.setColumnWidth(14, 140);// Can_Candidate_Profile
  userSheet.setColumnWidth(15, 90); // Can_Print
  userSheet.setColumnWidth(16, 110);// Can_Download
  userSheet.setColumnWidth(17, 150);// Created_At

  // -------------------------------------------------------------------------
  // 2. SETUP SETTINGS SHEET
  // -------------------------------------------------------------------------
  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_SETTINGS);
  } else {
    settingsSheet.clear();
  }

  var settingsData = [
    ["सेटिंग कुंजी (Setting Key)", "मान (Value)", "विवरण (Description)"],
    ["PORTAL_NAME", "पंचायती राज आम चुनाव - 2026 सचित्र मतदाता पर्ची पोर्टल", "पोर्टल का मुख्य शीर्षक"],
    ["BLOCK_NAME", "भिनाय (अजमेर)", "पंचायत समिति / ब्लॉक"],
    ["ELECTION_DATE", "15 अक्टूबर 2026", "मतदान की आधिकारिक तिथि"],
    ["ELECTION_TIME", "प्रातः 7:00 बजे से सायं 5:00 बजे तक", "मतदान समय"],
    ["GITHUB_REPO", "https://github.com/voterslip/slip", "GitHub रिपॉजिटरी लिंक"],
    ["LIVE_PORTAL_URL", "https://voterslip.github.io/slip/", "GitHub Pages लाइव यूआरएल"],
    ["ALLOW_WITH_PIC_GLOBAL", "TRUE", "सचित्र पर्ची मॉड्यूल का मास्टर ऑन/ऑफ स्विच"],
    ["DEFAULT_THERMAL_BATCH", "18", "थर्मल पर्ची डिफ़ॉल्ट प्रिंट बैच (15, 18, 21)"],
    ["DEFAULT_SORT_ORDER", "alpha", "पर्ची क्रम: alpha (A-Z अंग्रेजी) अथवा serial (क्रमांक)"]
  ];

  settingsSheet.getRange(1, 1, settingsData.length, 3).setValues(settingsData);
  var setHeader = settingsSheet.getRange(1, 1, 1, 3);
  setHeader.setBackground("#0f766e"); // Teal
  setHeader.setFontColor("#ffffff");
  setHeader.setFontWeight("bold");
  settingsSheet.setColumnWidth(1, 220);
  settingsSheet.setColumnWidth(2, 350);
  settingsSheet.setColumnWidth(3, 300);
  settingsSheet.setRowHeight(1, 35);
  settingsSheet.setFrozenRows(1);

  // -------------------------------------------------------------------------
  // 3. SETUP SUMMARY SHEET (30 GRAM PANCHAYATS)
  // -------------------------------------------------------------------------
  var summarySheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SHEET_SUMMARY);
  } else {
    summarySheet.clear();
  }

  var gpList = [
    ["Gram_Panchayat_EN", "Gram_Panchayat_HI", "Total_Wards"],
    ["Badgaon", "बड़गांव", 13], ["Badli", "बड़ली", 11], ["Bagrai", "बगराई", 9],
    ["Bandanwara", "बांदनवाड़ा", 17], ["Bhinay", "भिनाय", 19], ["Boobkiya", "बूबकिया", 11],
    ["Chapaneri", "चापानेरी", 11], ["Chhachhundra", "छछून्दरा", 11], ["DEVPURA", "देवपुरा", 11],
    ["Devliyakalan", "देवलियाकलां", 13], ["Dhantol", "धांतोल", 9], ["Ekalsingha", "एकलसिंहा", 9],
    ["Ghana", "घणा", 9], ["GudhaKhurd", "गुढाखुर्द", 11], ["Hiyaliya", "हियालिया", 9],
    ["Kanaikalan", "कनईकला", 11], ["Karanti", "करांटी", 9], ["Kerot", "कैरोंट", 11],
    ["Khedi", "खेडी", 11], ["Kumhariya", "कुम्हारिया", 13], ["Lamgra", "लामगरा", 11],
    ["Nagola", "नागोला", 13], ["Nandsi", "नान्दसी", 13], ["Padanga", "पड़ांगा", 9],
    ["Padliya", "पाडलिया", 11], ["Rammaliya", "राममालिया", 9], ["Ratakot", "राताकोट", 11],
    ["Singawal", "सिंगावल", 11], ["Sobdi", "सोबडी", 9], ["Solkhurd", "सोलखुर्द", 9]
  ];

  summarySheet.getRange(1, 1, gpList.length, 3).setValues(gpList);
  var sumHeader = summarySheet.getRange(1, 1, 1, 3);
  sumHeader.setBackground("#4338ca"); // Indigo
  sumHeader.setFontColor("#ffffff");
  sumHeader.setFontWeight("bold");
  summarySheet.setColumnWidth(1, 180);
  summarySheet.setColumnWidth(2, 180);
  summarySheet.setColumnWidth(3, 120);
  summarySheet.setRowHeight(1, 35);
  summarySheet.setFrozenRows(1);

  SpreadsheetApp.flush();

  var ui = SpreadsheetApp.getUi();
  ui.alert("✅ बधाई! चुनाव पोर्टल Google Sheet ऑटो-सेटअप पूर्ण हो चुका है।\n\n- Users शीट: सभी चेकबॉक्स व कॉलम तैयार\n- Settings शीट: पोर्टल सेटिंग्स तैयार\n- Summary शीट: 30 ग्राम पंचायतें तैयार\n\nअब 'Deploy' > 'New deployment' करके Web App URL पोर्टल में जोड़ें।");
}

/**
 * Refresh checkboxes validation on Users sheet
 */
function refreshCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return;

  var lastRow = Math.max(100, sheet.getLastRow() + 20);
  var checkboxCols = [6, 9, 10, 11, 12, 13, 14, 15, 16];
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  for (var c = 0; c < checkboxCols.length; c++) {
    sheet.getRange(2, checkboxCols[c], lastRow - 1, 1).setDataValidation(checkRule);
  }

  SpreadsheetApp.getUi().alert("✅ समस्त चेकबॉक्स सफलतापूर्वक रीफ्रेश कर दिए गए हैं।");
}

/**
 * Standard HTTP GET Handler (REST API)
 */
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || "ping";
  var callback = params.callback;

  var responseData = {};

  try {
    switch (action) {
      case "ping":
        responseData = {
          success: true,
          message: "Panchayat Election 2026 Voter Slip API is Live!",
          version: "3.0.0",
          timestamp: new Date().toISOString()
        };
        break;

      case "login":
        responseData = handleLogin(params.username, params.password);
        break;

      case "getUsers":
        responseData = handleGetUsers(params.adminUsername, params.adminPassword);
        break;

      case "getSettings":
        responseData = handleGetSettings();
        break;

      case "getPanchayats":
        responseData = handleGetPanchayats();
        break;

      default:
        responseData = {
          success: false,
          error: "अमान्य Action: " + action
        };
    }
  } catch (err) {
    responseData = {
      success: false,
      error: err.toString()
    };
  }

  return createJsonResponse(responseData, callback);
}

/**
 * Standard HTTP POST Handler
 */
function doPost(e) {
  var responseData = {};

  try {
    var payload = {};
    if (e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (ex) {
        payload = e.parameter || {};
      }
    } else {
      payload = e.parameter || {};
    }

    var action = payload.action || "";

    switch (action) {
      case "updatePassword":
        responseData = handleUpdatePassword(payload.username, payload.newPassword, payload.adminUsername);
        break;

      case "addUser":
        responseData = handleAddUser(payload.userData, payload.adminUsername);
        break;

      case "toggleUserStatus":
        responseData = handleToggleUserStatus(payload.username, payload.status, payload.adminUsername);
        break;

      case "updateUserScope":
        responseData = handleUpdateUserScope(payload.username, payload.assignedPanchayats, payload.assignedWards, payload.adminUsername);
        break;

      case "updateUserPermissions":
        responseData = handleUpdateUserPermissions(payload.username, payload.permissions, payload.adminUsername);
        break;

      default:
        responseData = {
          success: false,
          error: "अमान्य POST Action: " + action
        };
    }
  } catch (err) {
    responseData = {
      success: false,
      error: err.toString()
    };
  }

  return createJsonResponse(responseData);
}

// ================================================================================================
// API HANDLER FUNCTIONS
// ================================================================================================

/**
 * Handle User Login & Permission Parsing
 */
function handleLogin(username, password) {
  if (!username || !password) {
    return { success: false, error: "यूजर आईडी एवं पासवर्ड दर्ज करना आवश्यक है!" };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    return { success: false, error: "Users शीट नहीं मिली! कृपया पहले 1-क्लिक सेटअप चलाएं।" };
  }

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { success: false, error: "Users शीट में कोई रिकॉर्ड नहीं है!" };
  }

  var headers = data[0];
  var uCol = headers.indexOf("Username");
  var pCol = headers.indexOf("Password");
  var nCol = headers.indexOf("Full_Name");
  var mCol = headers.indexOf("Mobile");
  var rCol = headers.indexOf("Role");
  var sCol = headers.indexOf("Status");
  var gpCol = headers.indexOf("Allowed_Panchayats");
  var wCol = headers.indexOf("Allowed_Wards");

  var cSearch = headers.indexOf("Can_Search");
  var cAlpha = headers.indexOf("Can_Alpha");
  var cBulk = headers.indexOf("Can_Bulk_Slips");
  var cWithPic = headers.indexOf("Can_With_Pic_Slips");
  var cDir = headers.indexOf("Can_Directory");
  var cCand = headers.indexOf("Can_Candidate_Profile");
  var cPrint = headers.indexOf("Can_Print");
  var cDown = headers.indexOf("Can_Download");

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var u = String(row[uCol] || "").trim().toLowerCase();
    var p = String(row[pCol] || "").trim();

    if (u === String(username).trim().toLowerCase()) {
      if (p !== String(password).trim()) {
        return { success: false, error: "गलत पासवर्ड! कृपया सही पासवर्ड दर्ज करें।" };
      }

      var isStatusActive = (row[sCol] === true || String(row[sCol]).toUpperCase() === "ACTIVE" || String(row[sCol]).toUpperCase() === "TRUE");
      if (!isStatusActive) {
        return { success: false, error: "यह खाता निष्क्रिय (Inactive) कर दिया गया है! मुख्य व्यवस्थापक से संपर्क करें।" };
      }

      var role = rCol !== -1 ? String(row[rCol] || "CANDIDATE").toUpperCase() : "CANDIDATE";
      var isSuper = (role === "SUPER_ADMIN");

      // Build Allowed Tabs dynamically from checkboxes
      var allowedTabs = [];
      if (isSuper) {
        allowedTabs = [
          "dashboardTab", "searchTab", "alphaTab", "bulkSlipTab", 
          "withPicSlipTab", "directoryTab", "candidateProfileTab", 
          "adminControlTab", "settingsTab"
        ];
      } else {
        if (cSearch !== -1 && row[cSearch] === true) allowedTabs.push("searchTab");
        if (cAlpha !== -1 && row[cAlpha] === true) allowedTabs.push("alphaTab");
        if (cBulk !== -1 && row[cBulk] === true) allowedTabs.push("bulkSlipTab");
        if (cWithPic !== -1 && row[cWithPic] === true) allowedTabs.push("withPicSlipTab");
        if (cDir !== -1 && row[cDir] === true) allowedTabs.push("directoryTab");
        if (cCand !== -1 && row[cCand] === true) allowedTabs.push("candidateProfileTab");

        // Fallback default
        if (allowedTabs.length === 0) allowedTabs = ["searchTab"];
      }

      return {
        success: true,
        user: {
          id: String(row[uCol]),
          username: String(row[uCol]),
          fullName: nCol !== -1 ? String(row[nCol] || row[uCol]) : String(row[uCol]),
          mobile: mCol !== -1 ? String(row[mCol] || "") : "",
          role: role,
          allowed_panchayats: gpCol !== -1 ? String(row[gpCol] || "ALL").trim() : "ALL",
          allowed_wards: wCol !== -1 ? String(row[wCol] || "ALL").trim() : "ALL",
          can_search: (isSuper || (cSearch !== -1 && row[cSearch] === true)),
          can_alpha: (isSuper || (cAlpha !== -1 && row[cAlpha] === true)),
          can_bulk_slips: (isSuper || (cBulk !== -1 && row[cBulk] === true)),
          can_with_pic_slips: (isSuper || (cWithPic !== -1 && row[cWithPic] === true)),
          can_directory: (isSuper || (cDir !== -1 && row[cDir] === true)),
          can_candidate_profile: (isSuper || (cCand !== -1 && row[cCand] === true)),
          can_print: (isSuper || (cPrint !== -1 && row[cPrint] === true)),
          can_download: (isSuper || (cDown !== -1 && row[cDown] === true)),
          allowed_tabs: allowedTabs,
          status: "ACTIVE"
        }
      };
    }
  }

  return { success: false, error: "उपयोगकर्ता (" + username + ") नहीं मिला!" };
}

/**
 * Handle Get Users for Super Admin Dashboard
 */
function handleGetUsers(adminUsername, adminPassword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return { success: true, users: [] };

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, users: [] };

  var headers = data[0];
  var uCol = headers.indexOf("Username");
  var pCol = headers.indexOf("Password");
  var nCol = headers.indexOf("Full_Name");
  var mCol = headers.indexOf("Mobile");
  var rCol = headers.indexOf("Role");
  var sCol = headers.indexOf("Status");
  var gpCol = headers.indexOf("Allowed_Panchayats");
  var wCol = headers.indexOf("Allowed_Wards");
  var cWithPic = headers.indexOf("Can_With_Pic_Slips");

  var users = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[uCol]) continue;

    users.push({
      rowIndex: i + 1,
      username: String(row[uCol]),
      password: String(row[pCol] || ""),
      fullName: nCol !== -1 ? String(row[nCol] || "") : "",
      mobile: mCol !== -1 ? String(row[mCol] || "") : "",
      role: rCol !== -1 ? String(row[rCol] || "CANDIDATE") : "CANDIDATE",
      status: (row[sCol] === true || String(row[sCol]).toUpperCase() === "ACTIVE") ? "ACTIVE" : "INACTIVE",
      allowed_panchayats: gpCol !== -1 ? String(row[gpCol] || "ALL") : "ALL",
      allowed_wards: wCol !== -1 ? String(row[wCol] || "ALL") : "ALL",
      can_with_pic_slips: (cWithPic !== -1 && row[cWithPic] === true)
    });
  }

  return { success: true, users: users };
}

/**
 * Handle Get Settings
 */
function handleGetSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
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

/**
 * Handle Get Panchayats
 */
function handleGetPanchayats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!sheet) return { success: true, panchayats: [] };

  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    list.push({
      name_en: String(data[i][0] || ""),
      name_hi: String(data[i][1] || ""),
      total_wards: Number(data[i][2] || 11)
    });
  }
  return { success: true, panchayats: list };
}

/**
 * Handle Update Password
 */
function handleUpdatePassword(username, newPassword, adminUsername) {
  if (!username || !newPassword) {
    return { success: false, error: "Username एवं नया पासवर्ड आवश्यक हैं!" };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  var data = sheet.getDataRange().getValues();
  var uCol = data[0].indexOf("Username");
  var pCol = data[0].indexOf("Password");

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uCol]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
      sheet.getRange(i + 1, pCol + 1).setValue(String(newPassword).trim());
      SpreadsheetApp.flush();
      return { success: true, message: username + " का पासवर्ड सफलतापूर्वक बदल दिया गया!" };
    }
  }
  return { success: false, error: "यूजर नहीं मिला!" };
}

/**
 * Handle Add User
 */
function handleAddUser(userData, adminUsername) {
  if (!userData || !userData.username || !userData.password) {
    return { success: false, error: "Username एवं Password आवश्यक हैं!" };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  var data = sheet.getDataRange().getValues();
  var uCol = data[0].indexOf("Username");

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uCol]).trim().toLowerCase() === String(userData.username).trim().toLowerCase()) {
      return { success: false, error: "यह Username (" + userData.username + ") पहले से मौजूद है!" };
    }
  }

  var nowStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
  var newRow = [
    String(userData.username).trim(),
    String(userData.password).trim(),
    String(userData.fullName || userData.username).trim(),
    String(userData.mobile || "").trim(),
    String(userData.role || "CANDIDATE").trim().toUpperCase(),
    true, // Status active
    String(userData.allowed_panchayats || "ALL").trim(),
    String(userData.allowed_wards || "ALL").trim(),
    true, // Can_Search
    true, // Can_Alpha
    true, // Can_Bulk_Slips
    (userData.can_with_pic_slips !== false), // Can_With_Pic_Slips
    true, // Can_Directory
    true, // Can_Candidate_Profile
    true, // Can_Print
    true, // Can_Download
    nowStr
  ];

  sheet.appendRow(newRow);
  SpreadsheetApp.flush();
  return { success: true, message: "नया यूजर सफलतापूर्वक जोड़ा गया!" };
}

/**
 * Handle Toggle User Status
 */
function handleToggleUserStatus(username, status, adminUsername) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  var data = sheet.getDataRange().getValues();
  var uCol = data[0].indexOf("Username");
  var sCol = data[0].indexOf("Status");

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uCol]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
      var newBool = (status === true || status === "ACTIVE" || status === "true");
      sheet.getRange(i + 1, sCol + 1).setValue(newBool);
      SpreadsheetApp.flush();
      return { success: true, message: username + " की स्थिति अपडेट की गई!" };
    }
  }
  return { success: false, error: "यूजर नहीं मिला!" };
}

/**
 * Helper to create JSON or JSONP Response
 */
function createJsonResponse(data, callback) {
  var outputStr = JSON.stringify(data);
  if (callback) {
    outputStr = callback + "(" + outputStr + ");";
    return ContentService.createTextOutput(outputStr).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(outputStr).setMimeType(ContentService.MimeType.JSON);
}

/**
 * UI Prompt for Adding User
 */
function promptAddUser() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt("➕ नया यूजर जोड़ें", "यूजरनेम दर्ज करें (उदा. cand_devpura):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() === ui.Button.OK) {
    var uname = resp.getResponseText().trim();
    if (!uname) return;
    var pResp = ui.prompt("पासवर्ड दर्ज करें", "पासवर्ड दर्ज करें:", ui.ButtonSet.OK_CANCEL);
    if (pResp.getSelectedButton() === ui.Button.OK) {
      var pass = pResp.getResponseText().trim();
      handleAddUser({ username: uname, password: pass }, "admin");
      ui.alert("✅ यूजर " + uname + " जोड़ दिया गया है!");
    }
  }
}

/**
 * Show API Info Dialog
 */
function showApiInfo() {
  var ui = SpreadsheetApp.getUi();
  ui.alert("ℹ️ Google Apps Script Web App Deployment:\n\n1. ऊपर 'Deploy' > 'New deployment' पर क्लिक करें।\n2. Type: Web app चुनें।\n3. Execute as: Me\n4. Who has access: Anyone\n5. URL कॉपी करके पोर्टल सेटिंग्स में पेस्ट करें।");
}
