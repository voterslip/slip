/**
 * ================================================================================================
 * 🗳️ राजस्थान राज्य निर्वाचन आयोग - पंचायत आम चुनाव 2026
 * कार्यालय उपखंड मजिस्ट्रेट (SDM), भिनाय (अजमेर)
 * मास्टर Google Apps Script बैकएंड व मल्टी-सेलेक्ट कंट्रोलर (Code.gs)
 * ================================================================================================
 */

var SHEET_SETTINGS = "Settings";
var SHEET_USERS = "Users";
var SHEET_CANDIDATES = "Candidates";

// समस्त 30 अधिकृत ग्राम पंचायतें
var OFFICIAL_PANCHAYATS = [
  "कनईकला", "करांटी", "कैरोंट", "खेडी", "कुम्हारिया",
  "गुढाखुर्द", "घणा", "चापानेरी", "छछून्दरा", "धांतोल",
  "देवपुरा", "देवलियाकलां", "नागोला", "नान्दसी", "पड़ांगा",
  "पाडलिया", "बगराई", "बड़गांव", "बड़ली", "बांदनवाड़ा",
  "बूबकिया", "भिनाय", "राममालिया", "राताकोट", "लामगरा",
  "सिंगावल", "सोबडी", "सोलखुर्द", "हियालिया", "एकलसिंहा"
];

// अधिकृत चुनाव चिन्ह
var OFFICIAL_SYMBOLS = [
  "हाथ", "कमल", "उगता सूरज", "पतंग", "बोतल", "साइकिल",
  "गिलास", "कार", "ताला-चाबी", "बल्ला", "पंखा", "अलमारी", "सिलेंडर", "मोमबत्ती"
];

// ================================================================================================
// 1. CUSTOM MENU & SIDEBAR / DIALOG PICKER
// ================================================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu("🗳️ चुनाव पोर्टल कंट्रोल")
    .addItem("🎯 पंचायत व वार्ड चयनकर्ता (Multi-Select Dialog)", "openPanchayatWardPicker")
    .addSeparator()
    .addItem("🚀 1-क्लिक सम्पूर्ण शीट ऑटो-सेटअप (Master Setup)", "setupAllSheets")
    .addItem("⚙️ केवल Settings शीट रीसेट/अपडेट करें", "setupSettingsSheet")
    .addItem("🔄 ड्रॉपडाउन व चेकबॉक्स रीफ्रेश करें", "refreshCheckboxes")
    .addSeparator()
    .addItem("ℹ️ वेब ऐप URL व सहायता", "showAppInfo")
    .addToUi();
}

/**
 * मल्टी-सेलेक्ट पॉपअप मॉडल: जहाँ यूजर बिना टाइप किए चेकबॉक्स द्वारा पंचायत व वार्ड चुन सकता है
 */
function openPanchayatWardPicker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getName();
  
  if (sheetName !== SHEET_USERS && sheetName !== SHEET_CANDIDATES) {
    SpreadsheetApp.getUi().alert("⚠️ कृपया पहले 'Users' या 'Candidates' शीट में उस पंक्ति (Row) पर क्लिक करें जिसे आप अपडेट करना चाहते हैं!");
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("⚠️ कृपया शीर्षक (Header) के नीचे किसी प्रत्याशी या यूजर की पंक्ति चुनें!");
    return;
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  var userName = String(rowData[headers.indexOf("Candidate_Name")] || rowData[headers.indexOf("Full_Name")] || rowData[headers.indexOf("Username")] || "प्रत्याशी");
  var post = String(rowData[headers.indexOf("Post")] || "पंचायत समिति सदस्य");
  var curPrintWard = String(rowData[headers.indexOf("Print_Ward_No")] || "2");
  var curGps = String(rowData[headers.indexOf("Allowed_Panchayats")] || "");
  var curWardScope = String(rowData[headers.indexOf("Allowed_Panchayat_Wards")] || "");

  var htmlTemplate = HtmlService.createTemplate(getPickerHtmlTemplate());
  htmlTemplate.row = row;
  htmlTemplate.sheetName = sheetName;
  htmlTemplate.userName = userName;
  htmlTemplate.post = post;
  htmlTemplate.curPrintWard = curPrintWard;
  htmlTemplate.curGps = curGps;
  htmlTemplate.curWardScope = curWardScope;
  htmlTemplate.panchayatsJson = JSON.stringify(OFFICIAL_PANCHAYATS);

  var html = htmlTemplate.evaluate()
    .setWidth(720)
    .setHeight(620)
    .setTitle("🎯 मल्टीपल पंचायत व आंशिक वार्ड चयनकर्ता");

  SpreadsheetApp.getUi().showModalDialog(html, "🎯 " + userName + " (" + post + ") - पंचायत व वार्ड चयनकर्ता");
}

/**
 * मॉडल डायलॉग से सेव बटन दबाने पर सीधे सक्रिय पंक्ति में डेटा लिखने वाला फंक्शन
 */
function savePickerData(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(data.sheetName);
    if (!sheet) return { success: false, error: "शीट नहीं मिली!" };

    var row = parseInt(data.row, 10);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var pWardCol = headers.indexOf("Print_Ward_No");
    var gpsCol = headers.indexOf("Allowed_Panchayats");
    var wardsCol = headers.indexOf("Allowed_Panchayat_Wards");

    if (pWardCol !== -1) sheet.getRange(row, pWardCol + 1).setValue(data.printWardNo || "ALL");
    if (gpsCol !== -1) sheet.getRange(row, gpsCol + 1).setValue(data.allowedPanchayats || "ALL");
    if (wardsCol !== -1) sheet.getRange(row, wardsCol + 1).setValue(data.allowedPanchayatWards || "ALL");

    SpreadsheetApp.flush();
    return { success: true, message: "पंक्ति " + row + " में पंचायत व वार्ड विवरण सफलतापूर्वक सुरक्षित हो गया!" };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * ऑन-एडिट ट्रिगर: जब यूजर Allowed_Panchayats के ड्रॉपडाउन से एक से अधिक पंचायतें चुनता है,
 * तो यह बिना टाइप किए कॉमा लगाकर जोड़ता जाता है।
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  if (sheetName !== SHEET_USERS && sheetName !== SHEET_CANDIDATES) return;

  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (row <= 1) return;

  var header = sheet.getRange(1, col).getValue();
  if (header === "Allowed_Panchayats") {
    var newVal = String(e.value || "").trim();
    var oldVal = String(e.oldValue || "").trim();

    if (!newVal || newVal === "ALL" || !oldVal || oldVal === "ALL") {
      return;
    }

    var list = oldVal.split(",").map(function(s) { return s.trim(); }).filter(Boolean);
    var idx = list.indexOf(newVal);
    if (idx !== -1) {
      list.splice(idx, 1); // अगर दोबारा चुना तो हट जाएगा
    } else {
      list.push(newVal); // नया चुना तो जुड़ जाएगा
    }

    e.range.setValue(list.join(", "));
  }
}

// ================================================================================================
// 2. MASTER 1-CLICK SETUP FOR ALL SHEETS
// ================================================================================================
function setupAllSheets() {
  setupSettingsSheet();
  setupUsersSheet();
  setupCandidatesSheet();
  
  SpreadsheetApp.flush();
  try {
    SpreadsheetApp.getUi().alert(
      "✅ सम्पूर्ण मास्टर सेटअप सफल!\n\n" +
      "1. 'Settings' शीट: पोर्टल नाम, मतदान तिथि, समय और मॉड्यूल कंट्रोल तैयार हैं।\n" +
      "2. 'Users' शीट: कमला देवी (केवल बांदनवाड़ा), जिला परिषद, सरपंच व वार्ड पंच तैयार हैं।\n" +
      "3. 'Candidates' शीट: प्रत्याशी विवरण (ड्रॉपडाउन चयन व आंशिक वार्ड सहित) तैयार है।\n\n" +
      "💡 टिप: किसी भी प्रत्याशी के आगे 'Allowed_Panchayats' वाले सेल में ड्रॉपडाउन से एक-एक करके मल्टीपल पंचायतें चुन सकते हैं या ऊपर मेनू '🎯 पंचायत व वार्ड चयनकर्ता' का उपयोग कर सकते हैं!"
    );
  } catch(e) {}
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
    ["ENABLE_ADMIN_CONTROL", true, "CHECKBOX", "एडमिन कंट्रोल हब (सुपर एडमिन केवल)"],
    ["ALLOW_PUBLIC_SEARCH", true, "CHECKBOX", "बिना लॉगिन पब्लिक सर्च अनुमति"],
    ["ALLOW_SLIP_DOWNLOAD", true, "CHECKBOX", "पर्ची डाउनलोड अनुमति (मास्टर स्विच)"],
    ["ALLOW_SLIP_PRINT", true, "CHECKBOX", "पर्ची प्रिंट अनुमति (मास्टर स्विच)"],
    ["SHOW_WATERMARK", true, "CHECKBOX", "पर्ची पर वाटरमार्क दिखाएं"],
    ["DEFAULT_USER_PASSWORD", "bhinai2026", "TEXT", "नये उपयोगकर्ताओं का डिफ़ॉल्ट पासवर्ड"]
  ];

  sheet.getRange(2, 1, settingsRows.length, 4).setValues(settingsRows);

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
    "Status", "Post", "Print_Ward_No", "Allowed_Panchayats", "Allowed_Panchayat_Wards",
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
    ["admin", "admin2026", "मुख्य व्यवस्थापक (Super Admin)", "7023293283", "SUPER_ADMIN", true, "मुख्य व्यवस्थापक", "ALL", "ALL", "ALL", true, true, true, true, true, true, true, nowStr],

    // 2. पंचायत समिति सदस्य: कमला देवी (केवल बांदनवाड़ा पंचायत) - PS वार्ड 2
    ["kamala", "bhinai2026", "कमला देवी", "9876543210", "CANDIDATE", true, "पंचायत समिति सदस्य", "2", "बांदनवाड़ा", "बांदनवाड़ा: ALL", true, true, true, true, true, true, true, nowStr],
    ["kamla", "bhinai2026", "कमला देवी", "9876543210", "CANDIDATE", true, "पंचायत समिति सदस्य", "2", "बांदनवाड़ा", "बांदनवाड़ा: ALL", true, true, true, true, true, true, true, nowStr],

    // 3. जिला परिषद सदस्य: राजेन्द्र चौधरी (ZP वार्ड 15) - 3 पंचायतें
    ["zp_rajendra", "zp2026", "राजेन्द्र चौधरी", "9829000015", "CANDIDATE", true, "जिला परिषद सदस्य", "15", "बांदनवाड़ा, देवलियाकलां, कनईकला", "बांदनवाड़ा: 1,2,3 | देवलियाकलां: 1,2,3,4 | कनईकला: 11,12,13", true, true, true, true, true, true, true, nowStr],

    // 4. सरपंच प्रत्याशी: रामेश्वर जाट (कनईकलां) - पूरी ग्राम पंचायत, सारे वार्ड खुलेंगे (ALL)
    ["sarpanch_kanai", "kanai2026", "रामेश्वर जाट", "9829000001", "CANDIDATE", true, "सरपंच", "ALL", "कनईकला", "ALL", true, true, true, true, true, true, true, nowStr],

    // 5. वार्ड पंच प्रत्याशी: सुभाष गुर्जर (कनईकलां - वार्ड 4) - केवल वार्ड 4 खुलेगा और लॉक रहेगा
    ["panch_w4", "panch2026", "सुभाष गुर्जर", "9829000004", "CANDIDATE", true, "वार्ड पंच", "4", "कनईकला", "4", true, true, true, true, true, true, true, nowStr]
  ];

  sheet.getRange(2, 1, defaultUsers.length, headers.length).setValues(defaultUsers);

  // 1. Checkboxes
  var checkCols = [6, 11, 12, 13, 14, 15, 16, 17];
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  for (var i = 0; i < checkCols.length; i++) {
    sheet.getRange(2, checkCols[i], 100, 1).setDataValidation(checkRule);
  }

  // 2. Dropdowns for Role, Post, Allowed_Panchayats
  applySheetDropdowns(sheet, headers);

  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 70);
  sheet.setColumnWidth(7, 160);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 240);
  sheet.setColumnWidth(10, 340);
  sheet.setFrozenRows(1);
}

function setupCandidatesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CANDIDATES) || ss.insertSheet(SHEET_CANDIDATES);
  sheet.clear();

  var headers = [
    "Username", "Password", "Candidate_Name", "Post", "Print_Ward_No",
    "Allowed_Panchayats", "Allowed_Panchayat_Wards", "Symbol_Name", "Slogan",
    "Election_Date", "Election_Time", "Mobile", "Status", "Can_Print", "Can_Download"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#4338ca")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var candData = [
    // 1. पंचायत समिति सदस्य प्रत्याशी: कमला देवी (केवल बांदनवाड़ा पंचायत, वार्ड 2 छपेगा)
    [
      "kamala", "bhinai2026", "कमला देवी", "पंचायत समिति सदस्य", "2",
      "बांदनवाड़ा", "बांदनवाड़ा: ALL", "हाथ",
      "क्षेत्र के सर्वांगीण विकास एवं खुशहाली हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।",
      "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "9876543210", true, true, true
    ],
    // 2. जिला परिषद सदस्य प्रत्याशी: राजेन्द्र चौधरी (वार्ड 15 छपेगा, 3 पंचायतों के आंशिक वार्ड खुलेंगे)
    [
      "zp_rajendra", "zp2026", "राजेन्द्र चौधरी", "जिला परिषद सदस्य", "15",
      "बांदनवाड़ा, देवलियाकलां, कनईकला", "बांदनवाड़ा: 1,2,3 | देवलियाकलां: 1,2,3,4 | कनईकला: 11,12,13", "कमल",
      "जिला परिषद क्षेत्र के चहुंमुखी विकास हेतु आपका अपना कर्मठ एवं विश्वस्त साथी।",
      "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "9829000015", true, true, true
    ],
    // 3. सरपंच प्रत्याशी: रामेश्वर जाट (पूरी ग्राम पंचायत कनईकलां, सारे वार्ड खुलेंगे, पर्ची पर केवल ग्राम पंचायत छपेगी)
    [
      "sarpanch_kanai", "kanai2026", "रामेश्वर जाट", "सरपंच", "ALL",
      "कनईकला", "ALL", "उगता सूरज",
      "ग्राम पंचायत के सर्वांगीण विकास हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।",
      "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "9829000001", true, true, true
    ],
    // 4. वार्ड पंच प्रत्याशी: सुभाष गुर्जर (कनईकलां - केवल वार्ड 4 खुलेगा और पर्ची पर वार्ड 4 छपेगा)
    [
      "panch_w4", "panch2026", "सुभाष गुर्जर", "वार्ड पंच", "4",
      "कनईकला", "4", "पतंग",
      "वार्ड के विकास एवं स्वच्छता हेतु भारी मतों से विजयी बनावें।",
      "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "9829000004", true, true, true
    ]
  ];

  sheet.getRange(2, 1, candData.length, headers.length).setValues(candData);
  
  // Checkboxes
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(2, 13, 50, 3).setDataValidation(checkRule);

  // Dropdowns
  applySheetDropdowns(sheet, headers);

  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 240);
  sheet.setColumnWidth(7, 340);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 280);
  sheet.setFrozenRows(1);
}

/**
 * पद, पंचायत और चुनाव चिन्ह के लिए ऑटोमैटिक ड्रॉपडाउन नियम लागू करना
 */
function applySheetDropdowns(sheet, headers) {
  var postCol = headers.indexOf("Post");
  if (postCol !== -1) {
    var postList = ["पंचायत समिति सदस्य", "जिला परिषद सदस्य", "सरपंच", "वार्ड पंच", "मुख्य व्यवस्थापक", "बी.एल.ओ."];
    var postRule = SpreadsheetApp.newDataValidation().requireValueInList(postList, true).build();
    sheet.getRange(2, postCol + 1, 100, 1).setDataValidation(postRule);
  }

  var gpCol = headers.indexOf("Allowed_Panchayats");
  if (gpCol !== -1) {
    var gpOptions = ["ALL"].concat(OFFICIAL_PANCHAYATS);
    var gpRule = SpreadsheetApp.newDataValidation().requireValueInList(gpOptions, true).build();
    sheet.getRange(2, gpCol + 1, 100, 1).setDataValidation(gpRule);
  }

  var symCol = headers.indexOf("Symbol_Name");
  if (symCol !== -1) {
    var symRule = SpreadsheetApp.newDataValidation().requireValueInList(OFFICIAL_SYMBOLS, true).build();
    sheet.getRange(2, symCol + 1, 100, 1).setDataValidation(symRule);
  }
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
    var uHeaders = uSheet.getRange(1, 1, 1, uSheet.getLastColumn()).getValues()[0];
    var checkCols = [6, 11, 12, 13, 14, 15, 16, 17];
    for (var c = 0; c < checkCols.length; c++) {
      uSheet.getRange(2, checkCols[c], 100, 1).setDataValidation(checkRule);
    }
    applySheetDropdowns(uSheet, uHeaders);
  }

  var cSheet = ss.getSheetByName(SHEET_CANDIDATES);
  if (cSheet) {
    var cHeaders = cSheet.getRange(1, 1, 1, cSheet.getLastColumn()).getValues()[0];
    cSheet.getRange(2, 13, 50, 3).setDataValidation(checkRule);
    applySheetDropdowns(cSheet, cHeaders);
  }

  try {
    SpreadsheetApp.getUi().alert("✅ समस्त चेकबॉक्स एवं ड्रॉपडाउन सूचियां सफलतापूर्वक रीफ्रेश हो गई हैं!");
  } catch(e) {}
}

function showAppInfo() {
  var msg = "ℹ️ वेब ऐप सेटअप निर्देश:\n\n" +
    "1. ऊपर नीले बटन 'Deploy' > 'New deployment' पर क्लिक करें।\n" +
    "2. 'Select type' में 'Web app' चुनें।\n" +
    "3. 'Execute as' में 'Me' चुनें।\n" +
    "4. 'Who has access' में 'Anyone' (सभी) चुनें।\n" +
    "5. 'Deploy' पर क्लिक करके Web App URL कॉपी कर लें।\n\n" +
    "💡 टिप: बिना टाइपिंग के पंचायत व वार्ड चुनने हेतु '🗳️ चुनाव पोर्टल कंट्रोल' मेनू में '🎯 पंचायत व वार्ड चयनकर्ता' का उपयोग करें!";
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch(e) {}
}

// ================================================================================================
// 3. HTTP GET & POST HANDLERS (PORTAL REAL-TIME SYNC)
// ================================================================================================
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
      post: String(row[headers.indexOf("Post")] || ""),
      print_ward_no: String(row[headers.indexOf("Print_Ward_No")] || ""),
      allowed_panchayats: String(row[headers.indexOf("Allowed_Panchayats")] || "ALL"),
      allowed_panchayat_wards: String(row[headers.indexOf("Allowed_Panchayat_Wards")] || "ALL"),
      assigned_panchayats: String(row[headers.indexOf("Allowed_Panchayats")] || "ALL"),
      assigned_wards: String(row[headers.indexOf("Allowed_Panchayat_Wards")] || "ALL"),
      allowed_wards: String(row[headers.indexOf("Allowed_Panchayat_Wards")] || "ALL"),
      ward: String(row[headers.indexOf("Print_Ward_No")] || ""),
      allowed_tabs: tabs,
      can_print: (row[headers.indexOf("Can_Print")] === true),
      can_download: (row[headers.indexOf("Can_Download")] === true)
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
  var headers = data[0];
  var candidates = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    candidates.push({
      username: String(row[0]),
      password: String(row[1]),
      candidate_name: String(row[headers.indexOf("Candidate_Name")]),
      post: String(row[headers.indexOf("Post")]),
      print_ward_no: String(row[headers.indexOf("Print_Ward_No")]),
      ward: String(row[headers.indexOf("Print_Ward_No")]),
      allowed_panchayats: String(row[headers.indexOf("Allowed_Panchayats")]),
      panchayat: String(row[headers.indexOf("Allowed_Panchayats")]),
      allowed_panchayat_wards: String(row[headers.indexOf("Allowed_Panchayat_Wards")]),
      allowed_wards: String(row[headers.indexOf("Allowed_Panchayat_Wards")]),
      symbol_name: String(row[headers.indexOf("Symbol_Name")]),
      slogan: String(row[headers.indexOf("Slogan")]),
      election_date: String(row[headers.indexOf("Election_Date")]),
      election_time: String(row[headers.indexOf("Election_Time")]),
      mobile: String(row[headers.indexOf("Mobile")]),
      status: (row[headers.indexOf("Status")] === true || String(row[headers.indexOf("Status")]).toUpperCase() === "TRUE") ? "ACTIVE" : "INACTIVE",
      can_print: (row[headers.indexOf("Can_Print")] === true),
      can_download: (row[headers.indexOf("Can_Download")] === true),
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
          post: String(row[headers.indexOf("Post")] || ""),
          print_ward_no: String(row[headers.indexOf("Print_Ward_No")] || ""),
          ward: String(row[headers.indexOf("Print_Ward_No")] || ""),
          allowed_panchayats: String(row[headers.indexOf("Allowed_Panchayats")]),
          allowed_panchayat_wards: String(row[headers.indexOf("Allowed_Panchayat_Wards")]),
          allowed_wards: String(row[headers.indexOf("Allowed_Panchayat_Wards")]),
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
  var uCol = 0;
  var targetRow = -1;

  var uName = cand.username || cand.user_id;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uCol]).toLowerCase() === String(uName).toLowerCase()) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    targetRow = data.length + 1;
    sheet.getRange(targetRow, 1).setValue(uName);
  }

  var headers = data[0];
  var fieldMap = {
    "Candidate_Name": cand.candidate_name,
    "Post": cand.post,
    "Print_Ward_No": cand.print_ward_no || cand.ward,
    "Allowed_Panchayats": cand.allowed_panchayats || cand.panchayat,
    "Allowed_Panchayat_Wards": cand.allowed_panchayat_wards || cand.allowed_wards,
    "Symbol_Name": cand.symbol_name,
    "Slogan": cand.slogan,
    "Election_Date": cand.election_date,
    "Election_Time": cand.election_time,
    "Mobile": cand.mobile
  };

  for (var key in fieldMap) {
    var cIdx = headers.indexOf(key);
    if (cIdx !== -1 && fieldMap[key] !== undefined) {
      sheet.getRange(targetRow, cIdx + 1).setValue(fieldMap[key]);
    }
  }

  return { success: true, message: "प्रत्याशी विवरण सुरक्षित हो गया!" };
}

// ================================================================================================
// 4. EMBEDDED HTML TEMPLATE FOR MULTI-SELECT PICKER DIALOG
// ================================================================================================
function getPickerHtmlTemplate() {
  return '<!DOCTYPE html>' +
  '<html>' +
  '<head>' +
  '<base target="_top">' +
  '<style>' +
  '  * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; }' +
  '  body { margin: 0; padding: 16px; background: #f8fafc; color: #1e293b; font-size: 13px; }' +
  '  .header-box { background: #4338ca; color: #fff; padding: 12px 16px; border-radius: 8px; margin-bottom: 12px; }' +
  '  .header-title { font-size: 15px; font-weight: 700; margin: 0; }' +
  '  .header-sub { font-size: 12px; opacity: 0.9; margin-top: 3px; }' +
  '  .top-controls { display: flex; gap: 10px; margin-bottom: 12px; align-items: center; background: #fff; padding: 10px; border-radius: 6px; border: 1px solid #e2e8f0; }' +
  '  .search-input { flex: 1; padding: 7px 10px; border: 1px solid #cbd5e1; border-radius: 5px; font-size: 13px; outline: none; }' +
  '  .search-input:focus { border-color: #4338ca; }' +
  '  .gp-container { max-height: 340px; overflow-y: auto; background: #fff; border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px; }' +
  '  .gp-card { border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px 10px; margin-bottom: 6px; background: #fafafa; transition: all 0.15s; }' +
  '  .gp-card.active { border-color: #6366f1; background: #eef2ff; }' +
  '  .gp-head { display: flex; align-items: center; justify-content: space-between; cursor: pointer; }' +
  '  .gp-label { font-weight: 700; font-size: 13px; color: #1e1b4b; display: flex; align-items: center; gap: 6px; cursor: pointer; }' +
  '  .ward-box { margin-top: 6px; padding-top: 6px; border-top: 1px dashed #cbd5e1; display: none; }' +
  '  .gp-card.active .ward-box { display: block; }' +
  '  .ward-chips { display: flex; flex-wrap: wrap; gap: 4px; align-items: center; }' +
  '  .ward-btn { font-size: 11px; padding: 2px 7px; border: 1px solid #cbd5e1; border-radius: 4px; background: #fff; cursor: pointer; user-select: none; }' +
  '  .ward-btn.selected { background: #4338ca; color: #fff; border-color: #4338ca; font-weight: 700; }' +
  '  .ward-btn.all-btn { background: #059669; color: #fff; border-color: #059669; font-weight: 700; }' +
  '  .footer-box { margin-top: 12px; display: flex; justify-content: space-between; align-items: center; padding-top: 10px; border-top: 1px solid #e2e8f0; }' +
  '  .save-btn { background: #059669; color: #fff; border: none; padding: 9px 20px; font-size: 13px; font-weight: 700; border-radius: 6px; cursor: pointer; }' +
  '  .save-btn:hover { background: #047857; }' +
  '  .cancel-btn { background: #e2e8f0; color: #475569; border: none; padding: 9px 16px; border-radius: 6px; cursor: pointer; font-weight: 600; }' +
  '  .preview-bar { background: #f1f5f9; padding: 6px 10px; border-radius: 5px; font-size: 11px; color: #334155; margin-top: 8px; border: 1px solid #cbd5e1; word-break: break-all; }' +
  '</style>' +
  '</head>' +
  '<body>' +
  '  <div class="header-box">' +
  '    <div class="header-title">🎯 <?= userName ?> • <?= post ?> (पंक्ति <?= row ?>)</div>' +
  '    <div class="header-sub">नीचे से पंचायतें व उनके अधिकृत वार्ड चुनें। टाइपिंग की कोई आवश्यकता नहीं है!</div>' +
  '  </div>' +
  '  <div class="top-controls">' +
  '    <input type="text" id="searchInput" class="search-input" placeholder="🔍 पंचायत खोजें..." oninput="filterGps(this.value)">' +
  '    <div style="display:flex; align-items:center; gap:6px;">' +
  '      <label for="printWardInput" style="font-weight:700; font-size:12px;">पर्ची पर छपने वाला वार्ड:</label>' +
  '      <input type="text" id="printWardInput" value="<?= curPrintWard ?>" style="width:70px; padding:6px 8px; border:1px solid #cbd5e1; border-radius:4px; font-weight:700; text-align:center;">' +
  '    </div>' +
  '  </div>' +
  '  <div class="gp-container" id="gpList"></div>' +
  '  <div class="preview-bar" id="previewText">लोड हो रहा है...</div>' +
  '  <div class="footer-box">' +
  '    <button type="button" class="cancel-btn" onclick="google.script.host.close()">रद्द करें</button>' +
  '    <button type="button" class="save-btn" id="saveBtn" onclick="doSave()">💾 शीट में सुरक्षित करें</button>' +
  '  </div>' +
  '  <script>' +
  '    var allGps = JSON.parse(<?= panchayatsJson ?>);' +
  '    var curGpsStr = "<?= curGps ?>";' +
  '    var curWardsStr = "<?= curWardScope ?>";' +
  '    var rowNum = "<?= row ?>";' +
  '    var sheetName = "<?= sheetName ?>";' +
  '    var selectedMap = {};' +
  '    parseInitial();' +
  '    renderList();' +
  '    updatePreview();' +
  '    function parseInitial() {' +
  '      if (curWardsStr === "ALL") {' +
  '        var gpArr = curGpsStr.split(",").map(function(s){return s.trim();}).filter(Boolean);' +
  '        gpArr.forEach(function(g){ selectedMap[g] = "ALL"; });' +
  '        return;' +
  '      }' +
  '      if (curWardsStr.includes(":")) {' +
  '        var parts = curWardsStr.split("|");' +
  '        parts.forEach(function(p){' +
  '          var sp = p.split(":");' +
  '          if (sp.length === 2) {' +
  '            var g = sp[0].trim();' +
  '            var w = sp[1].trim();' +
  '            if (w === "ALL") selectedMap[g] = "ALL";' +
  '            else selectedMap[g] = w.split(",").map(function(x){return x.trim();}).filter(Boolean);' +
  '          }' +
  '        });' +
  '      } else if (curGpsStr) {' +
  '        var gpArr = curGpsStr.split(",").map(function(s){return s.trim();}).filter(Boolean);' +
  '        gpArr.forEach(function(g){ selectedMap[g] = "ALL"; });' +
  '      }' +
  '    }' +
  '    function renderList(filter) {' +
  '      var container = document.getElementById("gpList");' +
  '      container.innerHTML = "";' +
  '      var fClean = (filter || "").toLowerCase().trim();' +
  '      allGps.forEach(function(gp){' +
  '        if (fClean && gp.toLowerCase().indexOf(fClean) === -1) return;' +
  '        var isSel = (selectedMap[gp] !== undefined);' +
  '        var curVal = selectedMap[gp];' +
  '        var card = document.createElement("div");' +
  '        card.className = "gp-card " + (isSel ? "active" : "");' +
  '        card.id = "card_" + gp;' +
  '        var head = document.createElement("div");' +
  '        head.className = "gp-head";' +
  '        head.innerHTML = "<label class=\'gp-label\'><input type=\'checkbox\' " + (isSel ? "checked" : "") + " onchange=\'toggleGp(\\\"" + gp + "\\\", this.checked)\'> 🏛️ " + gp + "</label>";' +
  '        var wardBox = document.createElement("div");' +
  '        wardBox.className = "ward-box";' +
  '        var chips = "<div class=\'ward-chips\'><span style=\'font-size:11px; font-weight:700; color:#475569;\'>वार्ड:</span> ";' +
  '        var isAll = (curVal === "ALL" || !curVal);' +
  '        chips += "<button type=\'button\' class=\'ward-btn " + (isAll ? "all-btn" : "") + "\' onclick=\'setAllWards(\\\"" + gp + "\\\")\'>समस्त (ALL)</button>";' +
  '        for (var w = 1; w <= 15; w++) {' +
  '          var wStr = String(w);' +
  '          var isWSelected = (!isAll && Array.isArray(curVal) && curVal.indexOf(wStr) !== -1);' +
  '          chips += "<button type=\'button\' class=\'ward-btn " + (isWSelected ? "selected" : "") + "\' onclick=\'toggleWardChip(\\\"" + gp + "\\\", \\\"" + wStr + "\\\")\'>" + w + "</button>";' +
  '        }' +
  '        chips += "</div>";' +
  '        wardBox.innerHTML = chips;' +
  '        card.appendChild(head);' +
  '        card.appendChild(wardBox);' +
  '        container.appendChild(card);' +
  '      });' +
  '    }' +
  '    function filterGps(val) { renderList(val); }' +
  '    function toggleGp(gp, checked) {' +
  '      if (checked) { selectedMap[gp] = "ALL"; } else { delete selectedMap[gp]; }' +
  '      renderList(document.getElementById("searchInput").value);' +
  '      updatePreview();' +
  '    }' +
  '    function setAllWards(gp) {' +
  '      selectedMap[gp] = "ALL";' +
  '      renderList(document.getElementById("searchInput").value);' +
  '      updatePreview();' +
  '    }' +
  '    function toggleWardChip(gp, wStr) {' +
  '      var cur = selectedMap[gp];' +
  '      if (cur === "ALL" || !Array.isArray(cur)) { cur = [wStr]; }' +
  '      else {' +
  '        var idx = cur.indexOf(wStr);' +
  '        if (idx !== -1) cur.splice(idx, 1); else cur.push(wStr);' +
  '        if (cur.length === 0) cur = "ALL";' +
  '      }' +
  '      selectedMap[gp] = cur;' +
  '      renderList(document.getElementById("searchInput").value);' +
  '      updatePreview();' +
  '    }' +
  '    function updatePreview() {' +
  '      var keys = Object.keys(selectedMap);' +
  '      if (keys.length === 0) {' +
  '        document.getElementById("previewText").innerHTML = "<strong>दायरा:</strong> कोई पंचायत चयनित नहीं (कृपया ऊपर चेक करें)";' +
  '        return;' +
  '      }' +
  '      var gpListStr = keys.join(", ");' +
  '      var wardScopeParts = [];' +
  '      keys.forEach(function(g){' +
  '        var val = selectedMap[g];' +
  '        if (val === "ALL" || !val) { wardScopeParts.push(g + ": ALL"); }' +
  '        else { wardScopeParts.push(g + ": " + (Array.isArray(val) ? val.join(",") : val)); }' +
  '      });' +
  '      document.getElementById("previewText").innerHTML = "<strong>पंचायतें:</strong> " + gpListStr + "<br><strong>आंशिक वार्ड:</strong> " + wardScopeParts.join(" | ");' +
  '    }' +
  '    function doSave() {' +
  '      var keys = Object.keys(selectedMap);' +
  '      if (keys.length === 0) { alert("कृपया कम से कम एक पंचायत का चयन अवश्य करें!"); return; }' +
  '      var gpListStr = keys.join(", ");' +
  '      var wardScopeParts = [];' +
  '      keys.forEach(function(g){' +
  '        var val = selectedMap[g];' +
  '        if (val === "ALL" || !val) { wardScopeParts.push(g + ": ALL"); }' +
  '        else { wardScopeParts.push(g + ": " + (Array.isArray(val) ? val.join(",") : val)); }' +
  '      });' +
  '      var pWard = document.getElementById("printWardInput").value.trim() || "2";' +
  '      var btn = document.getElementById("saveBtn");' +
  '      btn.disabled = true; btn.textContent = "सुरक्षित हो रहा है...";' +
  '      google.script.run' +
  '        .withSuccessHandler(function(res){' +
  '          if (res && res.success) { alert("✅ " + res.message); google.script.host.close(); }' +
  '          else { alert("⚠️ त्रुटि: " + (res.error || "अज्ञात")); btn.disabled = false; btn.textContent = "सुरक्षित करें"; }' +
  '        })' +
  '        .withFailureHandler(function(err){' +
  '          alert("⚠️ विफलता: " + err); btn.disabled = false; btn.textContent = "सुरक्षित करें";' +
  '        })' +
  '        .savePickerData({' +
  '          row: rowNum,' +
  '          sheetName: sheetName,' +
  '          printWardNo: pWard,' +
  '          allowedPanchayats: gpListStr,' +
  '          allowedPanchayatWards: wardScopeParts.join(" | ")' +
  '        });' +
  '    }' +
  '  </script>' +
  '</body>' +
  '</html>';
}
