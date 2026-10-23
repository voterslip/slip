/**
 * ================================================================================================
 * 🗳️ राजस्थान राज्य निर्वाचन आयोग - पंचायत आम चुनाव 2026
 * कार्यालय उपखंड मजिस्ट्रेट (SDM), भिनाय (अजमेर)
 * 
 * मास्टर Google Apps Script - एकल शीट (Single Master Sheet) व नो-टाइपिंग कंट्रोलर
 * 
 * विशेषताएँ:
 * 1. केवल 2 शीट्स: 'Settings' (पोर्टल सेटिंग्स) और 'Users' (समस्त प्रत्याशी व यूजर एक ही जगह)
 * 2. नो-टाइपिंग सिस्टम: साइडबार व मॉडल डायलॉग से 1-क्लिक में मल्टीपल पंचायतें व आंशिक वार्ड चुनें
 * 3. कमला देवी केवल 'बांदनवाड़ा' पंचायत (PS वार्ड 2)
 * 4. जिला परिषद (राजेन्द्र चौधरी): 3 पंचायतों के आंशिक वार्ड (बांदनवाड़ा: 1,2,3 | देवलियाकलां: 1,2,3,4 | कनईकला: 11,12,13)
 * 5. सरपंच (रामेश्वर जाट): पूरी ग्राम पंचायत, सारे वार्ड (ALL), पर्ची पर वार्ड नहीं छपेगा
 * 6. वार्ड पंच (सुभाष गुर्जर): केवल वार्ड 4 खुलेगा व लॉक रहेगा
 * ================================================================================================
 */

var SHEET_SETTINGS = "Settings";
var SHEET_USERS = "Users";
var SHEET_CANDIDATES = "Candidates"; // बैकवर्ड कम्पैटिबिलिटी हेतु

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
// 1. CUSTOM MENU & SIDEBAR / DIALOG LAUNCHERS
// ================================================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu("🎯 पंचायत व वार्ड चयनकर्ता (NO-TYPING)")
    .addItem("📱 साइडबार खोलें (Sidebar - दाईं तरफ)", "openPanchayatWardSidebar")
    .addItem("🪟 बड़ा पॉपअप खोलें (Popup Dialog - केंद्र में)", "openPanchayatWardPicker")
    .addSeparator()
    .addItem("🚀 1-क्लिक सम्पूर्ण मास्टर सेटअप (Single Sheet Setup)", "setupAllSheets")
    .addItem("🔄 चेकबॉक्स व ड्रॉपडाउन रीफ्रेश करें", "refreshCheckboxes")
    .addSeparator()
    .addItem("ℹ️ वेब ऐप URL व सहायता", "showAppInfo")
    .addToUi();
}

/**
 * साइडबार: शीट देखते-देखते दाईं तरफ से बिना टाइप किए पंचायत व वार्ड चुनने हेतु
 */
function openPanchayatWardSidebar() {
  var html = buildPickerHtml(true);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * मॉडल डायलॉग: बड़ी स्क्रीन पर पॉपअप द्वारा पंचायत व वार्ड चुनने हेतु
 */
function openPanchayatWardPicker() {
  var html = buildPickerHtml(false);
  SpreadsheetApp.getUi().showModalDialog(html, "🎯 पंचायत व आंशिक वार्ड चयनकर्ता (NO-TYPING SYSTEM)");
}

/**
 * पिकर का HTML तैयार करने वाला इंजन
 */
function buildPickerHtml(isSidebar) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS) || ss.getActiveSheet();
  var sheetName = sheet.getName();

  // शीट में मौजूद सभी उपयोगकर्ताओं/प्रत्याशियों की सूची निकालें
  var candidatesList = [];
  var activeRow = 2;
  try {
    var curRange = ss.getActiveRange();
    if (curRange && curRange.getRow() > 1) {
      activeRow = curRange.getRow();
    }
  } catch(e) {}

  if (sheetName === SHEET_USERS || sheetName === SHEET_CANDIDATES) {
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var uCol = headers.indexOf("Username");
    var nameCol = headers.indexOf("Candidate_Name") !== -1 ? headers.indexOf("Candidate_Name") : headers.indexOf("Full_Name");
    var postCol = headers.indexOf("Post");
    var pWardCol = headers.indexOf("Print_Ward_No");
    var gpsCol = headers.indexOf("Allowed_Panchayats");
    var wardsCol = headers.indexOf("Allowed_Panchayat_Wards");

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[uCol]) continue;
      candidatesList.push({
        row: i + 1,
        username: String(row[uCol]),
        name: String(row[nameCol] || row[uCol]),
        post: String(row[postCol] || "पंचायत समिति सदस्य"),
        print_ward_no: String(row[pWardCol] || "2"),
        allowed_panchayats: String(row[gpsCol] || ""),
        allowed_panchayat_wards: String(row[wardsCol] || "")
      });
    }
  }

  var htmlTemplate = HtmlService.createTemplate(getPickerHtmlTemplate());
  htmlTemplate.candidatesJson = JSON.stringify(candidatesList);
  htmlTemplate.panchayatsJson = JSON.stringify(OFFICIAL_PANCHAYATS);
  htmlTemplate.activeRow = activeRow;
  htmlTemplate.isSidebar = isSidebar;

  var html = htmlTemplate.evaluate();
  if (isSidebar) {
    html.setTitle("🎯 पंचायत व वार्ड चयनकर्ता");
  } else {
    html.setWidth(750).setHeight(650);
  }
  return html;
}

/**
 * साइडबार/डायलॉग से डेटा शीट में सुरक्षित करने वाला बैकएंड फंक्शन
 */
function savePickerData(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_USERS) || ss.getSheetByName(SHEET_CANDIDATES);
    if (!sheet) return { success: false, error: "Users शीट नहीं मिली!" };

    var row = parseInt(data.row, 10);
    if (isNaN(row) || row <= 1) {
      return { success: false, error: "कृपया मान्य पंक्ति (Row) चुनें!" };
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var pWardCol = headers.indexOf("Print_Ward_No");
    var gpsCol = headers.indexOf("Allowed_Panchayats");
    var wardsCol = headers.indexOf("Allowed_Panchayat_Wards");
    var postCol = headers.indexOf("Post");

    if (pWardCol !== -1 && data.printWardNo !== undefined) {
      sheet.getRange(row, pWardCol + 1).setValue(data.printWardNo || "ALL");
    }
    if (gpsCol !== -1 && data.allowedPanchayats !== undefined) {
      sheet.getRange(row, gpsCol + 1).setValue(data.allowedPanchayats || "ALL");
    }
    if (wardsCol !== -1 && data.allowedPanchayatWards !== undefined) {
      sheet.getRange(row, wardsCol + 1).setValue(data.allowedPanchayatWards || "ALL");
    }
    if (postCol !== -1 && data.post !== undefined && data.post) {
      sheet.getRange(row, postCol + 1).setValue(data.post);
    }

    SpreadsheetApp.flush();
    return { 
      success: true, 
      message: "पंक्ति " + row + " (" + (data.candidateName || "") + ") में पंचायत व वार्ड विवरण सफलतापूर्वक सुरक्षित हो गया!" 
    };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * ऑन-एडिट ट्रिगर: सेल में ड्रॉपडाउन चुनते समय मल्टी-सेलेक्ट ऑटोमैटिक कॉमा लगाना
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
      list.splice(idx, 1);
    } else {
      list.push(newVal);
    }

    e.range.setValue(list.join(", "));
  }
}

// ================================================================================================
// 2. MASTER 1-CLICK SETUP FOR ALL SHEETS (SINGLE MASTER USERS SHEET)
// ================================================================================================
function setupAllSheets() {
  setupSettingsSheet();
  setupUnifiedUsersSheet();
  
  // पुरानी Candidates शीट यदि हो तो उसे यूजर्स शीट में मर्ज करके हटाएँ या छिपाएँ
  cleanupOldCandidatesSheet();

  SpreadsheetApp.flush();
  try {
    SpreadsheetApp.getUi().alert(
      "✅ सम्पूर्ण एकल-शीट मास्टर सेटअप सफल!\n\n" +
      "1. 'Settings' शीट: पोर्टल नाम, मतदान दिनांक, समय व मॉड्यूल स्विच।\n" +
      "2. 'Users' शीट: एकमात्र मास्टर शीट जिसमें लॉगिन, प्रत्याशी प्रोफाइल, आंशिक वार्ड व चुनाव चिन्ह सब कुछ एक जगह है!\n\n" +
      "🎯 नो-टाइपिंग सुविधा:\n" +
      "अब आपको कुछ भी टाइप करने की आवश्यकता नहीं है! ऊपर मेनू में '🎯 पंचायत व वार्ड चयनकर्ता' > '📱 साइडबार खोलें' पर क्लिक करें और माउस से क्लिक करके पंचायत व वार्ड सुरक्षित करें!"
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

/**
 * एकीकृत मास्टर Users शीट (लॉगिन + प्रत्याशी + वार्ड दायरा + चुनाव चिन्ह सब एक ही जगह)
 */
function setupUnifiedUsersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS) || ss.insertSheet(SHEET_USERS);
  sheet.clear();

  var headers = [
    "Username", "Password", "Candidate_Name", "Mobile", "Post",
    "Print_Ward_No", "Allowed_Panchayats", "Allowed_Panchayat_Wards",
    "Symbol_Name", "Slogan", "Election_Date", "Election_Time",
    "Role", "Status", "Can_Search", "Can_Alpha", "Can_Bulk_Slips",
    "Can_With_Pic_Slips", "Can_Candidate_Profile", "Can_Print", "Can_Download", "Created_At"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#1e3a8a")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var nowStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
  
  var defaultUnifiedRows = [
    // 1. Super Admin: मुख्य व्यवस्थापक (समस्त 30 पंचायतें, सभी वार्ड)
    [
      "admin", "admin2026", "मुख्य व्यवस्थापक (Super Admin)", "7023293283", "मुख्य व्यवस्थापक",
      "ALL", "ALL", "ALL",
      "कमल", "पारदर्शी एवं निष्पक्ष निर्वाचन 2026", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "SUPER_ADMIN", true, true, true, true, true, true, true, true, nowStr
    ],

    // 2. पंचायत समिति सदस्य: कमला देवी (केवल बांदनवाड़ा पंचायत, PS वार्ड 2)
    [
      "kamala", "bhinai2026", "कमला देवी", "9876543210", "पंचायत समिति सदस्य",
      "2", "बांदनवाड़ा", "बांदनवाड़ा: ALL",
      "हाथ", "क्षेत्र के सर्वांगीण विकास एवं खुशहाली हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ],
    [
      "kamla", "bhinai2026", "कमला देवी", "9876543210", "पंचायत समिति सदस्य",
      "2", "बांदनवाड़ा", "बांदनवाड़ा: ALL",
      "हाथ", "क्षेत्र के सर्वांगीण विकास एवं खुशहाली हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ],

    // 3. जिला परिषद सदस्य: राजेन्द्र चौधरी (ZP वार्ड 15 छपेगा, 3 पंचायतों के आंशिक वार्ड खुलेंगे)
    [
      "zp_rajendra", "zp2026", "राजेन्द्र चौधरी", "9829000015", "जिला परिषद सदस्य",
      "15", "बांदनवाड़ा, देवलियाकलां, कनईकला", "बांदनवाड़ा: 1,2,3 | देवलियाकलां: 1,2,3,4 | कनईकला: 11,12,13",
      "कमल", "जिला परिषद क्षेत्र के चहुंमुखी विकास हेतु आपका अपना कर्मठ एवं विश्वस्त साथी।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ],

    // 4. सरपंच प्रत्याशी: रामेश्वर जाट (कनईकलां - पूरी ग्राम पंचायत, सारे वार्ड खुलेंगे, पर्ची पर वार्ड नहीं छपेगा)
    [
      "sarpanch_kanai", "kanai2026", "रामेश्वर जाट", "9829000001", "सरपंच",
      "ALL", "कनईकला", "ALL",
      "उगता सूरज", "ग्राम पंचायत के सर्वांगीण विकास हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ],

    // 5. वार्ड पंच प्रत्याशी: सुभाष गुर्जर (कनईकलां - केवल वार्ड 4 खुलेगा व लॉक रहेगा)
    [
      "panch_w4", "panch2026", "सुभाष गुर्जर", "9829000004", "वार्ड पंच",
      "4", "कनईकला", "4",
      "पतंग", "वार्ड के विकास एवं स्वच्छता हेतु भारी मतों से विजयी बनावें।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ]
  ];

  sheet.getRange(2, 1, defaultUnifiedRows.length, headers.length).setValues(defaultUnifiedRows);

  // Checkbox Validations: Status (14), Can_Search(15), Can_Alpha(16), Can_Bulk_Slips(17), Can_With_Pic_Slips(18), Can_Candidate_Profile(19), Can_Print(20), Can_Download(21)
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  for (var c = 14; c <= 21; c++) {
    sheet.getRange(2, c, 100, 1).setDataValidation(checkRule);
  }

  // Dropdown Validations: Post, Allowed_Panchayats, Symbol_Name
  applySheetDropdowns(sheet, headers);

  sheet.setColumnWidth(1, 130); // Username
  sheet.setColumnWidth(2, 110); // Password
  sheet.setColumnWidth(3, 170); // Candidate_Name
  sheet.setColumnWidth(4, 110); // Mobile
  sheet.setColumnWidth(5, 160); // Post
  sheet.setColumnWidth(6, 120); // Print_Ward_No
  sheet.setColumnWidth(7, 240); // Allowed_Panchayats
  sheet.setColumnWidth(8, 360); // Allowed_Panchayat_Wards
  sheet.setColumnWidth(9, 120); // Symbol_Name
  sheet.setColumnWidth(10, 260); // Slogan
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
}

/**
 * यदि पुरानी Candidates शीट मौजूद हो तो उसे हटाएँ या साफ करें ताकि दो शीट्स का भ्रम न रहे
 */
function cleanupOldCandidatesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cSheet = ss.getSheetByName(SHEET_CANDIDATES);
  if (cSheet) {
    try {
      ss.deleteSheet(cSheet);
    } catch(e) {
      cSheet.hideSheet();
    }
  }
}

/**
 * पद, पंचायत और चुनाव चिन्ह हेतु ड्रॉपडाउन
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
    for (var c = 14; c <= 21; c++) {
      uSheet.getRange(2, c, 100, 1).setDataValidation(checkRule);
    }
    applySheetDropdowns(uSheet, uHeaders);
  }

  try {
    SpreadsheetApp.getUi().alert("✅ समस्त चेकबॉक्स एवं ड्रॉपडाउन सूचियां रीफ्रेश हो गई हैं!");
  } catch(e) {}
}

function showAppInfo() {
  var msg = "ℹ️ वेब ऐप सेटअप निर्देश:\n\n" +
    "1. ऊपर नीले बटन 'Deploy' > 'New deployment' पर क्लिक करें।\n" +
    "2. 'Select type' में 'Web app' चुनें।\n" +
    "3. 'Execute as' में 'Me' चुनें।\n" +
    "4. 'Who has access' में 'Anyone' चुनें।\n" +
    "5. 'Deploy' पर क्लिक करें।\n\n" +
    "🎯 नो-टाइपिंग सुविधा:\n" +
    "पंचायत व आंशिक वार्ड चुनने के लिए '🎯 पंचायत व वार्ड चयनकर्ता' मेनू से साइडबार खोलें!";
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch(e) {}
}

// ================================================================================================
// 3. HTTP GET & POST HANDLERS (PORTAL REAL-TIME SYNC FROM SINGLE MASTER USERS SHEET)
// ================================================================================================
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || "ping";
  var callback = params.callback;
  var res = {};

  try {
    switch (action) {
      case "ping":
        res = { success: true, message: "Panchayat Election API Live (Single Sheet Master)!", timestamp: new Date().toISOString() };
        break;

      case "init":
      case "setup":
      case "setupAllSheets":
        setupSettingsSheet();
        setupUnifiedUsersSheet();
        cleanupOldCandidatesSheet();
        res = { success: true, message: "एकल मास्टर शीट (Settings, Users) सफलतापूर्वक तैयार हो चुकी है!" };
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

  return { success: true, settings: settingsObj };
}

function handleSaveSettings(newSettings) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
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

/**
 * Users शीट से ही उपयोगकर्ताओं की सूची देना
 */
function handleGetUsers(adminUsername, adminPassword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet || sheet.getDataRange().getLastRow() <= 1) {
    setupUnifiedUsersSheet();
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

    var uNameCol = headers.indexOf("Candidate_Name") !== -1 ? headers.indexOf("Candidate_Name") : headers.indexOf("Full_Name");

    users.push({
      username: String(row[0]),
      password: String(row[1]),
      fullName: String(row[uNameCol] || row[0]),
      mobile: String(row[headers.indexOf("Mobile")] || ""),
      role: String(row[headers.indexOf("Role")] || "CANDIDATE"),
      status: (row[headers.indexOf("Status")] === true || String(row[headers.indexOf("Status")]).toUpperCase() === "TRUE") ? "ACTIVE" : "INACTIVE",
      post: String(row[headers.indexOf("Post")] || ""),
      print_ward_no: String(row[headers.indexOf("Print_Ward_No")] || ""),
      allowed_panchayats: String(row[headers.indexOf("Allowed_Panchayats")] || "ALL"),
      allowed_panchayat_wards: String(row[headers.indexOf("Allowed_Panchayat_Wards")] || "ALL"),
      assigned_panchayats: String(row[headers.indexOf("Allowed_Panchayats")] || "ALL"),
      assigned_wards: String(row[headers.indexOf("Allowed_Panchayat_Wards")] || "ALL"),
      allowed_wards: String(row[headers.indexOf("Allowed_Panchayat_Wards")] || "ALL"),
      ward: String(row[headers.indexOf("Print_Ward_No")] || ""),
      symbol_name: String(row[headers.indexOf("Symbol_Name")] || ""),
      slogan: String(row[headers.indexOf("Slogan")] || ""),
      election_date: String(row[headers.indexOf("Election_Date")] || ""),
      election_time: String(row[headers.indexOf("Election_Time")] || ""),
      allowed_tabs: tabs,
      can_print: (row[headers.indexOf("Can_Print")] === true),
      can_download: (row[headers.indexOf("Can_Download")] === true)
    });
  }

  return { success: true, users: users };
}

/**
 * Users शीट से ही समस्त प्रत्याशियों की सूची देना (अब 2 शीट्स की जरूरत नहीं!)
 */
function handleGetCandidates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet || sheet.getDataRange().getLastRow() <= 1) {
    setupUnifiedUsersSheet();
    sheet = ss.getSheetByName(SHEET_USERS);
  }
  if (!sheet) return { success: true, candidates: [] };

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var candidates = [];

  var uCol = headers.indexOf("Username");
  var pCol = headers.indexOf("Password");
  var nameCol = headers.indexOf("Candidate_Name") !== -1 ? headers.indexOf("Candidate_Name") : headers.indexOf("Full_Name");
  var postCol = headers.indexOf("Post");
  var pWardCol = headers.indexOf("Print_Ward_No");
  var gpsCol = headers.indexOf("Allowed_Panchayats");
  var wardsCol = headers.indexOf("Allowed_Panchayat_Wards");
  var symCol = headers.indexOf("Symbol_Name");
  var sloganCol = headers.indexOf("Slogan");
  var dateCol = headers.indexOf("Election_Date");
  var timeCol = headers.indexOf("Election_Time");
  var mobCol = headers.indexOf("Mobile");
  var statusCol = headers.indexOf("Status");
  var printCol = headers.indexOf("Can_Print");
  var dlCol = headers.indexOf("Can_Download");

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[uCol]) continue;

    var cName = String(row[nameCol] || row[uCol]);
    var cPost = String(row[postCol] || "");

    candidates.push({
      username: String(row[uCol]),
      password: String(row[pCol] || ""),
      candidate_name: cName,
      full_name: cName,
      post: cPost,
      print_ward_no: String(row[pWardCol] || ""),
      ward: String(row[pWardCol] || ""),
      allowed_panchayats: String(row[gpsCol] || "ALL"),
      panchayat: String(row[gpsCol] || "ALL"),
      allowed_panchayat_wards: String(row[wardsCol] || "ALL"),
      allowed_wards: String(row[wardsCol] || "ALL"),
      symbol_name: String(row[symCol] || "हाथ"),
      slogan: String(row[sloganCol] || ""),
      election_date: String(row[dateCol] || "15 अक्टूबर 2026"),
      election_time: String(row[timeCol] || "प्रातः 7:00 बजे से सायं 5:00 बजे तक"),
      mobile: String(row[mobCol] || ""),
      status: (row[statusCol] === true || String(row[statusCol]).toUpperCase() === "TRUE") ? "ACTIVE" : "INACTIVE",
      can_print: (row[printCol] === true),
      can_download: (row[dlCol] === true),
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

      var nameCol = headers.indexOf("Candidate_Name") !== -1 ? headers.indexOf("Candidate_Name") : headers.indexOf("Full_Name");

      return {
        success: true,
        user: {
          username: String(row[uCol]),
          fullName: String(row[nameCol] || row[uCol]),
          role: String(row[headers.indexOf("Role")]),
          post: String(row[headers.indexOf("Post")] || ""),
          print_ward_no: String(row[headers.indexOf("Print_Ward_No")] || ""),
          ward: String(row[headers.indexOf("Print_Ward_No")] || ""),
          allowed_panchayats: String(row[headers.indexOf("Allowed_Panchayats")]),
          allowed_panchayat_wards: String(row[headers.indexOf("Allowed_Panchayat_Wards")]),
          allowed_wards: String(row[headers.indexOf("Allowed_Panchayat_Wards")]),
          symbol_name: String(row[headers.indexOf("Symbol_Name")] || ""),
          slogan: String(row[headers.indexOf("Slogan")] || ""),
          election_date: String(row[headers.indexOf("Election_Date")] || ""),
          election_time: String(row[headers.indexOf("Election_Time")] || ""),
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, error: "Users शीट नहीं मिली!" };

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

  SpreadsheetApp.flush();
  return { success: true, message: "प्रत्याशी विवरण सुरक्षित हो गया!" };
}

// ================================================================================================
// 4. EMBEDDED HTML TEMPLATE FOR NO-TYPING SIDEBAR & MODAL PICKER
// ================================================================================================
function getPickerHtmlTemplate() {
  return '<!DOCTYPE html>' +
  '<html>' +
  '<head>' +
  '<base target="_top">' +
  '<style>' +
  '  * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; }' +
  '  body { margin: 0; padding: 12px; background: #f8fafc; color: #1e293b; font-size: 13px; }' +
  '  .header-box { background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%); color: #fff; padding: 12px 14px; border-radius: 8px; margin-bottom: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }' +
  '  .header-title { font-size: 14px; font-weight: 700; margin: 0; display: flex; align-items: center; justify-content: space-between; }' +
  '  .header-sub { font-size: 11px; opacity: 0.9; margin-top: 3px; }' +
  '  .cand-select-bar { background: #fff; border: 1px solid #cbd5e1; border-radius: 6px; padding: 10px; margin-bottom: 10px; }' +
  '  .control-label { font-weight: 700; font-size: 11px; color: #475569; margin-bottom: 4px; display: block; text-transform: uppercase; letter-spacing: 0.5px; }' +
  '  .cand-select { width: 100%; padding: 7px 10px; border: 1px solid #94a3b8; border-radius: 5px; font-size: 13px; font-weight: 600; background: #fff; outline: none; }' +
  '  .cand-select:focus { border-color: #2563eb; ring: 2px solid #bfdbfe; }' +
  '  .row-flex { display: flex; gap: 8px; margin-top: 8px; }' +
  '  .col-half { flex: 1; }' +
  '  .quick-presets { display: flex; flex-wrap: wrap; gap: 5px; margin-bottom: 10px; }' +
  '  .preset-btn { font-size: 11px; padding: 4px 8px; border: 1px solid #cbd5e1; border-radius: 4px; background: #f1f5f9; cursor: pointer; font-weight: 600; color: #334155; }' +
  '  .preset-btn:hover { background: #e2e8f0; color: #0f172a; }' +
  '  .preset-btn.highlight { background: #e0e7ff; color: #3730a3; border-color: #a5b4fc; }' +
  '  .search-box { position: relative; margin-bottom: 10px; }' +
  '  .search-input { width: 100%; padding: 7px 10px 7px 28px; border: 1px solid #cbd5e1; border-radius: 5px; font-size: 12px; outline: none; }' +
  '  .search-icon { position: absolute; left: 8px; top: 8px; font-size: 12px; color: #94a3b8; }' +
  '  .gp-container { max-height: 280px; overflow-y: auto; background: #fff; border: 1px solid #cbd5e1; border-radius: 6px; padding: 6px; margin-bottom: 10px; }' +
  '  .gp-card { border: 1px solid #e2e8f0; border-radius: 5px; padding: 7px 9px; margin-bottom: 6px; background: #fafafa; transition: all 0.15s; }' +
  '  .gp-card.active { border-color: #3b82f6; background: #eff6ff; }' +
  '  .gp-head { display: flex; align-items: center; justify-content: space-between; cursor: pointer; }' +
  '  .gp-label { font-weight: 700; font-size: 12px; color: #1e293b; display: flex; align-items: center; gap: 6px; cursor: pointer; user-select: none; }' +
  '  .ward-box { margin-top: 6px; padding-top: 6px; border-top: 1px dashed #cbd5e1; display: none; }' +
  '  .gp-card.active .ward-box { display: block; }' +
  '  .ward-chips { display: flex; flex-wrap: wrap; gap: 3px; align-items: center; margin-top: 4px; }' +
  '  .ward-btn { font-size: 10px; padding: 2px 6px; border: 1px solid #cbd5e1; border-radius: 3px; background: #fff; cursor: pointer; user-select: none; font-weight: 600; color: #334155; }' +
  '  .ward-btn:hover { background: #f1f5f9; }' +
  '  .ward-btn.selected { background: #2563eb; color: #fff; border-color: #2563eb; }' +
  '  .ward-btn.all-btn { background: #059669; color: #fff; border-color: #059669; }' +
  '  .ward-group-btn { font-size: 9px; padding: 2px 5px; border: 1px solid #cbd5e1; border-radius: 3px; background: #e0e7ff; color: #3730a3; cursor: pointer; font-weight: 600; }' +
  '  .preview-bar { background: #f8fafc; padding: 8px 10px; border-radius: 6px; font-size: 11px; color: #334155; margin-bottom: 10px; border: 1px solid #cbd5e1; line-height: 1.5; }' +
  '  .preview-title { font-weight: 700; color: #0f172a; margin-bottom: 2px; }' +
  '  .preview-val { color: #2563eb; font-weight: 600; word-break: break-all; }' +
  '  .footer-box { display: flex; gap: 8px; align-items: center; }' +
  '  .save-btn { flex: 1; background: #059669; color: #fff; border: none; padding: 10px 14px; font-size: 13px; font-weight: 700; border-radius: 6px; cursor: pointer; text-align: center; box-shadow: 0 2px 4px rgba(5,150,105,0.3); }' +
  '  .save-btn:hover { background: #047857; }' +
  '  .cancel-btn { background: #e2e8f0; color: #475569; border: none; padding: 10px 14px; border-radius: 6px; cursor: pointer; font-weight: 600; }' +
  '  .status-msg { margin-top: 6px; font-size: 11px; font-weight: 600; text-align: center; }' +
  '  .status-msg.success { color: #059669; }' +
  '  .status-msg.error { color: #dc2626; }' +
  '</style>' +
  '</head>' +
  '<body>' +
  '  <div class="header-box">' +
  '    <div class="header-title">' +
  '      <span>🎯 पंचायत व आंशिक वार्ड चयनकर्ता</span>' +
  '      <span style="font-size:10px; background:rgba(255,255,255,0.2); padding:2px 6px; border-radius:10px;">NO-TYPING</span>' +
  '    </div>' +
  '    <div class="header-sub">टाइपिंग की कोई जरूरत नहीं! चेकबॉक्स व वार्ड नंबर दबाकर सुरक्षित करें।</div>' +
  '  </div>' +
  '  <div class="cand-select-bar">' +
  '    <label class="control-label" for="candDropdown">👤 प्रत्याशी / यूजर चुनें:</label>' +
  '    <select id="candDropdown" class="cand-select" onchange="onCandidateChange()"></select>' +
  '    <div class="row-flex">' +
  '      <div class="col-half">' +
  '        <label class="control-label" for="postSelect">🎖️ पद (Post):</label>' +
  '        <select id="postSelect" class="cand-select" style="font-size:11px;" onchange="updatePreview()">' +
  '          <option value="पंचायत समिति सदस्य">पंचायत समिति सदस्य</option>' +
  '          <option value="जिला परिषद सदस्य">जिला परिषद सदस्य</option>' +
  '          <option value="सरपंच">सरपंच</option>' +
  '          <option value="वार्ड पंच">वार्ड पंच</option>' +
  '          <option value="मुख्य व्यवस्थापक">मुख्य व्यवस्थापक</option>' +
  '        </select>' +
  '      </div>' +
  '      <div class="col-half">' +
  '        <label class="control-label" for="printWardInput">🖨️ पर्ची पर छपने वाला वार्ड:</label>' +
  '        <input type="text" id="printWardInput" style="width:100%; padding:6px 8px; border:1px solid #94a3b8; border-radius:5px; font-weight:700; text-align:center;" oninput="updatePreview()">' +
  '      </div>' +
  '    </div>' +
  '  </div>' +
  '  <div class="quick-presets">' +
  '    <button type="button" class="preset-btn highlight" onclick="applyPresetKamala()">⚡ कमला देवी (केवल बांदनवाड़ा)</button>' +
  '    <button type="button" class="preset-btn highlight" onclick="applyPresetRajendra()">⚡ राजेन्द्र चौधरी (ZP 3 पंचायतें)</button>' +
  '    <button type="button" class="preset-btn" onclick="applyPresetSarpanch()">⚡ सरपंच (समस्त वार्ड)</button>' +
  '    <button type="button" class="preset-btn" onclick="clearAllSelections()" style="color:#dc2626;">❌ सब हटाएं</button>' +
  '  </div>' +
  '  <div class="search-box">' +
  '    <span class="search-icon">🔍</span>' +
  '    <input type="text" id="searchInput" class="search-input" placeholder="पंचायत खोजें (उदा. बांदनवाड़ा, देवलिया...)" oninput="filterGps(this.value)">' +
  '  </div>' +
  '  <div class="gp-container" id="gpList"></div>' +
  '  <div class="preview-bar" id="previewBox">' +
  '    <div class="preview-title">👁️ लाइव प्रीव्यू (यह शीट में लिखा जाएगा):</div>' +
  '    <div>🏛️ <strong>पंचायतें:</strong> <span class="preview-val" id="prevGps">-</span></div>' +
  '    <div>🎯 <strong>आंशिक वार्ड:</strong> <span class="preview-val" id="prevWards">-</span></div>' +
  '    <div>🖨️ <strong>पर्ची विवरण:</strong> <span class="preview-val" id="prevPrintSlip">-</span></div>' +
  '  </div>' +
  '  <div class="footer-box">' +
  '    <button type="button" class="save-btn" id="saveBtn" onclick="doSave()">💾 शीट में सुरक्षित करें</button>' +
  '  </div>' +
  '  <div class="status-msg" id="statusMsg"></div>' +
  '  <script>' +
  '    var candidates = JSON.parse(<?= candidatesJson ?>);' +
  '    var allGps = JSON.parse(<?= panchayatsJson ?>);' +
  '    var initialActiveRow = parseInt("<?= activeRow ?>", 10) || 2;' +
  '    var isSidebar = <?= isSidebar ?>;' +
  '    var selectedMap = {};' +
  '    initCandidateDropdown();' +
  '    function initCandidateDropdown() {' +
  '      var d = document.getElementById("candDropdown");' +
  '      d.innerHTML = "";' +
  '      if (candidates.length === 0) {' +
  '        var opt = document.createElement("option");' +
  '        opt.value = initialActiveRow;' +
  '        opt.textContent = "सक्रिय पंक्ति " + initialActiveRow;' +
  '        d.appendChild(opt);' +
  '      } else {' +
  '        var matchIdx = 0;' +
  '        candidates.forEach(function(c, idx){' +
  '          var opt = document.createElement("option");' +
  '          opt.value = c.row;' +
  '          opt.textContent = c.name + " (" + c.post + " - पंक्ति " + c.row + ")";' +
  '          d.appendChild(opt);' +
  '          if (c.row === initialActiveRow) matchIdx = idx;' +
  '        });' +
  '        d.selectedIndex = matchIdx;' +
  '      }' +
  '      loadCandidateData();' +
  '    }' +
  '    function onCandidateChange() {' +
  '      loadCandidateData();' +
  '    }' +
  '    function getSelectedCandidate() {' +
  '      var row = parseInt(document.getElementById("candDropdown").value, 10);' +
  '      for (var i = 0; i < candidates.length; i++) {' +
  '        if (candidates[i].row === row) return candidates[i];' +
  '      }' +
  '      return null;' +
  '    }' +
  '    function loadCandidateData() {' +
  '      var cand = getSelectedCandidate();' +
  '      selectedMap = {};' +
  '      if (!cand) {' +
  '        renderList();' +
  '        updatePreview();' +
  '        return;' +
  '      }' +
  '      document.getElementById("postSelect").value = cand.post || "पंचायत समिति सदस्य";' +
  '      document.getElementById("printWardInput").value = cand.print_ward_no || "2";' +
  '      parseWardsString(cand.allowed_panchayat_wards, cand.allowed_panchayats);' +
  '      renderList();' +
  '      updatePreview();' +
  '    }' +
  '    function parseWardsString(wardsStr, gpsStr) {' +
  '      wardsStr = (wardsStr || "").trim();' +
  '      gpsStr = (gpsStr || "").trim();' +
  '      if (!wardsStr && !gpsStr) return;' +
  '      if (wardsStr === "ALL") {' +
  '        var arr = gpsStr.split(",").map(function(s){ return s.trim(); }).filter(Boolean);' +
  '        arr.forEach(function(g){ selectedMap[g] = "ALL"; });' +
  '        return;' +
  '      }' +
  '      if (wardsStr.indexOf(":") !== -1) {' +
  '        var parts = wardsStr.split("|");' +
  '        parts.forEach(function(p){' +
  '          var sp = p.split(":");' +
  '          if (sp.length === 2) {' +
  '            var g = sp[0].trim();' +
  '            var w = sp[1].trim();' +
  '            if (w === "ALL") { selectedMap[g] = "ALL"; }' +
  '            else { selectedMap[g] = w.split(",").map(function(x){ return x.trim(); }).filter(Boolean); }' +
  '          }' +
  '        });' +
  '      } else if (gpsStr) {' +
  '        var arr = gpsStr.split(",").map(function(s){ return s.trim(); }).filter(Boolean);' +
  '        arr.forEach(function(g){' +
  '          if (wardsStr) selectedMap[g] = wardsStr.split(",").map(function(x){ return x.trim(); }).filter(Boolean);' +
  '          else selectedMap[g] = "ALL";' +
  '        });' +
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
  '        var isAll = (curVal === "ALL" || !curVal);' +
  '        var chips = "<div style=\'display:flex; justify-content:space-between; align-items:center; margin-bottom:3px;\'>" +' +
  '          "<span style=\'font-size:10px; font-weight:700; color:#475569;\'>वार्ड चयन:</span>" +' +
  '          "<div style=\'display:flex; gap:3px;\'>" +' +
  '            "<button type=\'button\' class=\'ward-group-btn\' onclick=\'setWardGroup(\\\"" + gp + "\\\", [1,2,3,4,5])\'>1-5</button>" +' +
  '            "<button type=\'button\' class=\'ward-group-btn\' onclick=\'setWardGroup(\\\"" + gp + "\\\", [6,7,8,9,10])\'>6-10</button>" +' +
  '            "<button type=\'button\' class=\'ward-group-btn\' onclick=\'setWardGroup(\\\"" + gp + "\\\", [11,12,13,14,15])\'>11-15</button>" +' +
  '          "</div>" +' +
  '        "</div>" +' +
  '        "<div class=\'ward-chips\'>" +' +
  '          "<button type=\'button\' class=\'ward-btn " + (isAll ? "all-btn" : "") + "\' onclick=\'setAllWards(\\\"" + gp + "\\\")\'>समस्त (ALL)</button>";' +
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
  '    function setWardGroup(gp, arr) {' +
  '      selectedMap[gp] = arr.map(String);' +
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
  '    function applyPresetKamala() {' +
  '      selectedMap = {};' +
  '      selectedMap["बांदनवाड़ा"] = "ALL";' +
  '      document.getElementById("postSelect").value = "पंचायत समिति सदस्य";' +
  '      document.getElementById("printWardInput").value = "2";' +
  '      renderList();' +
  '      updatePreview();' +
  '    }' +
  '    function applyPresetRajendra() {' +
  '      selectedMap = {};' +
  '      selectedMap["बांदनवाड़ा"] = ["1", "2", "3"];' +
  '      selectedMap["देवलियाकलां"] = ["1", "2", "3", "4"];' +
  '      selectedMap["कनईकला"] = ["11", "12", "13"];' +
  '      document.getElementById("postSelect").value = "जिला परिषद सदस्य";' +
  '      document.getElementById("printWardInput").value = "15";' +
  '      renderList();' +
  '      updatePreview();' +
  '    }' +
  '    function applyPresetSarpanch() {' +
  '      var keys = Object.keys(selectedMap);' +
  '      if (keys.length === 0) selectedMap["कनईकला"] = "ALL";' +
  '      else keys.forEach(function(k){ selectedMap[k] = "ALL"; });' +
  '      document.getElementById("postSelect").value = "सरपंच";' +
  '      document.getElementById("printWardInput").value = "ALL";' +
  '      renderList();' +
  '      updatePreview();' +
  '    }' +
  '    function clearAllSelections() {' +
  '      selectedMap = {};' +
  '      renderList();' +
  '      updatePreview();' +
  '    }' +
  '    function updatePreview() {' +
  '      var keys = Object.keys(selectedMap);' +
  '      var post = document.getElementById("postSelect").value;' +
  '      var pWard = document.getElementById("printWardInput").value.trim();' +
  '      if (keys.length === 0) {' +
  '        document.getElementById("prevGps").textContent = "कोई पंचायत नहीं चुनी गई";' +
  '        document.getElementById("prevWards").textContent = "-";' +
  '        document.getElementById("prevPrintSlip").textContent = "-";' +
  '        return;' +
  '      }' +
  '      var gpStr = keys.join(", ");' +
  '      var wardScopeParts = [];' +
  '      keys.forEach(function(g){' +
  '        var val = selectedMap[g];' +
  '        if (val === "ALL" || !val) { wardScopeParts.push(g + ": ALL"); }' +
  '        else { wardScopeParts.push(g + ": " + (Array.isArray(val) ? val.join(",") : val)); }' +
  '      });' +
  '      var slipText = post;' +
  '      if (pWard && pWard !== "ALL" && post !== "सरपंच") {' +
  '        slipText += " प्रत्याशी • वार्ड नं.-" + pWard;' +
  '      } else {' +
  '        slipText += " प्रत्याशी";' +
  '      }' +
  '      document.getElementById("prevGps").textContent = gpStr;' +
  '      document.getElementById("prevWards").textContent = wardScopeParts.join(" | ");' +
  '      document.getElementById("prevPrintSlip").textContent = slipText;' +
  '    }' +
  '    function doSave() {' +
  '      var keys = Object.keys(selectedMap);' +
  '      if (keys.length === 0) { alert("⚠️ कृपया कम से कम एक पंचायत का चयन करें!"); return; }' +
  '      var row = parseInt(document.getElementById("candDropdown").value, 10);' +
  '      var post = document.getElementById("postSelect").value;' +
  '      var pWard = document.getElementById("printWardInput").value.trim() || "ALL";' +
  '      var gpStr = keys.join(", ");' +
  '      var wardScopeParts = [];' +
  '      keys.forEach(function(g){' +
  '        var val = selectedMap[g];' +
  '        if (val === "ALL" || !val) { wardScopeParts.push(g + ": ALL"); }' +
  '        else { wardScopeParts.push(g + ": " + (Array.isArray(val) ? val.join(",") : val)); }' +
  '      });' +
  '      var btn = document.getElementById("saveBtn");' +
  '      var status = document.getElementById("statusMsg");' +
  '      btn.disabled = true; btn.textContent = "⏳ सुरक्षित हो रहा है...";' +
  '      status.className = "status-msg"; status.textContent = "";' +
  '      var candObj = getSelectedCandidate();' +
  '      var cName = candObj ? candObj.name : ("पंक्ति " + row);' +
  '      google.script.run' +
  '        .withSuccessHandler(function(res){' +
  '          btn.disabled = false; btn.textContent = "💾 शीट में सुरक्षित करें";' +
  '          if (res && res.success) {' +
  '            status.className = "status-msg success";' +
  '            status.textContent = "✅ " + res.message;' +
  '            if (candObj) {' +
  '              candObj.post = post;' +
  '              candObj.print_ward_no = pWard;' +
  '              candObj.allowed_panchayats = gpStr;' +
  '              candObj.allowed_panchayat_wards = wardScopeParts.join(" | ");' +
  '            }' +
  '          } else {' +
  '            status.className = "status-msg error";' +
  '            status.textContent = "⚠️ त्रुटि: " + (res.error || "अज्ञात त्रुटि");' +
  '          }' +
  '        })' +
  '        .withFailureHandler(function(err){' +
  '          btn.disabled = false; btn.textContent = "💾 शीट में सुरक्षित करें";' +
  '          status.className = "status-msg error";' +
  '          status.textContent = "⚠️ विफलता: " + err;' +
  '        })' +
  '        .savePickerData({' +
  '          row: row,' +
  '          candidateName: cName,' +
  '          post: post,' +
  '          printWardNo: pWard,' +
  '          allowedPanchayats: gpStr,' +
  '          allowedPanchayatWards: wardScopeParts.join(" | ")' +
  '        });' +
  '    }' +
  '  </script>' +
  '</body>' +
  '</html>';
}
