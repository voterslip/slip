/**
 * ================================================================================================
 * 🗳️ राजस्थान राज्य निर्वाचन आयोग - पंचायत आम चुनाव 2026
 * कार्यालय उपखंड मजिस्ट्रेट (SDM), भिनाय (अजमेर)
 * 
 * मास्टर Google Apps Script - एकल शीट (Single Master Sheet) व प्रत्याशी नियंत्रण कमान कक्ष (Web App)
 * 
 * विशेषताएँ:
 * 1. प्रत्याशी नियंत्रण कमान कक्ष (Web App HTML): सुपर एडमिन किसी भी प्रत्याशी की सेटिंग बदल सकते हैं,
 *    नया प्रत्याशी जोड़ सकते हैं, और फोटो अपलोड कर सकते हैं जो सीधे गूगल शीट और पोर्टल सेटिंग्स में सिंक होती है।
 * 2. केवल 2 शीट्स: 'Settings' (पोर्टल सेटिंग्स) और 'Users' (समस्त प्रत्याशी, फोटो व यूजर एक ही जगह)
 * 3. नो-टाइपिंग सिस्टम: साइडबार व मॉडल डायलॉग से 1-क्लिक में मल्टीपल पंचायतें व आंशिक वार्ड चुनें
 * 4. प्रत्येक प्रत्याशी का अलग-अलग प्रोफाइल (कमला देवी केवल 'बांदनवाड़ा', अन्य प्रत्याशी अपनी-अपनी पंचायत/वार्ड)
 * 5. फोटो अपलोड सुविधा: वेब ऐप से रंगीन फोटो अपलोड होकर बेस64 डाटा के रूप में गूगल शीट में सुरक्षित होती है।
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

// अधिकृत 22 चुनाव चिन्ह
var OFFICIAL_SYMBOLS = [
  "हाथ", "कमल", "उगता सूरज", "पतंग", "बोतल", "साइकिल",
  "गिलास", "कार", "ताला-चाबी", "बल्ला", "पंखा", "अलमारी", "सिलेंडर", "मोमबत्ती",
  "टेलीविजन", "मेज", "कुर्सी", "कप-प्लेट", "कैमरा", "किताब", "ट्रेक्टर", "रेल का इंजन"
];

// ================================================================================================
// 1. CUSTOM MENU & SIDEBAR / DIALOG LAUNCHERS
// ================================================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu("🎯 पंचायत व प्रत्याशी नियंत्रण कमान कक्ष")
    .addItem("🌐 प्रत्याशी प्रबंधन एवं सेटिंग्स वेब ऐप (Web App)", "openCandidateSettingsDialog")
    .addItem("📱 साइडबार वार्ड चयनकर्ता (Sidebar - दाईं तरफ)", "openPanchayatWardSidebar")
    .addItem("🪟 बड़ा पॉपअप वार्ड चयनकर्ता (Popup Dialog - केंद्र में)", "openPanchayatWardPicker")
    .addSeparator()
    .addItem("🚀 1-क्लिक सम्पूर्ण मास्टर सेटअप (Setup All Sheets)", "setupAllSheets")
    .addItem("🔄 चेकबॉक्स व ड्रॉपडाउन रीफ्रेश करें", "refreshCheckboxes")
    .addSeparator()
    .addItem("ℹ️ वेब ऐप URL व सहायता", "showAppInfo")
    .addToUi();
}

/**
 * सुपर एडमिन प्रत्याशी प्रबंधन एवं सेटिंग्स वेब ऐप को डायलॉग रूप में खोलना
 */
function openCandidateSettingsDialog() {
  var html = HtmlService.createHtmlOutput(buildCandidateSettingsWebHtml())
    .setWidth(1150)
    .setHeight(800)
    .setTitle("🗳️ राजस्थान पंचायत चुनाव 2026 - सुपर एडमिन प्रत्याशी नियंत्रण व सेटिंग्स कमान कक्ष");
  SpreadsheetApp.getUi().showModalDialog(html, "🗳️ सुपर एडमिन प्रत्याशी नियंत्रण व सेटिंग्स कमान कक्ष");
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
        post: String(row[postCol] || ""),
        print_ward_no: String(row[pWardCol] || ""),
        allowed_panchayats: String(row[gpsCol] || ""),
        allowed_panchayat_wards: String(row[wardsCol] || "")
      });
    }
  }

  var tmpl = HtmlService.createTemplate(getPickerHtmlTemplate());
  tmpl.candidatesJson = JSON.stringify(candidatesList);
  tmpl.panchayatsJson = JSON.stringify(OFFICIAL_PANCHAYATS);
  tmpl.activeRow = activeRow;
  tmpl.isSidebar = isSidebar;

  var output = tmpl.evaluate();
  if (isSidebar) {
    output.setTitle("🎯 पंचायत व वार्ड चयनकर्ता");
  } else {
    output.setWidth(680).setHeight(620);
  }
  return output;
}

/**
 * HTML साइडबार/मॉडल से डेटा शीट में सुरक्षित करने वाला बैकएंड फंक्शन
 */
function savePickerData(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) {
      return { success: false, error: "Users शीट नहीं मिली!" };
    }

    var row = parseInt(data.row, 10);
    if (isNaN(row) || row < 2) {
      return { success: false, error: "अमान्य पंक्ति संख्या: " + data.row };
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
  cleanupOldCandidatesSheet();

  SpreadsheetApp.flush();
  try {
    SpreadsheetApp.getUi().alert(
      "✅ सम्पूर्ण एकल-शीट मास्टर सेटअप सफल!\n\n" +
      "1. 'Settings' शीट: पोर्टल नाम, मतदान दिनांक, समय व मॉड्यूल स्विच।\n" +
      "2. 'Users' शीट: एकमात्र मास्टर शीट जिसमें लॉगिन, प्रत्याशी प्रोफाइल, फोटो (Photo_Url), आंशिक वार्ड व चुनाव चिन्ह सब कुछ एक जगह है!\n\n" +
      "🌐 प्रत्याशी नियंत्रण वेब ऐप:\n" +
      "ऊपर मेनू में '🎯 पंचायत व प्रत्याशी नियंत्रण कमान कक्ष' > '🌐 प्रत्याशी प्रबंधन एवं सेटिंग्स वेब ऐप' पर क्लिक करें और नए प्रत्याशी जोड़ें या फोटो अपलोड करें!"
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
    ["ALLOW_VOTER_PHOTO_ON_NORMAL_SLIPS", false, "CHECKBOX", "सामान्य पर्चियों में मतदाता फोटो शामिल करें (TRUE/FALSE)"],
    ["SHOW_WATERMARK", true, "CHECKBOX", "पर्ची पर वाटरमार्क दिखाएं"],
    ["DEFAULT_USER_PASSWORD", "bhinai2026", "TEXT", "नये उपयोगकर्ताओं का डिफ़ॉल्ट पासवर्ड"],
    ["LAST_CANDIDATE_SYNC", new Date().toISOString(), "TEXT", "अंतिम प्रत्याशी सिंक समय"]
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
 * उपयोगकर्ताओं की शीट में आवश्यक कॉलम्स की जांच व स्वतः जोड़ना
 */
function ensureUserSheetHeaders(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var required = ["Custom_Symbol_Url", "Photo_Url"];
  
  required.forEach(function(req) {
    if (headers.indexOf(req) === -1) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(req)
        .setBackground("#1e3a8a")
        .setFontColor("#ffffff")
        .setFontWeight("bold");
    }
  });
}

/**
 * एकीकृत मास्टर Users शीट (लॉगिन + प्रत्याशी + फोटो + वार्ड दायरा + चुनाव चिन्ह सब एक ही जगह)
 */
function setupUnifiedUsersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS) || ss.insertSheet(SHEET_USERS);
  sheet.clear();

  var headers = [
    "Username", "Password", "Candidate_Name", "Mobile", "Post",
    "Print_Ward_No", "Allowed_Panchayats", "Allowed_Panchayat_Wards",
    "Symbol_Name", "Custom_Symbol_Url", "Photo_Url", "Slogan", "Election_Date", "Election_Time",
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
      "कमल", "", "", "पारदर्शी एवं निष्पक्ष निर्वाचन 2026", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "SUPER_ADMIN", true, true, true, true, true, true, true, true, nowStr
    ],

    // 2. पंचायत समिति सदस्य: कमला देवी (केवल बांदनवाड़ा पंचायत, PS वार्ड 2, हाथ चिन्ह)
    [
      "kamala", "bhinai2026", "कमला देवी", "9876543210", "पंचायत समिति सदस्य",
      "2", "बांदनवाड़ा", "बांदनवाड़ा: ALL",
      "हाथ", "", "./assets/kamala_photo.jpg", "क्षेत्र के सर्वांगीण विकास एवं खुशहाली हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ],
    [
      "kamla", "bhinai2026", "कमला देवी", "9876543210", "पंचायत समिति सदस्य",
      "2", "बांदनवाड़ा", "बांदनवाड़ा: ALL",
      "हाथ", "", "./assets/kamala_photo.jpg", "क्षेत्र के सर्वांगीण विकास एवं खुशहाली हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ],

    // 3. जिला परिषद सदस्य: राजेन्द्र चौधरी (ZP वार्ड 15 छपेगा, 3 पंचायतों के आंशिक वार्ड खुलेंगे)
    [
      "zp_rajendra", "zp2026", "राजेन्द्र चौधरी", "9829000015", "जिला परिषद सदस्य",
      "15", "बांदनवाड़ा, देवलियाकलां, कनईकला", "बांदनवाड़ा: 1,2,3 | देवलियाकलां: 1,2,3,4 | कनईकला: 11,12,13",
      "कमल", "", "", "जिला परिषद क्षेत्र के चहुंमुखी विकास हेतु आपका अपना कर्मठ एवं विश्वस्त साथी।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ],

    // 4. सरपंच प्रत्याशी: रामेश्वर जाट (कनईकलां - पूरी ग्राम पंचायत, सारे वार्ड खुलेंगे, पर्ची पर वार्ड नहीं छपेगा)
    [
      "sarpanch_kanai", "kanai2026", "रामेश्वर जाट", "9829000001", "सरपंच",
      "ALL", "कनईकला", "ALL",
      "उगता सूरज", "", "", "ग्राम पंचायत के सर्वांगीण विकास हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ],

    // 5. वार्ड पंच प्रत्याशी: सुभाष गुर्जर (कनईकलां - केवल वार्ड 4 खुलेगा व लॉक रहेगा)
    [
      "panch_w4", "panch2026", "सुभाष गुर्जर", "9829000004", "वार्ड पंच",
      "4", "कनईकला", "4",
      "पतंग", "", "", "वार्ड के विकास एवं स्वच्छता हेतु भारी मतों से विजयी बनावें।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ]
  ];

  sheet.getRange(2, 1, defaultUnifiedRows.length, headers.length).setValues(defaultUnifiedRows);

  // Checkbox Validations: Status (16), Can_Search(17) to Can_Download(23)
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  for (var c = 16; c <= 23; c++) {
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
  sheet.setColumnWidth(10, 160); // Custom_Symbol_Url
  sheet.setColumnWidth(11, 200); // Photo_Url
  sheet.setColumnWidth(12, 260); // Slogan
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
}

function cleanupOldCandidatesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cSheet = ss.getSheetByName(SHEET_CANDIDATES);
  if (cSheet) {
    try {
      cSheet.hideSheet();
    } catch(e) {}
  }
}

function applySheetDropdowns(sheet, headers) {
  var postCol = headers.indexOf("Post") + 1;
  var symbolCol = headers.indexOf("Symbol_Name") + 1;
  var gpCol = headers.indexOf("Allowed_Panchayats") + 1;

  var postRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["सरपंच", "वार्ड पंच", "उपसरपंच", "पंचायत समिति सदस्य", "जिला परिषद सदस्य", "मुख्य व्यवस्थापक", "बी.एल.ओ."], true)
    .build();
  if (postCol > 0) sheet.getRange(2, postCol, 100, 1).setDataValidation(postRule);

  var symbolRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(OFFICIAL_SYMBOLS, true)
    .build();
  if (symbolCol > 0) sheet.getRange(2, symbolCol, 100, 1).setDataValidation(symbolRule);

  var gpRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["ALL"].concat(OFFICIAL_PANCHAYATS), true)
    .build();
  if (gpCol > 0) sheet.getRange(2, gpCol, 100, 1).setDataValidation(gpRule);
}

function refreshCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return;
  ensureUserSheetHeaders(sheet);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  applySheetDropdowns(sheet, headers);
  SpreadsheetApp.getUi().alert("✅ ड्रॉपडाउन एवं कॉलम्स सफलतापूर्वक रीफ्रेश हो गए!");
}

function showAppInfo() {
  var url = ScriptApp.getService().getUrl();
  var msg = "🌐 वेब ऐप लाइव URL:\n" + (url || "वेब ऐप को अभी डिप्लॉय करें (Deploy as Web App)") + "\n\n" +
            "📌 उपयोग निर्देश:\n" +
            "1. सुपर एडमिन वेब ऐप खोलने हेतु ऊपर मेनू में '🌐 प्रत्याशी प्रबंधन एवं सेटिंग्स वेब ऐप' चुनें।\n" +
            "2. वेब ऐप से आप सीधे फोटो अपलोड कर सकते हैं, नया प्रत्याशी जोड़ सकते हैं और सेटिंग्स बदल सकते हैं।";
  SpreadsheetApp.getUi().alert(msg);
}

// ================================================================================================
// 3. HTTP GET & POST HANDLERS (WEB APP & API)
// ================================================================================================

/**
 * doGet: वेब ऐप HTML पेज सर्व करता है अथवा JSON API रिस्पॉन्स देता है
 */
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = (params.action || "").trim();
  var callback = params.callback;

  // JSON API कॉल्स की सूची
  var jsonActions = ["ping", "init", "setup", "setupAllSheets", "getSettings", "saveSettings", "updateSetting", "getUsers", "getCandidates", "getPortalJson", "getFullJson", "syncJson", "login", "saveCandidate", "deleteCandidate", "updatePassword", "changePassword"];

  // 1. यदि कोई JSON एक्शन नहीं है अथवा वेब पेज मांगा गया है, तो सीधे सम्पूर्ण प्रत्याशी प्रबंधन वेब ऐप HTML आउटपुट दें!
  if (!action || !jsonActions.includes(action) || action === "candidateSettings" || action === "candidate_settings" || action === "html" || action === "admin") {
    var htmlContent = "";
    try {
      htmlContent = HtmlService.createTemplateFromFile("CandidateSettings").evaluate().getContent();
    } catch(e1) {
      try {
        htmlContent = HtmlService.createTemplateFromFile("Index").evaluate().getContent();
      } catch(e2) {
        htmlContent = buildCandidateSettingsWebHtml();
      }
    }
    return HtmlService.createHtmlOutput(htmlContent)
      .setTitle("🗳️ राजस्थान पंचायत आम चुनाव 2026 - सुपर एडमिन प्रत्याशी सेटिंग्स एवं नियंत्रण कक्ष")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 2. API कॉल्स का प्रबंधन
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

      case "saveSettings":
      case "updateSetting":
        var sData = {};
        if (params.settings) {
          try { sData = JSON.parse(params.settings); } catch(ex) { sData = params; }
        } else if (params.key && params.value !== undefined) {
          var pVal = params.value;
          if (params.value === "true" || params.value === true) pVal = true;
          else if (params.value === "false" || params.value === false) pVal = false;
          sData[params.key] = pVal;
        } else {
          sData = params;
        }
        res = handleSaveSettings(sData);
        break;

      case "getUsers":
        res = handleGetUsers(params.adminUsername, params.adminPassword);
        break;

      case "getCandidates":
        res = handleGetCandidates();
        break;

      case "getPortalJson":
      case "getFullJson":
      case "syncJson":
        res = handleGetPortalJson();
        break;

      case "login":
        res = handleLogin(params.username, params.password);
        break;

      case "saveCandidate":
        // GET द्वारा भी प्रत्याशी सुरक्षित करने की सुविधा (CORS फ्री)
        var cData = {};
        if (params.data) {
          try { cData = JSON.parse(params.data); } catch(ex) { cData = params; }
        } else {
          cData = params;
        }
        res = handleSaveCandidate(cData);
        break;

      case "deleteCandidate":
        res = handleDeleteCandidate(params.username || params.user_id);
        break;

      case "updatePassword":
      case "changePassword":
        res = handleUpdatePassword(params.username || params.user_id, params.newPassword || params.password);
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

/**
 * doPost: सेटिंग्स, प्रत्याशी विवरण व फोटो अपलोड सुरक्षित करने हेतु
 */
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
    } else if (action === "saveCandidate" || action === "saveCandidateSelfEdit") {
      res = handleSaveCandidate(payload.candidateData || payload.candidate || payload);
    } else if (action === "updatePassword" || action === "changePassword") {
      res = handleUpdatePassword(payload.username || payload.user_id, payload.newPassword || payload.password);
    } else if (action === "deleteCandidate") {
      res = handleDeleteCandidate(payload.username || payload.user_id);
    } else if (action === "syncPortalJson") {
      res = handleGetPortalJson();
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
  if (!sheet || sheet.getDataRange().getLastRow() <= 1) {
    setupSettingsSheet();
    sheet = ss.getSheetByName(SHEET_SETTINGS);
  }
  if (!sheet) return { success: false, error: "Settings शीट नहीं मिली!" };

  var data = sheet.getDataRange().getValues();
  var existingKeys = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    if (key) existingKeys[key] = i + 1;
  }

  for (var k in newSettings) {
    if (newSettings[k] !== undefined) {
      if (existingKeys[k]) {
        sheet.getRange(existingKeys[k], 2).setValue(newSettings[k]);
      } else {
        var newRow = sheet.getLastRow() + 1;
        var type = (typeof newSettings[k] === "boolean") ? "CHECKBOX" : "TEXT";
        sheet.getRange(newRow, 1, 1, 4).setValues([[k, newSettings[k], type, "अपडेटेड सेटिंग"]]);
      }
    }
  }
  SpreadsheetApp.flush();
  return { success: true, message: "सेटिंग्स गूगल शीट में सफलतापूर्वक अपडेट हुईं!", settings: handleGetSettings().settings };
}

function apiSaveSettings(newSettings) {
  return handleSaveSettings(newSettings);
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
  ensureUserSheetHeaders(sheet);

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var users = [];

  var photoCol = headers.indexOf("Photo_Url");
  var custSymCol = headers.indexOf("Custom_Symbol_Url");

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
    var uNameStr = String(row[0]).toLowerCase();

    var photoVal = (photoCol !== -1 && row[photoCol]) ? String(row[photoCol]) : "";
    if (!photoVal && (uNameStr === "kamala" || uNameStr === "kamla")) {
      photoVal = "./assets/kamala_photo.jpg";
    }

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
      custom_symbol_url: (custSymCol !== -1 && row[custSymCol]) ? String(row[custSymCol]) : "",
      photo_url: photoVal,
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
 * Users शीट से ही समस्त प्रत्याशियों की सूची देना
 */
function handleGetCandidates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet || sheet.getDataRange().getLastRow() <= 1) {
    setupUnifiedUsersSheet();
    sheet = ss.getSheetByName(SHEET_USERS);
  }
  if (!sheet) return { success: true, candidates: [] };
  ensureUserSheetHeaders(sheet);

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
  var custSymCol = headers.indexOf("Custom_Symbol_Url");
  var photoCol = headers.indexOf("Photo_Url");
  var sloganCol = headers.indexOf("Slogan");
  var dateCol = headers.indexOf("Election_Date");
  var timeCol = headers.indexOf("Election_Time");
  var mobCol = headers.indexOf("Mobile");
  var statusCol = headers.indexOf("Status");
  var printCol = headers.indexOf("Can_Print");
  var dlCol = headers.indexOf("Can_Download");
  var roleCol = headers.indexOf("Role");

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[uCol]) continue;

    var uName = String(row[uCol]);
    var uNameLower = uName.toLowerCase();
    var cName = String(row[nameCol] || row[uCol]);
    var cPost = String(row[postCol] || "");

    // फोटो रेजोल्यूशन: यदि कमला देवी है तो डिफ़ॉल्ट कमला फोटो, अन्यथा शीट में सुरक्षित फोटो
    var photoVal = (photoCol !== -1 && row[photoCol]) ? String(row[photoCol]) : "";
    if (!photoVal && (uNameLower === "kamala" || uNameLower === "kamla" || cName.indexOf("कमला") !== -1)) {
      photoVal = "./assets/kamala_photo.jpg";
    }

    candidates.push({
      id: "cand_" + uName,
      user_id: "cand_" + uName,
      username: uName,
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
      symbol_name: String(row[symCol] || (uNameLower.indexOf("kamla") !== -1 ? "हाथ" : "उगता सूरज")),
      symbol_icon: (symCol !== -1 && row[symCol]) ? String(row[symCol]) : "sun",
      custom_symbol_url: (custSymCol !== -1 && row[custSymCol]) ? String(row[custSymCol]) : "",
      photo_url: photoVal,
      slogan: String(row[sloganCol] || ""),
      election_date: String(row[dateCol] || "15 अक्टूबर 2026"),
      election_time: String(row[timeCol] || "प्रातः 7:00 बजे से सायं 5:00 बजे तक"),
      mobile: String(row[mobCol] || ""),
      role: (roleCol !== -1 && row[roleCol]) ? String(row[roleCol]) : "CANDIDATE",
      status: (row[statusCol] === true || String(row[statusCol]).toUpperCase() === "TRUE") ? "ACTIVE" : "INACTIVE",
      can_print: (row[printCol] === true),
      can_download: (row[dlCol] === true),
      show_banner_on_slip: true,
      allowed_tabs: ["searchTab", "alphaTab", "bulkSlipTab", "withPicSlipTab", "candidateProfileTab"]
    });
  }

  return { success: true, candidates: candidates };
}

/**
 * सम्पूर्ण पोर्टल डेटा JSON (portal_users.json फॉर्मेट) तैयार करके देना
 */
function handleGetPortalJson() {
  var candRes = handleGetCandidates();
  var settRes = handleGetSettings();

  var usersList = [];
  var candidatesMap = {};

  (candRes.candidates || []).forEach(function(c) {
    var uId = c.id || ("cand_" + c.username);
    var uName = c.username;

    var candProfile = {
      user_id: uId,
      username: uName,
      candidate_name: c.candidate_name,
      full_name: c.candidate_name,
      mobile: c.mobile || "",
      post: c.post || "सरपंच",
      panchayat: c.panchayat || c.allowed_panchayats || "ALL",
      ward: c.print_ward_no || c.ward || "",
      print_ward_no: c.print_ward_no || c.ward || "",
      allowed_panchayats: c.allowed_panchayats || "ALL",
      allowed_wards: c.allowed_panchayat_wards || c.allowed_wards || "ALL",
      allowed_panchayat_wards: c.allowed_panchayat_wards || c.allowed_wards || "ALL",
      symbol_name: c.symbol_name || (uName === "kamla" ? "हाथ" : "उगता सूरज"),
      symbol_icon: c.symbol_icon || "sun",
      custom_symbol_url: c.custom_symbol_url || "",
      photo_url: c.photo_url || (uName === "kamla" || uName === "kamala" ? "./assets/kamala_photo.jpg" : ""),
      slogan: c.slogan || "क्षेत्र के सर्वांगीण विकास एवं खुशहाली हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।",
      election_date: c.election_date || "15 अक्टूबर 2026",
      election_time: c.election_time || "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      show_banner_on_slip: true
    };

    candidatesMap[uId] = candProfile;
    candidatesMap[uName] = candProfile;

    usersList.push({
      id: uId,
      username: uName,
      password: c.password || "bhinai2026",
      full_name: c.candidate_name,
      mobile: c.mobile || "",
      role: c.role || "CANDIDATE",
      status: c.status || "ACTIVE",
      allowed_panchayats: c.allowed_panchayats || "ALL",
      allowed_wards: c.allowed_panchayat_wards || "ALL",
      allowed_tabs: c.allowed_tabs || ["searchTab", "alphaTab", "bulkSlipTab", "withPicSlipTab", "candidateProfileTab"],
      candidate_mode: "user_edit",
      can_print: c.can_print !== false,
      can_download: c.can_download !== false,
      can_search: true,
      can_view: true,
      candidate: candProfile
    });
  });

  return {
    success: true,
    portal_users: {
      version: "2.1.0",
      election_date: (settRes.settings && settRes.settings.ELECTION_DATE) || "15 अक्टूबर 2026",
      election_time: (settRes.settings && settRes.settings.ELECTION_TIME) || "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      last_updated: new Date().toISOString(),
      users: usersList,
      candidates: candidatesMap
    },
    candidates: candRes.candidates || [],
    settings: settRes.settings || {}
  };
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
      var photoCol = headers.indexOf("Photo_Url");
      var custSymCol = headers.indexOf("Custom_Symbol_Url");

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
          custom_symbol_url: (custSymCol !== -1 && row[custSymCol]) ? String(row[custSymCol]) : "",
          photo_url: (photoCol !== -1 && row[photoCol]) ? String(row[photoCol]) : "",
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

/**
 * सुपर एडमिन द्वारा प्रत्याशी जोड़ना, संपादित करना व फोटो अपलोड गूगल शीट में सुरक्षित करना
 */
function handleSaveCandidate(cand) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) {
    setupUnifiedUsersSheet();
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  }
  if (!sheet) return { success: false, error: "Users शीट नहीं मिली!" };

  ensureUserSheetHeaders(sheet);

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var uCol = headers.indexOf("Username");
  var targetRow = -1;

  var uName = String(cand.username || cand.user_id || "").trim();
  if (!uName) {
    // नया यूजरनेम स्वतः जनरेट करें
    var prefix = cand.post ? cand.post.toLowerCase().replace(/[^a-z0-9]/g, "") : "cand";
    uName = prefix + "_" + Math.floor(1000 + Math.random() * 9000);
  }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uCol]).toLowerCase() === uName.toLowerCase()) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, uCol + 1).setValue(uName);
    var nowStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
    var cAtCol = headers.indexOf("Created_At");
    if (cAtCol !== -1) sheet.getRange(targetRow, cAtCol + 1).setValue(nowStr);
  }

  var fieldMap = {
    "Candidate_Name": cand.candidate_name || cand.full_name || cand.fullName,
    "Post": cand.post,
    "Print_Ward_No": cand.print_ward_no || cand.ward,
    "Allowed_Panchayats": cand.allowed_panchayats || cand.panchayat,
    "Allowed_Panchayat_Wards": cand.allowed_panchayat_wards || cand.allowed_wards,
    "Symbol_Name": cand.symbol_name,
    "Custom_Symbol_Url": cand.custom_symbol_url || cand.customSymbolUrl || "",
    "Photo_Url": cand.photo_url || cand.photo || cand.photoUrl || "",
    "Slogan": cand.slogan,
    "Election_Date": cand.election_date || cand.electionDate,
    "Election_Time": cand.election_time || cand.electionTime,
    "Mobile": cand.mobile,
    "Role": cand.role || "CANDIDATE",
    "Status": (cand.status === false || String(cand.status).toUpperCase() === "INACTIVE") ? false : true,
    "Can_Print": (cand.can_print !== false),
    "Can_Download": (cand.can_download !== false),
    "Can_Search": true,
    "Can_Alpha": true,
    "Can_Bulk_Slips": true,
    "Can_With_Pic_Slips": true,
    "Can_Candidate_Profile": true
  };

  if (cand.password) {
    fieldMap["Password"] = cand.password;
  } else if (targetRow > data.length) {
    fieldMap["Password"] = "bhinai2026";
  }

  for (var key in fieldMap) {
    var cIdx = headers.indexOf(key);
    if (cIdx !== -1 && fieldMap[key] !== undefined) {
      sheet.getRange(targetRow, cIdx + 1).setValue(fieldMap[key]);
    }
  }

  // Settings शीट में अंतिम सिंक समय अपडेट करें
  try {
    var setSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
    if (setSheet) {
      var sData = setSheet.getDataRange().getValues();
      for (var s = 1; s < sData.length; s++) {
        if (String(sData[s][0]).trim() === "LAST_CANDIDATE_SYNC") {
          setSheet.getRange(s + 1, 2).setValue(new Date().toISOString());
          break;
        }
      }
    }
  } catch(e) {}

  SpreadsheetApp.flush();

  var updatedPortal = handleGetPortalJson();
  return { 
    success: true, 
    message: "प्रत्याशी विवरण व फोटो गूगल शीट में सफलतापूर्वक सुरक्षित हो गया!",
    username: uName,
    portal_users: updatedPortal.portal_users,
    candidates: updatedPortal.candidates
  };
}

/**
 * प्रत्याशी खाता हटाना
 */
function handleDeleteCandidate(username) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, error: "Users शीट नहीं मिली!" };

  var data = sheet.getDataRange().getValues();
  var uCol = data[0].indexOf("Username");
  if (uCol === -1) return { success: false, error: "Username कॉलम नहीं मिला!" };

  var targetRow = -1;
  var uName = String(username).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uCol]).trim().toLowerCase() === uName) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    return { success: false, error: "प्रत्याशी '" + username + "' नहीं मिला!" };
  }

  sheet.deleteRow(targetRow);
  SpreadsheetApp.flush();

  return { success: true, message: "प्रत्याशी '" + username + "' सफलतापूर्वक हटा दिया गया!" };
}

/**
 * यूजर/प्रत्याशी का पासवर्ड गूगल शीट में सीधे अपडेट करना
 */
function handleUpdatePassword(username, newPassword) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, error: "Users शीट नहीं मिली!" };

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var uCol = headers.indexOf("Username");
  var pCol = headers.indexOf("Password");
  if (uCol === -1 || pCol === -1) return { success: false, error: "Username या Password कॉलम नहीं मिला!" };

  var uName = String(username || "").trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uCol]).trim().toLowerCase() === uName) {
      sheet.getRange(i + 1, pCol + 1).setValue(newPassword);
      SpreadsheetApp.flush();
      return { success: true, message: "पासवर्ड गूगल शीट में सफलतापूर्वक अपडेट हो गया!" };
    }
  }

  return { success: false, error: "यूजर '" + username + "' Users शीट में नहीं मिला!" };
}

// Google Script Run API Wrappers (वेब ऐप फ्रंटएंड से सीधे कॉल हेतु)
function apiGetCandidates() {
  return handleGetCandidates();
}

function apiSaveCandidate(candData) {
  return handleSaveCandidate(candData);
}

function apiDeleteCandidate(username) {
  return handleDeleteCandidate(username);
}

function apiGetPortalJson() {
  return handleGetPortalJson();
}

// ================================================================================================
// 4. EMBEDDED HTML TEMPLATE FOR NO-TYPING SIDEBAR & MODAL PICKER
// ================================================================================================
function getPickerHtmlTemplate() {
  return '<!DOCTYPE html>' +
  '<html><head><base target="_top">' +
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
  '  .preview-bar { background: #f8fafc; padding: 8px 10px; border-radius: 6px; font-size: 11px; color: #334155; margin-bottom: 10px; border: 1px solid #cbd5e1; line-height: 1.5; }' +
  '  .preview-title { font-weight: 700; color: #0f172a; margin-bottom: 2px; }' +
  '  .preview-val { color: #2563eb; font-weight: 600; word-break: break-all; }' +
  '  .footer-box { display: flex; gap: 8px; align-items: center; }' +
  '  .save-btn { flex: 1; background: #059669; color: #fff; border: none; padding: 10px 14px; font-size: 13px; font-weight: 700; border-radius: 6px; cursor: pointer; text-align: center; box-shadow: 0 2px 4px rgba(5,150,105,0.3); }' +
  '  .save-btn:hover { background: #047857; }' +
  '  .status-msg { margin-top: 6px; font-size: 11px; font-weight: 600; text-align: center; }' +
  '  .status-msg.success { color: #059669; }' +
  '  .status-msg.error { color: #dc2626; }' +
  '</style></head>' +
  '<body>' +
  '  <div class="header-box">' +
  '    <div class="header-title"><span>🎯 पंचायत व आंशिक वार्ड चयनकर्ता</span><span style="font-size:10px; background:rgba(255,255,255,0.2); padding:2px 6px; border-radius:10px;">NO-TYPING</span></div>' +
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
  '    <div class="preview-title">👁️ लाइव प्रीव्यू:</div>' +
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
  '    var selectedMap = {};' +
  '    initCandidateDropdown();' +
  '    function initCandidateDropdown() {' +
  '      var d = document.getElementById("candDropdown");' +
  '      d.innerHTML = "";' +
  '      candidates.forEach(function(c, idx){' +
  '        var opt = document.createElement("option");' +
  '        opt.value = c.row;' +
  '        opt.textContent = c.name + " (" + c.post + " - पंक्ति " + c.row + ")";' +
  '        d.appendChild(opt);' +
  '        if (c.row === initialActiveRow) d.selectedIndex = idx;' +
  '      });' +
  '      loadCandidateData();' +
  '    }' +
  '    function onCandidateChange() { loadCandidateData(); }' +
  '    function getSelectedCandidate() {' +
  '      var row = parseInt(document.getElementById("candDropdown").value, 10);' +
  '      for (var i = 0; i < candidates.length; i++) { if (candidates[i].row === row) return candidates[i]; }' +
  '      return null;' +
  '    }' +
  '    function loadCandidateData() {' +
  '      var c = getSelectedCandidate();' +
  '      selectedMap = {};' +
  '      if (!c) { renderGpList(); updatePreview(); return; }' +
  '      if (c.post) document.getElementById("postSelect").value = c.post;' +
  '      document.getElementById("printWardInput").value = c.print_ward_no || "";' +
  '      var gps = (c.allowed_panchayats || "").split(",").map(function(s){return s.trim();}).filter(Boolean);' +
  '      var rawWards = c.allowed_panchayat_wards || "";' +
  '      if (c.allowed_panchayats === "ALL") {' +
  '        allGps.forEach(function(g){ selectedMap[g] = "ALL"; });' +
  '      } else {' +
  '        gps.forEach(function(g){ selectedMap[g] = "ALL"; });' +
  '        if (rawWards && rawWards !== "ALL") {' +
  '          var parts = rawWards.split("|").map(function(s){return s.trim();});' +
  '          parts.forEach(function(p){' +
  '            var colonIdx = p.indexOf(":");' +
  '            if (colonIdx !== -1) {' +
  '              var gName = p.substring(0, colonIdx).trim();' +
  '              var wStr = p.substring(colonIdx + 1).trim();' +
  '              if (wStr === "ALL") { selectedMap[gName] = "ALL"; }' +
  '              else { selectedMap[gName] = wStr.split(",").map(function(s){return s.trim();}).filter(Boolean); }' +
  '            }' +
  '          });' +
  '        }' +
  '      }' +
  '      renderGpList();' +
  '      updatePreview();' +
  '    }' +
  '    function renderGpList(filterText) {' +
  '      var container = document.getElementById("gpList");' +
  '      container.innerHTML = "";' +
  '      var q = (filterText || "").trim().toLowerCase();' +
  '      allGps.forEach(function(gp){' +
  '        if (q && gp.toLowerCase().indexOf(q) === -1) return;' +
  '        var isSel = (selectedMap[gp] !== undefined);' +
  '        var card = document.createElement("div");' +
  '        card.className = "gp-card" + (isSel ? " active" : "");' +
  '        card.id = "card_" + gp;' +
  '        var head = document.createElement("div");' +
  '        head.className = "gp-head";' +
  '        var label = document.createElement("label");' +
  '        label.className = "gp-label";' +
  '        var cb = document.createElement("input");' +
  '        cb.type = "checkbox";' +
  '        cb.checked = isSel;' +
  '        cb.onchange = function(e){ toggleGp(gp, cb.checked); };' +
  '        label.appendChild(cb);' +
  '        label.appendChild(document.createTextNode(" " + gp));' +
  '        head.appendChild(label);' +
  '        card.appendChild(head);' +
  '        var wBox = document.createElement("div");' +
  '        wBox.className = "ward-box";' +
  '        var chips = document.createElement("div");' +
  '        chips.className = "ward-chips";' +
  '        var allBtn = document.createElement("button");' +
  '        allBtn.type = "button";' +
  '        allBtn.className = "ward-btn all-btn";' +
  '        allBtn.textContent = "सारे वार्ड (ALL)";' +
  '        allBtn.onclick = function(){ selectGpAllWards(gp); };' +
  '        chips.appendChild(allBtn);' +
  '        for (var w = 1; w <= 15; w++) {' +
  '          (function(wNum){' +
  '            var btn = document.createElement("button");' +
  '            btn.type = "button";' +
  '            var curVal = selectedMap[gp];' +
  '            var wSel = (Array.isArray(curVal) && curVal.indexOf(String(wNum)) !== -1);' +
  '            btn.className = "ward-btn" + (wSel ? " selected" : "");' +
  '            btn.textContent = wNum;' +
  '            btn.onclick = function(){ toggleWard(gp, String(wNum)); };' +
  '            chips.appendChild(btn);' +
  '          })(w);' +
  '        }' +
  '        wBox.appendChild(chips);' +
  '        card.appendChild(wBox);' +
  '        container.appendChild(card);' +
  '      });' +
  '    }' +
  '    function toggleGp(gp, check) {' +
  '      if (check) { selectedMap[gp] = "ALL"; } else { delete selectedMap[gp]; }' +
  '      var c = document.getElementById("card_" + gp);' +
  '      if (c) { if (check) c.classList.add("active"); else c.classList.remove("active"); }' +
  '      updatePreview();' +
  '    }' +
  '    function selectGpAllWards(gp) {' +
  '      selectedMap[gp] = "ALL";' +
  '      var c = document.getElementById("card_" + gp);' +
  '      if (c) {' +
  '        c.classList.add("active");' +
  '        var btns = c.querySelectorAll(".ward-btn:not(.all-btn)");' +
  '        btns.forEach(function(b){ b.classList.remove("selected"); });' +
  '      }' +
  '      updatePreview();' +
  '    }' +
  '    function toggleWard(gp, wStr) {' +
  '      var cur = selectedMap[gp];' +
  '      var arr = [];' +
  '      if (Array.isArray(cur)) { arr = cur.slice(); }' +
  '      var idx = arr.indexOf(wStr);' +
  '      if (idx !== -1) { arr.splice(idx, 1); } else { arr.push(wStr); }' +
  '      if (arr.length === 0) { selectedMap[gp] = "ALL"; } else { selectedMap[gp] = arr; }' +
  '      var c = document.getElementById("card_" + gp);' +
  '      if (c) {' +
  '        c.classList.add("active");' +
  '        var btns = c.querySelectorAll(".ward-btn:not(.all-btn)");' +
  '        btns.forEach(function(b){' +
  '          if (arr.indexOf(b.textContent) !== -1) b.classList.add("selected");' +
  '          else b.classList.remove("selected");' +
  '        });' +
  '      }' +
  '      updatePreview();' +
  '    }' +
  '    function applyPresetKamala() {' +
  '      selectedMap = { "बांदनवाड़ा": "ALL" };' +
  '      document.getElementById("postSelect").value = "पंचायत समिति सदस्य";' +
  '      document.getElementById("printWardInput").value = "2";' +
  '      renderGpList(); updatePreview();' +
  '    }' +
  '    function applyPresetRajendra() {' +
  '      selectedMap = {' +
  '        "बांदनवाड़ा": ["1", "2", "3"],' +
  '        "देवलियाकलां": ["1", "2", "3", "4"],' +
  '        "कनईकला": ["11", "12", "13"]' +
  '      };' +
  '      document.getElementById("postSelect").value = "जिला परिषद सदस्य";' +
  '      document.getElementById("printWardInput").value = "15";' +
  '      renderGpList(); updatePreview();' +
  '    }' +
  '    function applyPresetSarpanch() {' +
  '      var curGp = Object.keys(selectedMap)[0] || "कनईकला";' +
  '      selectedMap = {};' +
  '      selectedMap[curGp] = "ALL";' +
  '      document.getElementById("postSelect").value = "सरपंच";' +
  '      document.getElementById("printWardInput").value = "ALL";' +
  '      renderGpList(); updatePreview();' +
  '    }' +
  '    function clearAllSelections() { selectedMap = {}; renderGpList(); updatePreview(); }' +
  '    function filterGps(val) { renderGpList(val); }' +
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
  '      if (pWard && pWard !== "ALL" && post !== "सरपंच") { slipText += " प्रत्याशी • वार्ड नं.-" + pWard; }' +
  '      else { slipText += " प्रत्याशी"; }' +
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
  '</body></html>';
}

// ================================================================================================
// 5. MASTER CANDIDATE SETTINGS & PHOTO UPLOAD WEB APPLICATION (HTML)
// ================================================================================================
function buildCandidateSettingsWebHtml() {
  var candRes = handleGetCandidates();
  var candidatesJson = JSON.stringify(candRes.candidates || []);
  var gpsJson = JSON.stringify(OFFICIAL_PANCHAYATS);
  var symbolsJson = JSON.stringify(OFFICIAL_SYMBOLS);

  return '<!DOCTYPE html>\n' +
  '<html lang="hi">\n' +
  '<head>\n' +
  '  <meta charset="UTF-8">\n' +
  '  <base target="_top">\n' +
  '  <meta name="viewport" content="width=device-width, initial-scale=1.0">\n' +
  '  <title>प्रत्याशी प्रबंधन एवं सेटिंग्स कमान कक्ष - पंचायत चुनाव 2026</title>\n' +
  '  <style>\n' +
  '    :root {\n' +
  '      --primary: #1e3a8a; --primary-light: #3b82f6; --success: #059669;\n' +
  '      --danger: #dc2626; --warning: #d97706; --bg: #f8fafc;\n' +
  '      --card-bg: #ffffff; --text: #0f172a; --text-muted: #64748b;\n' +
  '      --border: #cbd5e1;\n' +
  '    }\n' +
  '    * { box-sizing: border-box; margin:0; padding:0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; }\n' +
  '    body { background: var(--bg); color: var(--text); padding: 16px; font-size: 13px; line-height: 1.5; }\n' +
  '    .top-header { background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%); color: #fff; padding: 14px 20px; border-radius: 12px; margin-bottom: 16px; box-shadow: 0 4px 12px rgba(30,58,138,0.2); display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; }\n' +
  '    .top-title { font-size: 1.25rem; font-weight: 800; display:flex; align-items:center; gap:8px; }\n' +
  '    .top-sub { font-size: 0.8rem; opacity: 0.9; margin-top: 3px; }\n' +
  '    .top-actions { display:flex; gap:8px; flex-wrap:wrap; align-items:center; }\n' +
  '    .btn { padding: 7px 14px; border-radius: 6px; font-weight: 700; font-size: 12px; cursor: pointer; border: 1px solid transparent; transition: all 0.15s; display:inline-flex; align-items:center; gap:5px; }\n' +
  '    .btn-primary { background: var(--primary-light); color:#fff; }\n' +
  '    .btn-primary:hover { background: #2563eb; }\n' +
  '    .btn-success { background: var(--success); color:#fff; }\n' +
  '    .btn-success:hover { background: #047857; }\n' +
  '    .btn-warning { background: var(--warning); color:#fff; }\n' +
  '    .btn-danger { background: var(--danger); color:#fff; }\n' +
  '    .btn-outline { background: #fff; border-color: var(--border); color: var(--text); }\n' +
  '    .btn-outline:hover { background: #f1f5f9; }\n' +
  '    .main-grid { display: grid; grid-template-columns: 340px 1fr; gap: 16px; align-items: start; }\n' +
  '    @media(max-width: 900px) { .main-grid { grid-template-columns: 1fr; } }\n' +
  '    .panel-card { background: var(--card-bg); border: 1px solid var(--border); border-radius: 10px; padding: 16px; box-shadow: 0 2px 6px rgba(0,0,0,0.05); }\n' +
  '    .panel-head { font-size: 14px; font-weight: 800; color: #1e3a8a; margin-bottom: 12px; display:flex; justify-content:space-between; align-items:center; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; }\n' +
  '    .cand-search-input { width: 100%; padding: 8px 12px; border: 1px solid var(--border); border-radius: 6px; font-size: 12px; margin-bottom: 10px; outline: none; }\n' +
  '    .cand-search-input:focus { border-color: var(--primary-light); }\n' +
  '    .cand-list { max-height: 680px; overflow-y: auto; display: flex; flex-direction: column; gap: 8px; padding-right: 4px; }\n' +
  '    .cand-item { border: 1px solid var(--border); border-radius: 8px; padding: 10px; background: #fff; cursor: pointer; transition: all 0.15s; display: flex; gap: 10px; align-items: center; position: relative; }\n' +
  '    .cand-item:hover { border-color: var(--primary-light); background: #f0f7ff; }\n' +
  '    .cand-item.active { border-color: #2563eb; background: #eff6ff; box-shadow: 0 0 0 2px rgba(37,99,235,0.2); }\n' +
  '    .cand-thumb { width: 44px; height: 52px; border-radius: 6px; object-fit: cover; background: #e2e8f0; border: 1px solid #cbd5e1; flex-shrink: 0; }\n' +
  '    .cand-item-info { flex: 1; min-width: 0; }\n' +
  '    .cand-item-name { font-weight: 800; font-size: 13px; color: #0f172a; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }\n' +
  '    .cand-item-meta { font-size: 11px; color: var(--text-muted); margin-top: 2px; }\n' +
  '    .cand-badge { display: inline-block; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 4px; margin-top: 3px; }\n' +
  '    .badge-active { background: #dcfce7; color: #15803d; }\n' +
  '    .badge-inactive { background: #fee2e2; color: #b91c1c; }\n' +
  '    .badge-symbol { background: #fef3c7; color: #b45309; font-weight:700; }\n' +
  '    .form-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 12px; margin-bottom: 12px; }\n' +
  '    .form-group { display: flex; flex-direction: column; gap: 4px; }\n' +
  '    .form-group.full-width { grid-column: 1 / -1; }\n' +
  '    .form-label { font-size: 11px; font-weight: 700; color: #334155; text-transform: uppercase; letter-spacing: 0.3px; }\n' +
  '    .form-input, .form-select, .form-textarea { width: 100%; padding: 8px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 13px; font-weight: 600; outline: none; background: #fff; }\n' +
  '    .form-input:focus, .form-select:focus, .form-textarea:focus { border-color: #2563eb; box-shadow: 0 0 0 2px rgba(37,99,235,0.15); }\n' +
  '    .photo-upload-section { background: #f1f5f9; border: 2px dashed #94a3b8; border-radius: 10px; padding: 14px; text-align: center; display: flex; gap: 16px; align-items: center; justify-content: center; flex-wrap: wrap; margin-bottom: 16px; }\n' +
  '    .photo-preview-box { width: 100px; height: 120px; border-radius: 8px; border: 2px solid #3b82f6; overflow: hidden; background: #fff; position: relative; box-shadow: 0 2px 6px rgba(0,0,0,0.1); flex-shrink: 0; }\n' +
  '    .photo-preview-box img { width: 100%; height: 100%; object-fit: cover; display: block; }\n' +
  '    .photo-desc { text-align: left; max-width: 380px; }\n' +
  '    .specimen-card { background: #fff; border: 2px solid #1e3a8a; border-radius: 8px; overflow: hidden; max-width: 440px; margin: 16px auto; box-shadow: 0 4px 14px rgba(0,0,0,0.12); }\n' +
  '    .specimen-top { background: linear-gradient(to bottom, #ffffff, #fffbeb); border-bottom: 2px solid #1e3a8a; padding: 10px 12px; display: flex; gap: 10px; align-items: center; }\n' +
  '    .specimen-photo { width: 62px; height: 75px; border-radius: 4px; border: 1px solid #000; object-fit: cover; background: #f8fafc; flex-shrink: 0; }\n' +
  '    .specimen-info { flex: 1; text-align: center; }\n' +
  '    .specimen-title { font-size: 10px; font-weight: 700; color: #1e3a8a; }\n' +
  '    .specimen-name { font-size: 15px; font-weight: 900; color: #b91c1c; margin: 2px 0; }\n' +
  '    .specimen-post { font-size: 11px; font-weight: 800; background: #1e3a8a; color: #fff; padding: 2px 6px; border-radius: 3px; display: inline-block; }\n' +
  '    .specimen-gp { font-size: 10px; font-weight: 700; color: #000; margin-top: 3px; }\n' +
  '    .specimen-symbol { width: 56px; text-align: center; flex-shrink: 0; }\n' +
  '    .specimen-symbol-box { width: 52px; height: 52px; border: 1.5px solid #000; border-radius: 4px; padding: 2px; background: #fff; display: flex; align-items: center; justify-content: center; }\n' +
  '    .specimen-symbol-name { font-size: 10px; font-weight: 800; color: #000; margin-top: 2px; }\n' +
  '    .specimen-divider { border-top: 1px dashed #64748b; padding: 4px; text-align: center; font-size: 9px; font-weight: 700; color: #475569; background: #f8fafc; }\n' +
  '    .specimen-bottom { padding: 8px 10px; font-size: 10px; color: #334155; line-height: 1.4; }\n' +
  '    .toast { position: fixed; bottom: 20px; right: 20px; padding: 12px 20px; border-radius: 8px; font-weight: 700; font-size: 13px; color: #fff; box-shadow: 0 4px 12px rgba(0,0,0,0.2); z-index: 9999; display: none; }\n' +
  '    .toast.success { background: #059669; }\n' +
  '    .toast.error { background: #dc2626; }\n' +
  '  </style>\n' +
  '</head>\n' +
  '<body>\n' +
  '  <!-- Top Header Bar -->\n' +
  '  <div class="top-header">\n' +
  '    <div>\n' +
  '      <div class="top-title"><span>🗳️ राजस्थान पंचायती राज चुनाव - 2026</span> <span style="font-size:11px; background:rgba(255,255,255,0.2); padding:3px 8px; border-radius:12px;">SUPER ADMIN COMMAND</span></div>\n' +
  '      <div class="top-sub">सुपर एडमिन प्रत्याशी प्रबंधन, रंगीन फोटो अपलोड एवं एकल मास्टर शीट सिंक कमान कक्ष</div>\n' +
  '    </div>\n' +
  '    <div class="top-actions">\n' +
  '      <button type="button" class="btn btn-warning" onclick="createNewCandidate()">➕ नया प्रत्याशी जोड़ें</button>\n' +
  '      <button type="button" class="btn btn-primary" onclick="refreshCandidatesList()">🔄 रीफ्रेश सूची</button>\n' +
  '      <button type="button" class="btn btn-success" onclick="exportPortalJson()">📥 डाउनलोड portal_users.json</button>\n' +
  '      <button type="button" class="btn btn-outline" onclick="copyPortalJson()">📋 कॉपी JSON</button>\n' +
  '    </div>\n' +
  '  </div>\n' +
  '\n' +
  '  <div class="main-grid">\n' +
  '    <!-- Left Column: Candidates List -->\n' +
  '    <div class="panel-card">\n' +
  '      <div class="panel-head">\n' +
  '        <span>👤 पंजीकृत प्रत्याशी (<span id="candCountText">0</span>)</span>\n' +
  '        <button type="button" class="btn btn-xs btn-outline" onclick="createNewCandidate()" style="padding:2px 8px; font-size:11px;">+ नया</button>\n' +
  '      </div>\n' +
  '      <input type="text" id="candSearchInput" class="cand-search-input" placeholder="🔍 प्रत्याशी, पद या पंचायत खोजें..." oninput="filterCandidates(this.value)">\n' +
  '      <div class="cand-list" id="candidatesListContainer">\n' +
  '        <!-- Candidate Cards rendered here -->\n' +
  '      </div>\n' +
  '    </div>\n' +
  '\n' +
  '    <!-- Right Column: Candidate Editor & Live Specimen -->\n' +
  '    <div class="panel-card">\n' +
  '      <div class="panel-head">\n' +
  '        <span id="editorTitleText">✏️ प्रत्याशी विवरण एवं सेटिंग्स संपादन</span>\n' +
  '        <div>\n' +
  '          <button type="button" class="btn btn-danger btn-xs" id="deleteCandBtn" onclick="deleteCurrentCandidate()" style="display:none; padding:3px 10px;">🗑️ प्रत्याशी हटाएं</button>\n' +
  '        </div>\n' +
  '      </div>\n' +
  '\n' +
  '      <!-- Photo Upload Box (Key Requirement) -->\n' +
  '      <div class="photo-upload-section">\n' +
  '        <div class="photo-preview-box" onclick="document.getElementById(\'photoFileInput\').click()" title="फोटो बदलने हेतु क्लिक करें">\n' +
  '          <img id="photoPreviewImg" src="https://api.dicebear.com/7.x/identicon/svg?seed=candidate" alt="Photo">\n' +
  '        </div>\n' +
  '        <div class="photo-desc">\n' +
  '          <label style="font-weight:800; font-size:12px; color:#1e3a8a; display:block; margin-bottom:4px;">📸 प्रत्याशी की रंगीन फोटो (Color Photo):</label>\n' +
  '          <p style="font-size:11px; color:#64748b; margin-bottom:8px;">मतदाता पर्ची के शीर्ष प्रचार बैनर पर यह फोटो पूरी तरह रंगीन मुद्रित होगी। फोटो का आकार स्वतः ऑप्टिमाइज़ होकर गूगल शीट में सुरक्षित होगा।</p>\n' +
  '          <div style="display:flex; gap:8px; align-items:center;">\n' +
  '            <label for="photoFileInput" class="btn btn-primary" style="cursor:pointer; margin:0;">📁 फोटो चुनें / बदलें</label>\n' +
  '            <input type="file" id="photoFileInput" accept="image/*" style="display:none;" onchange="onPhotoFileSelected(event)">\n' +
  '            <button type="button" id="removePhotoBtn" class="btn btn-outline" onclick="removePhoto()" style="display:none;">❌ हटाएं</button>\n' +
  '          </div>\n' +
  '        </div>\n' +
  '      </div>\n' +
  '\n' +
  '      <!-- Form Details Grid -->\n' +
  '      <form id="candidateForm" onsubmit="handleFormSubmit(event)">\n' +
  '        <div class="form-grid">\n' +
  '          <div class="form-group">\n' +
  '            <label class="form-label" for="formUsername">यूजरनेम (Username) *:</label>\n' +
  '            <input type="text" id="formUsername" class="form-input" required placeholder="उदा. kamla / sarpanch01" oninput="updateLiveSpecimen()">\n' +
  '          </div>\n' +
  '          <div class="form-group">\n' +
  '            <label class="form-label" for="formPassword">पासवर्ड (Password) *:</label>\n' +
  '            <input type="text" id="formPassword" class="form-input" required value="bhinai2026">\n' +
  '          </div>\n' +
  '          <div class="form-group">\n' +
  '            <label class="form-label" for="formCandName">प्रत्याशी का पूरा नाम (Full Name) *:</label>\n' +
  '            <input type="text" id="formCandName" class="form-input" required placeholder="उदा. श्रीमती कमला देवी" oninput="updateLiveSpecimen()" style="color:#b91c1c; font-weight:800;">\n' +
  '          </div>\n' +
  '          <div class="form-group">\n' +
  '            <label class="form-label" for="formMobile">सम्पर्क / मोबाइल नं.:</label>\n' +
  '            <input type="tel" id="formMobile" class="form-input" placeholder="उदा. 9876543210" maxlength="10">\n' +
  '          </div>\n' +
  '        </div>\n' +
  '\n' +
  '        <div class="form-grid">\n' +
  '          <div class="form-group">\n' +
  '            <label class="form-label" for="formPost">चुनाव पद (Contesting Post) *:</label>\n' +
  '            <select id="formPost" class="form-select" onchange="onPostChanged(this.value)">\n' +
  '              <option value="सरपंच">सरपंच पद हेतु</option>\n' +
  '              <option value="पंचायत समिति सदस्य">पंचायत समिति सदस्य हेतु</option>\n' +
  '              <option value="जिला परिषद सदस्य">जिला परिषद सदस्य हेतु</option>\n' +
  '              <option value="वार्ड पंच">वार्ड पंच पद हेतु</option>\n' +
  '              <option value="उपसरपंच">उपसरपंच पद हेतु</option>\n' +
  '            </select>\n' +
  '          </div>\n' +
  '          <div class="form-group">\n' +
  '            <label class="form-label" for="formGp">मुख्य ग्राम पंचायत (Gram Panchayat) *:</label>\n' +
  '            <select id="formGp" class="form-select" onchange="updateLiveSpecimen()"></select>\n' +
  '          </div>\n' +
  '          <div class="form-group">\n' +
  '            <label class="form-label" for="formPrintWard">पर्ची पर छपने वाला वार्ड नं.:</label>\n' +
  '            <input type="text" id="formPrintWard" class="form-input" placeholder="उदा. 2 (सरपंच हेतु खाली/ALL)" oninput="updateLiveSpecimen()">\n' +
  '          </div>\n' +
  '          <div class="form-group">\n' +
  '            <label class="form-label" for="formSymbolSelect">चुनाव चिन्ह (Election Symbol) *:</label>\n' +
  '            <select id="formSymbolSelect" class="form-select" onchange="onSymbolChanged(this.value)"></select>\n' +
  '          </div>\n' +
  '        </div>\n' +
  '\n' +
  '        <div class="form-grid">\n' +
  '          <div class="form-group full-width">\n' +
  '            <label class="form-label" for="formAllowedGps">अनुमति प्राप्त पंचायतें (Allowed Panchayats):</label>\n' +
  '            <input type="text" id="formAllowedGps" class="form-input" placeholder="उदा. बांदनवाड़ा या बांदनवाड़ा, देवलियाकलां या ALL">\n' +
  '          </div>\n' +
  '          <div class="form-group full-width">\n' +
  '            <label class="form-label" for="formAllowedWards">आंशिक वार्ड दायरा (Allowed Panchayat Wards):</label>\n' +
  '            <input type="text" id="formAllowedWards" class="form-input" placeholder="उदा. बांदनवाड़ा: ALL या बांदनवाड़ा: 1,2,3 | देवलियाकलां: 1,2,3,4">\n' +
  '          </div>\n' +
  '        </div>\n' +
  '\n' +
  '        <div class="form-grid">\n' +
  '          <div class="form-group full-width">\n' +
  '            <label class="form-label" for="formSlogan">चुनाव प्रचार अपील / नारा (Campaign Slogan):</label>\n' +
  '            <textarea id="formSlogan" class="form-textarea" rows="2" placeholder="क्षेत्र के सर्वांगीण विकास एवं खुशहाली हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।" oninput="updateLiveSpecimen()"></textarea>\n' +
  '          </div>\n' +
  '        </div>\n' +
  '\n' +
  '        <div class="form-grid">\n' +
  '          <div class="form-group">\n' +
  '            <label class="form-label" for="formDate">मतदान दिनांक (Election Date):</label>\n' +
  '            <input type="text" id="formDate" class="form-input" value="15 अक्टूबर 2026" oninput="updateLiveSpecimen()">\n' +
  '          </div>\n' +
  '          <div class="form-group">\n' +
  '            <label class="form-label" for="formTime">मतदान समय (Voting Time):</label>\n' +
  '            <input type="text" id="formTime" class="form-input" value="प्रातः 7:00 बजे से सायं 5:00 बजे तक" oninput="updateLiveSpecimen()">\n' +
  '          </div>\n' +
  '          <div class="form-group" style="justify-content:center;">\n' +
  '            <label style="display:flex; align-items:center; gap:6px; cursor:pointer; font-weight:700; margin-top:14px;">\n' +
  '              <input type="checkbox" id="formStatus" checked style="width:16px; height:16px;"> 🟢 खाता सक्रिय रखें (ACTIVE)\n' +
  '            </label>\n' +
  '          </div>\n' +
  '        </div>\n' +
  '\n' +
  '        <!-- Live Perforated Slip Specimen Card -->\n' +
  '        <div style="margin-top:16px; background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:12px;">\n' +
  '          <div style="font-weight:800; font-size:12px; color:#334155; margin-bottom:6px; text-align:center;">📋 लाइव पर्ची प्रारूप नमूना (Live Voter Slip Specimen)</div>\n' +
  '          <div class="specimen-card" id="liveSpecimenCard">\n' +
  '            <div class="specimen-top">\n' +
  '              <img id="specPhoto" class="specimen-photo" src="https://api.dicebear.com/7.x/identicon/svg?seed=candidate" alt="Photo">\n' +
  '              <div class="specimen-info">\n' +
  '                <div class="specimen-title">।। श्री गणेशाय नमः ।। 🙏 विनम्र अपील</div>\n' +
  '                <div class="specimen-name" id="specName">कमला देवी</div>\n' +
  '                <div><span class="specimen-post" id="specPost">पंचायत समिति सदस्य प्रत्याशी</span></div>\n' +
  '                <div class="specimen-gp" id="specGpRow">ग्राम पंचायत: बांदनवाड़ा • वार्ड नं.: 2</div>\n' +
  '                <div style="font-size:9px; color:#475569; margin-top:2px;">📅 मतदान: <strong id="specDate">15 अक्टूबर 2026</strong></div>\n' +
  '              </div>\n' +
  '              <div class="specimen-symbol">\n' +
  '                <div class="specimen-symbol-box" id="specSymbolBox">✋</div>\n' +
  '                <div class="specimen-symbol-name" id="specSymbolName">हाथ</div>\n' +
  '              </div>\n' +
  '            </div>\n' +
  '            <div class="specimen-divider">✂️ --- यहाँ से काटकर मतदाता को दें --- ✂️</div>\n' +
  '            <div class="specimen-bottom">\n' +
  '              <div style="display:flex; justify-content:space-between; font-weight:700;"><span>मतदाता: रमेश कुमार</span> <span>क्रम: 104</span></div>\n' +
  '              <div style="font-size:9px; color:#64748b;">मतदान केंद्र: रा.उ.मा.वि. कक्ष-1</div>\n' +
  '            </div>\n' +
  '          </div>\n' +
  '        </div>\n' +
  '\n' +
  '        <!-- Action Buttons -->\n' +
  '        <div style="display:flex; gap:10px; margin-top:16px; flex-wrap:wrap;">\n' +
  '          <button type="submit" id="saveSubmitBtn" class="btn btn-success" style="flex:1; padding:12px; font-size:14px; justify-content:center;">\n' +
  '            💾 गूगल शीट में सुरक्षित करें (Save to Google Sheet)\n' +
  '          </button>\n' +
  '          <button type="button" class="btn btn-outline" onclick="createNewCandidate()" style="padding:12px 18px;">\n' +
  '            ➕ नया प्रत्याशी\n' +
  '          </button>\n' +
  '        </div>\n' +
  '      </form>\n' +
  '    </div>\n' +
  '  </div>\n' +
  '\n' +
  '  <!-- Toast notification -->\n' +
  '  <div id="toastBox" class="toast"></div>\n' +
  '\n' +
  '  <script>\n' +
  '    var candidates = ' + candidatesJson + ';\n' +
  '    var officialGps = ' + gpsJson + ';\n' +
  '    var officialSymbols = ' + symbolsJson + ';\n' +
  '    var activeCandidate = null;\n' +
  '    var currentPhotoBase64 = "";\n' +
  '\n' +
  '    initApp();\n' +
  '\n' +
  '    function initApp() {\n' +
  '      populateDropdowns();\n' +
  '      renderCandidatesList();\n' +
  '      if (candidates.length > 0) {\n' +
  '        selectCandidate(candidates[0].username);\n' +
  '      } else {\n' +
  '        createNewCandidate();\n' +
  '      }\n' +
  '    }\n' +
  '\n' +
  '    function populateDropdowns() {\n' +
  '      var gpSelect = document.getElementById("formGp");\n' +
  '      gpSelect.innerHTML = "";\n' +
  '      officialGps.forEach(function(gp) {\n' +
  '        var opt = document.createElement("option");\n' +
  '        opt.value = gp; opt.textContent = gp;\n' +
  '        gpSelect.appendChild(opt);\n' +
  '      });\n' +
  '\n' +
  '      var symSelect = document.getElementById("formSymbolSelect");\n' +
  '      symSelect.innerHTML = "";\n' +
  '      officialSymbols.forEach(function(sym) {\n' +
  '        var opt = document.createElement("option");\n' +
  '        opt.value = sym; opt.textContent = sym;\n' +
  '        symSelect.appendChild(opt);\n' +
  '      });\n' +
  '    }\n' +
  '\n' +
  '    function renderCandidatesList(filterText) {\n' +
  '      var container = document.getElementById("candidatesListContainer");\n' +
  '      container.innerHTML = "";\n' +
  '      var q = (filterText || "").trim().toLowerCase();\n' +
  '      var count = 0;\n' +
  '\n' +
  '      candidates.forEach(function(c) {\n' +
  '        var text = (c.candidate_name + " " + c.post + " " + c.panchayat + " " + c.username).toLowerCase();\n' +
  '        if (q && text.indexOf(q) === -1) return;\n' +
  '        count++;\n' +
  '\n' +
  '        var item = document.createElement("div");\n' +
  '        var isActive = activeCandidate && (activeCandidate.username === c.username);\n' +
  '        item.className = "cand-item" + (isActive ? " active" : "");\n' +
  '        item.onclick = function() { selectCandidate(c.username); };\n' +
  '\n' +
  '        var photoSrc = c.photo_url || (c.username === "kamla" || c.username === "kamala" ? "./assets/kamala_photo.jpg" : "https://api.dicebear.com/7.x/identicon/svg?seed=" + encodeURIComponent(c.username));\n' +
  '\n' +
  '        item.innerHTML = \n' +
  '          \'<img class="cand-thumb" src="\' + photoSrc + \'" alt="\' + c.candidate_name + \'" onerror="this.src=\\\'https://api.dicebear.com/7.x/identicon/svg?seed=cand\\\'">\' +\n' +
  '          \'<div class="cand-item-info">\' +\n' +
  '          \'  <div class="cand-item-name">\' + (c.candidate_name || c.username) + \'</div>\' +\n' +
  '          \'  <div class="cand-item-meta">\' + (c.post || "") + \' • \' + (c.panchayat || "") + \'</div>\' +\n' +
  '          \'  <div>\' +\n' +
  '          \'    <span class="cand-badge \' + (c.status === "ACTIVE" ? "badge-active" : "badge-inactive") + \'">\' + (c.status === "ACTIVE" ? "🟢 एक्टिव" : "🔴 बंद") + \'</span> \' +\n' +
  '          \'    <span class="cand-badge badge-symbol">🚩 \' + (c.symbol_name || "चिन्ह") + \'</span>\' +\n' +
  '          \'  </div>\' +\n' +
  '          \'</div>\';\n' +
  '\n' +
  '        container.appendChild(item);\n' +
  '      });\n' +
  '\n' +
  '      document.getElementById("candCountText").textContent = count;\n' +
  '    }\n' +
  '\n' +
  '    function filterCandidates(q) {\n' +
  '      renderCandidatesList(q);\n' +
  '    }\n' +
  '\n' +
  '    function selectCandidate(uName) {\n' +
  '      var c = candidates.find(function(x) { return x.username === uName; });\n' +
  '      if (!c) return;\n' +
  '      activeCandidate = c;\n' +
  '      currentPhotoBase64 = c.photo_url || "";\n' +
  '\n' +
  '      document.getElementById("editorTitleText").textContent = "✏️ " + (c.candidate_name || c.username) + " का विवरण संपादन";\n' +
  '      document.getElementById("deleteCandBtn").style.display = (c.username === "admin" ? "none" : "inline-block");\n' +
  '\n' +
  '      document.getElementById("formUsername").value = c.username;\n' +
  '      document.getElementById("formUsername").readOnly = (c.username === "admin" || c.username === "kamala" || c.username === "kamla");\n' +
  '      document.getElementById("formPassword").value = c.password || "bhinai2026";\n' +
  '      document.getElementById("formCandName").value = c.candidate_name || "";\n' +
  '      document.getElementById("formMobile").value = c.mobile || "";\n' +
  '      document.getElementById("formPost").value = c.post || "सरपंच";\n' +
  '      document.getElementById("formGp").value = c.panchayat || officialGps[0];\n' +
  '      document.getElementById("formPrintWard").value = c.print_ward_no || c.ward || "";\n' +
  '      document.getElementById("formSymbolSelect").value = c.symbol_name || "हाथ";\n' +
  '      document.getElementById("formAllowedGps").value = c.allowed_panchayats || c.panchayat || "ALL";\n' +
  '      document.getElementById("formAllowedWards").value = c.allowed_panchayat_wards || c.allowed_wards || "ALL";\n' +
  '      document.getElementById("formSlogan").value = c.slogan || "";\n' +
  '      document.getElementById("formDate").value = c.election_date || "15 अक्टूबर 2026";\n' +
  '      document.getElementById("formTime").value = c.election_time || "प्रातः 7:00 बजे से सायं 5:00 बजे तक";\n' +
  '      document.getElementById("formStatus").checked = (c.status === "ACTIVE");\n' +
  '\n' +
  '      var previewImg = document.getElementById("photoPreviewImg");\n' +
  '      if (c.photo_url) {\n' +
  '        previewImg.src = c.photo_url;\n' +
  '        document.getElementById("removePhotoBtn").style.display = "inline-block";\n' +
  '      } else if (c.username === "kamla" || c.username === "kamala") {\n' +
  '        previewImg.src = "./assets/kamala_photo.jpg";\n' +
  '        document.getElementById("removePhotoBtn").style.display = "inline-block";\n' +
  '      } else {\n' +
  '        previewImg.src = "https://api.dicebear.com/7.x/identicon/svg?seed=" + encodeURIComponent(c.username);\n' +
  '        document.getElementById("removePhotoBtn").style.display = "none";\n' +
  '      }\n' +
  '\n' +
  '      renderCandidatesList(document.getElementById("candSearchInput").value);\n' +
  '      updateLiveSpecimen();\n' +
  '    }\n' +
  '\n' +
  '    function createNewCandidate() {\n' +
  '      activeCandidate = null;\n' +
  '      currentPhotoBase64 = "";\n' +
  '      document.getElementById("editorTitleText").textContent = "➕ नया प्रत्याशी जोड़ें";\n' +
  '      document.getElementById("deleteCandBtn").style.display = "none";\n' +
  '\n' +
  '      document.getElementById("formUsername").value = "cand_" + Math.floor(100 + Math.random() * 900);\n' +
  '      document.getElementById("formUsername").readOnly = false;\n' +
  '      document.getElementById("formPassword").value = "bhinai2026";\n' +
  '      document.getElementById("formCandName").value = "";\n' +
  '      document.getElementById("formMobile").value = "";\n' +
  '      document.getElementById("formPost").value = "सरपंच";\n' +
  '      document.getElementById("formGp").value = officialGps[0];\n' +
  '      document.getElementById("formPrintWard").value = "";\n' +
  '      document.getElementById("formSymbolSelect").value = "उगता सूरज";\n' +
  '      document.getElementById("formAllowedGps").value = officialGps[0];\n' +
  '      document.getElementById("formAllowedWards").value = "ALL";\n' +
  '      document.getElementById("formSlogan").value = "क्षेत्र के सर्वांगीण विकास एवं खुशहाली हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।";\n' +
  '      document.getElementById("formDate").value = "15 अक्टूबर 2026";\n' +
  '      document.getElementById("formTime").value = "प्रातः 7:00 बजे से सायं 5:00 बजे तक";\n' +
  '      document.getElementById("formStatus").checked = true;\n' +
  '\n' +
  '      document.getElementById("photoPreviewImg").src = "https://api.dicebear.com/7.x/identicon/svg?seed=new";\n' +
  '      document.getElementById("removePhotoBtn").style.display = "none";\n' +
  '\n' +
  '      renderCandidatesList();\n' +
  '      updateLiveSpecimen();\n' +
  '    }\n' +
  '\n' +
  '    function onPostChanged(post) {\n' +
  '      var wardInput = document.getElementById("formPrintWard");\n' +
  '      if (post === "पंचायत समिति सदस्य") {\n' +
  '        if (!wardInput.value || wardInput.value === "ALL") wardInput.value = "2";\n' +
  '      } else if (post === "जिला परिषद सदस्य") {\n' +
  '        if (!wardInput.value || wardInput.value === "ALL") wardInput.value = "15";\n' +
  '      } else if (post === "सरपंच") {\n' +
  '        wardInput.value = "ALL";\n' +
  '      }\n' +
  '      updateLiveSpecimen();\n' +
  '    }\n' +
  '\n' +
  '    function onSymbolChanged(sym) {\n' +
  '      updateLiveSpecimen();\n' +
  '    }\n' +
  '\n' +
  '    // Photo Upload Handler with Client-Side Canvas Compression\n' +
  '    function onPhotoFileSelected(e) {\n' +
  '      var file = e.target.files && e.target.files[0];\n' +
  '      if (!file) return;\n' +
  '      showToast("⏳ फोटो प्रोसेस हो रही है...", "success");\n' +
  '\n' +
  '      var reader = new FileReader();\n' +
  '      reader.onload = function(evt) {\n' +
  '        var img = new Image();\n' +
  '        img.onload = function() {\n' +
  '          var maxW = 360, maxH = 450;\n' +
  '          var w = img.width, h = img.height;\n' +
  '          if (w > maxW || h > maxH) {\n' +
  '            if (w/h > maxW/maxH) {\n' +
  '              h = Math.round(h * (maxW / w));\n' +
  '              w = maxW;\n' +
  '            } else {\n' +
  '              w = Math.round(w * (maxH / h));\n' +
  '              h = maxH;\n' +
  '            }\n' +
  '          }\n' +
  '          var canvas = document.createElement("canvas");\n' +
  '          canvas.width = w;\n' +
  '          canvas.height = h;\n' +
  '          var ctx = canvas.getContext("2d");\n' +
  '          ctx.drawImage(img, 0, 0, w, h);\n' +
  '          var base64 = canvas.toDataURL("image/jpeg", 0.82);\n' +
  '          currentPhotoBase64 = base64;\n' +
  '          document.getElementById("photoPreviewImg").src = base64;\n' +
  '          document.getElementById("removePhotoBtn").style.display = "inline-block";\n' +
  '          updateLiveSpecimen();\n' +
  '          showToast("✅ फोटो तैयार! सुरक्षित करने हेतु नीचे \\\'सुरक्षित करें\\\' दबाएं।", "success");\n' +
  '        };\n' +
  '        img.src = evt.target.result;\n' +
  '      };\n' +
  '      reader.readAsDataURL(file);\n' +
  '    }\n' +
  '\n' +
  '    function removePhoto() {\n' +
  '      currentPhotoBase64 = "";\n' +
  '      document.getElementById("photoFileInput").value = "";\n' +
  '      document.getElementById("photoPreviewImg").src = "https://api.dicebear.com/7.x/identicon/svg?seed=" + encodeURIComponent(document.getElementById("formUsername").value);\n' +
  '      document.getElementById("removePhotoBtn").style.display = "none";\n' +
  '      updateLiveSpecimen();\n' +
  '    }\n' +
  '\n' +
  '    function updateLiveSpecimen() {\n' +
  '      var name = document.getElementById("formCandName").value.trim() || document.getElementById("formUsername").value;\n' +
  '      var post = document.getElementById("formPost").value;\n' +
  '      var gp = document.getElementById("formGp").value;\n' +
  '      var ward = document.getElementById("formPrintWard").value.trim();\n' +
  '      var sym = document.getElementById("formSymbolSelect").value;\n' +
  '      var date = document.getElementById("formDate").value;\n' +
  '\n' +
  '      document.getElementById("specName").textContent = name;\n' +
  '      document.getElementById("specPost").textContent = post + " प्रत्याशी";\n' +
  '\n' +
  '      var gpRowText = "ग्राम पंचायत: " + gp;\n' +
  '      if (ward && ward !== "ALL" && post !== "सरपंच") {\n' +
  '        gpRowText += " • वार्ड नं.: " + ward;\n' +
  '      }\n' +
  '      document.getElementById("specGpRow").textContent = gpRowText;\n' +
  '      document.getElementById("specDate").textContent = date;\n' +
  '      document.getElementById("specSymbolName").textContent = sym;\n' +
  '\n' +
  '      // Symbol Icon\n' +
  '      var symIconMap = { "हाथ": "✋", "कमल": "🪷", "उगता सूरज": "☀️", "पतंग": "🪁", "बोतल": "🍾", "साइकिल": "🚲", "गिलास": "🥛", "कार": "🚗", "ताला-चाबी": "🔐", "बल्ला": "🏏", "पंखा": "🌀", "सिलेंडर": "🛢️", "मोमबत्ती": "🕯️" };\n' +
  '      document.getElementById("specSymbolBox").textContent = symIconMap[sym] || "🚩";\n' +
  '\n' +
  '      // Photo\n' +
  '      var previewImg = document.getElementById("photoPreviewImg");\n' +
  '      document.getElementById("specPhoto").src = previewImg.src;\n' +
  '    }\n' +
  '\n' +
  '    function handleFormSubmit(e) {\n' +
  '      e.preventDefault();\n' +
  '      var btn = document.getElementById("saveSubmitBtn");\n' +
  '      btn.disabled = true; btn.textContent = "⏳ गूगल शीट में सुरक्षित हो रहा है...";\n' +
  '\n' +
  '      var candData = {\n' +
  '        username: document.getElementById("formUsername").value.trim(),\n' +
  '        password: document.getElementById("formPassword").value.trim(),\n' +
  '        candidate_name: document.getElementById("formCandName").value.trim(),\n' +
  '        mobile: document.getElementById("formMobile").value.trim(),\n' +
  '        post: document.getElementById("formPost").value,\n' +
  '        panchayat: document.getElementById("formGp").value,\n' +
  '        allowed_panchayats: document.getElementById("formAllowedGps").value.trim(),\n' +
  '        print_ward_no: document.getElementById("formPrintWard").value.trim(),\n' +
  '        ward: document.getElementById("formPrintWard").value.trim(),\n' +
  '        allowed_panchayat_wards: document.getElementById("formAllowedWards").value.trim(),\n' +
  '        symbol_name: document.getElementById("formSymbolSelect").value,\n' +
  '        photo_url: currentPhotoBase64,\n' +
  '        slogan: document.getElementById("formSlogan").value.trim(),\n' +
  '        election_date: document.getElementById("formDate").value.trim(),\n' +
  '        election_time: document.getElementById("formTime").value.trim(),\n' +
  '        status: document.getElementById("formStatus").checked ? "ACTIVE" : "INACTIVE"\n' +
  '      };\n' +
  '\n' +
  '      function onSaveSuccess(res) {\n' +
  '        btn.disabled = false; btn.textContent = "💾 गूगल शीट में सुरक्षित करें (Save to Google Sheet)";\n' +
  '        if (res && res.success) {\n' +
  '          showToast("✅ " + res.message, "success");\n' +
  '          if (res.candidates) candidates = res.candidates;\n' +
  '          else refreshCandidatesList();\n' +
  '          selectCandidate(candData.username);\n' +
  '        } else {\n' +
  '          showToast("⚠️ त्रुटि: " + (res.error || "सुरक्षित नहीं हो सका"), "error");\n' +
  '        }\n' +
  '      }\n' +
  '\n' +
  '      function onSaveError(err) {\n' +
  '        btn.disabled = false; btn.textContent = "💾 गूगल शीट में सुरक्षित करें (Save to Google Sheet)";\n' +
  '        showToast("⚠️ नेटवर्क विफलता: " + err, "error");\n' +
  '      }\n' +
  '\n' +
  '      if (typeof google !== "undefined" && google.script && google.script.run) {\n' +
  '        google.script.run\n' +
  '          .withSuccessHandler(onSaveSuccess)\n' +
  '          .withFailureHandler(onSaveError)\n' +
  '          .apiSaveCandidate(candData);\n' +
  '      } else {\n' +
  '        fetch(window.location.href, {\n' +
  '          method: "POST",\n' +
  '          headers: { "Content-Type": "application/json" },\n' +
  '          body: JSON.stringify({ action: "saveCandidate", candidateData: candData })\n' +
  '        })\n' +
  '        .then(function(r) { return r.json(); })\n' +
  '        .then(onSaveSuccess)\n' +
  '        .catch(onSaveError);\n' +
  '      }\n' +
  '    }\n' +
  '\n' +
  '    function deleteCurrentCandidate() {\n' +
  '      if (!activeCandidate) return;\n' +
  '      var u = activeCandidate.username;\n' +
  '      if (!confirm("क्या आप वाकई प्रत्याशी \'" + activeCandidate.candidate_name + " (" + u + ")\' को हटाना चाहते हैं?")) return;\n' +
  '\n' +
  '      showToast("⏳ हटाया जा रहा है...", "success");\n' +
  '      function onDelSuccess(res) {\n' +
  '        if (res && res.success) {\n' +
  '          showToast("✅ " + res.message, "success");\n' +
  '          candidates = candidates.filter(function(x) { return x.username !== u; });\n' +
  '          if (candidates.length > 0) selectCandidate(candidates[0].username);\n' +
  '          else createNewCandidate();\n' +
  '        } else {\n' +
  '          showToast("⚠️ त्रुटि: " + (res.error || "हट नहीं सका"), "error");\n' +
  '        }\n' +
  '      }\n' +
  '\n' +
  '      if (typeof google !== "undefined" && google.script && google.script.run) {\n' +
  '        google.script.run.withSuccessHandler(onDelSuccess).apiDeleteCandidate(u);\n' +
  '      } else {\n' +
  '        fetch(window.location.href, {\n' +
  '          method: "POST",\n' +
  '          headers: { "Content-Type": "application/json" },\n' +
  '          body: JSON.stringify({ action: "deleteCandidate", username: u })\n' +
  '        })\n' +
  '        .then(function(r) { return r.json(); })\n' +
  '        .then(onDelSuccess);\n' +
  '      }\n' +
  '    }\n' +
  '\n' +
  '    function refreshCandidatesList() {\n' +
  '      showToast("🔄 गूगल शीट से प्रत्याशी लोड हो रहे हैं...", "success");\n' +
  '      function onListSuccess(res) {\n' +
  '        if (res && res.candidates) {\n' +
  '          candidates = res.candidates;\n' +
  '          renderCandidatesList();\n' +
  '          showToast("✅ सूची रीफ्रेश हो गई (" + candidates.length + " प्रत्याशी)", "success");\n' +
  '        }\n' +
  '      }\n' +
  '      if (typeof google !== "undefined" && google.script && google.script.run) {\n' +
  '        google.script.run.withSuccessHandler(onListSuccess).apiGetCandidates();\n' +
  '      } else {\n' +
  '        fetch(window.location.href + (window.location.href.indexOf("?") === -1 ? "?" : "&") + "action=getCandidates")\n' +
  '          .then(function(r) { return r.json(); })\n' +
  '          .then(onListSuccess);\n' +
  '      }\n' +
  '    }\n' +
  '\n' +
  '    function exportPortalJson() {\n' +
  '      function onJsonSuccess(res) {\n' +
  '        var data = res.portal_users || res;\n' +
  '        var blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });\n' +
  '        var a = document.createElement("a");\n' +
  '        a.href = URL.createObjectURL(blob);\n' +
  '        a.download = "portal_users.json";\n' +
  '        a.click();\n' +
  '        showToast("✅ portal_users.json डाउनलोड हो गया!", "success");\n' +
  '      }\n' +
  '      if (typeof google !== "undefined" && google.script && google.script.run) {\n' +
  '        google.script.run.withSuccessHandler(onJsonSuccess).apiGetPortalJson();\n' +
  '      } else {\n' +
  '        fetch(window.location.href + (window.location.href.indexOf("?") === -1 ? "?" : "&") + "action=getPortalJson")\n' +
  '          .then(function(r) { return r.json(); })\n' +
  '          .then(onJsonSuccess);\n' +
  '      }\n' +
  '    }\n' +
  '\n' +
  '    function copyPortalJson() {\n' +
  '      function onJsonSuccess(res) {\n' +
  '        var data = res.portal_users || res;\n' +
  '        var str = JSON.stringify(data, null, 2);\n' +
  '        navigator.clipboard.writeText(str).then(function() {\n' +
  '          showToast("📋 portal_users.json क्लिपबोर्ड में कॉपी हो गया!", "success");\n' +
  '        });\n' +
  '      }\n' +
  '      if (typeof google !== "undefined" && google.script && google.script.run) {\n' +
  '        google.script.run.withSuccessHandler(onJsonSuccess).apiGetPortalJson();\n' +
  '      } else {\n' +
  '        fetch(window.location.href + (window.location.href.indexOf("?") === -1 ? "?" : "&") + "action=getPortalJson")\n' +
  '          .then(function(r) { return r.json(); })\n' +
  '          .then(onJsonSuccess);\n' +
  '      }\n' +
  '    }\n' +
  '\n' +
  '    function showToast(msg, type) {\n' +
  '      var t = document.getElementById("toastBox");\n' +
  '      t.textContent = msg;\n' +
  '      t.className = "toast " + (type || "success");\n' +
  '      t.style.display = "block";\n' +
  '      clearTimeout(window._toastTimer);\n' +
  '      window._toastTimer = setTimeout(function() { t.style.display = "none"; }, 4000);\n' +
  '    }\n' +
  '  </script>\n' +
  '</body>\n' +
  '</html>';
}
