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
  var required = ["Custom_Symbol_Url", "Photo_Url", "Allow_Voter_Photo"];
  
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
    "Allow_Voter_Photo", "Role", "Status", "Can_Search", "Can_Alpha", "Can_Bulk_Slips",
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
      "", "", "", "पारदर्शी एवं निष्पक्ष निर्वाचन 2026", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      true, "SUPER_ADMIN", true, true, true, true, true, true, true, true, nowStr
    ],

    // 2. पंचायत समिति सदस्य: कमला देवी (केवल बांदनवाड़ा पंचायत, PS वार्ड 2, हाथ चिन्ह)
    [
      "kamala", "bhinai2026", "कमला देवी", "9876543210", "पंचायत समिति सदस्य",
      "2", "बांदनवाड़ा", "बांदनवाड़ा: ALL",
      "हाथ", "", "./assets/kamala_photo.jpg", "क्षेत्र के सर्वांगीण विकास एवं खुशहाली हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      false, "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ],
    [
      "kamla", "bhinai2026", "कमला देवी", "9876543210", "पंचायत समिति सदस्य",
      "2", "बांदनवाड़ा", "बांदनवाड़ा: ALL",
      "हाथ", "", "./assets/kamala_photo.jpg", "क्षेत्र के सर्वांगीण विकास एवं खुशहाली हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      false, "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ],

    // 3. जिला परिषद सदस्य: राजेन्द्र चौधरी (ZP वार्ड 15 छपेगा, 3 पंचायतों के आंशिक वार्ड खुलेंगे)
    [
      "zp_rajendra", "zp2026", "राजेन्द्र चौधरी", "9829000015", "जिला परिषद सदस्य",
      "15", "बांदनवाड़ा, देवलियाकलां, कनईकला", "बांदनवाड़ा: 1,2,3 | देवलियाकलां: 1,2,3,4 | कनईकला: 11,12,13",
      "कमल", "", "", "जिला परिषद क्षेत्र के चहुंमुखी विकास हेतु आपका अपना कर्मठ एवं विश्वस्त साथी।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      false, "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ],

    // 4. सरपंच प्रत्याशी: रामेश्वर जाट (कनईकलां - पूरी ग्राम पंचायत, सारे वार्ड खुलेंगे, पर्ची पर वार्ड नहीं छपेगा)
    [
      "sarpanch_kanai", "kanai2026", "रामेश्वर जाट", "9829000001", "सरपंच",
      "ALL", "कनईकला", "ALL",
      "उगता सूरज", "", "", "ग्राम पंचायत के सर्वांगीण विकास हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      false, "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ],

    // 5. वार्ड पंच प्रत्याशी: सुभाष गुर्जर (कनईकलां - केवल वार्ड 4 खुलेगा व लॉक रहेगा)
    [
      "panch_w4", "panch2026", "सुभाष गुर्जर", "9829000004", "वार्ड पंच",
      "4", "कनईकला", "4",
      "पतंग", "", "", "वार्ड के विकास एवं स्वच्छता हेतु भारी मतों से विजयी बनावें।", "15 अक्टूबर 2026", "प्रातः 7:00 बजे से सायं 5:00 बजे तक",
      false, "CANDIDATE", true, true, true, true, true, true, true, true, nowStr
    ]
  ];

    sheet.getRange(2, 1, defaultUnifiedRows.length, headers.length).setValues(defaultUnifiedRows);

  // Checkbox Validations: Status (16), Can_Search(17) to Can_Download(23)
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(2, 15, 100, 1).setDataValidation(checkRule); // Allow_Voter_Photo
  for (var c = 17; c <= 24; c++) {
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
/**
 * प्रत्याशी सेटिंग्स एवं नियंत्रण कक्ष वेब ऐप HTML आउटपुट जनरेटर
 */
function getCandidateSettingsHtmlOutput() {
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

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = (params.action || "").trim();
  var callback = params.callback;

  // JSON API कॉल्स की सूची
  var jsonActions = ["ping", "init", "setup", "setupAllSheets", "getSettings", "saveSettings", "updateSetting", "getUsers", "getCandidates", "getPortalJson", "getFullJson", "syncJson", "login", "saveCandidate", "deleteCandidate", "updatePassword", "changePassword"];

  // 1. यदि कोई JSON एक्शन नहीं है अथवा वेब पेज मांगा गया है, तो सीधे सम्पूर्ण प्रत्याशी प्रबंधन वेब ऐप HTML आउटपुट दें!
  if (!action || !jsonActions.includes(action) || action === "candidateSettings" || action === "candidate_settings" || action === "html" || action === "admin" || action === "web") {
    return getCandidateSettingsHtmlOutput();
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

      case "html":
      case "candidateSettings":
      case "candidate_settings":
      case "admin":
      case "web":
        return getCandidateSettingsHtmlOutput();

      default:
        // यदि कोई अपरिचित एक्शन या वेब पेज खोला गया हो तो एरर देने के बजाय सीधे HTML वेब ऐप लोड करें
        return getCandidateSettingsHtmlOutput();
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
    if (action === "login" || action === "auth") {
      res = handleLogin(payload.username || payload.user_id, payload.password);
    } else if (action === "saveSettings") {
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

  return { success: true, settings: settingsObj, spreadsheet_url: ss.getUrl() };
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
  var allowPhotoCol = headers.indexOf("Allow_Voter_Photo") !== -1 ? headers.indexOf("Allow_Voter_Photo") : headers.indexOf("Allow_Voter_Photo_On_Normal_Slips");

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
      symbol_name: (symCol !== -1 && row[symCol] !== undefined && row[symCol] !== null) ? String(row[symCol]).trim() : "",
      symbol_icon: (symCol !== -1 && row[symCol] !== undefined && row[symCol] !== null) ? String(row[symCol]).trim() : "",
      custom_symbol_url: (custSymCol !== -1 && row[custSymCol]) ? String(row[custSymCol]) : "",
      photo_url: photoVal,
      slogan: String(row[sloganCol] || ""),
      election_date: String(row[dateCol] || "15 अक्टूबर 2026"),
      election_time: String(row[timeCol] || "प्रातः 7:00 बजे से सायं 5:00 बजे तक"),
      mobile: String(row[mobCol] || ""),
      allow_voter_photo: (allowPhotoCol !== -1 && (row[allowPhotoCol] === true || String(row[allowPhotoCol]).toUpperCase() === "TRUE")),
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
      symbol_name: (c.symbol_name !== undefined && c.symbol_name !== null) ? String(c.symbol_name).trim() : "",
      symbol_icon: (c.symbol_icon !== undefined && c.symbol_icon !== null) ? String(c.symbol_icon).trim() : "",
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

  var targetRows = [];
  var isKamla = (uName.toLowerCase() === "kamla" || uName.toLowerCase() === "kamala");

  for (var i = 1; i < data.length; i++) {
    var rowU = String(data[i][uCol]).toLowerCase();
    if (isKamla && (rowU === "kamla" || rowU === "kamala")) {
      targetRows.push(i + 1);
    } else if (rowU === uName.toLowerCase()) {
      targetRows.push(i + 1);
      break;
    }
  }

  if (targetRows.length === 0) {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, uCol + 1).setValue(uName);
    var nowStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
    var cAtCol = headers.indexOf("Created_At");
    if (cAtCol !== -1) sheet.getRange(newRow, cAtCol + 1).setValue(nowStr);
    targetRows.push(newRow);
  }

  var fieldMap = {
    "Candidate_Name": cand.candidate_name || cand.full_name || cand.fullName,
    "Post": cand.post,
    "Print_Ward_No": cand.print_ward_no || cand.ward,
    "Allowed_Panchayats": cand.allowed_panchayats || cand.panchayat,
    "Allowed_Panchayat_Wards": cand.allowed_panchayat_wards || cand.allowed_wards,
    "Symbol_Name": (cand.symbol_name !== undefined) ? cand.symbol_name : "",
    "Custom_Symbol_Url": cand.custom_symbol_url || cand.customSymbolUrl || "",
    "Photo_Url": cand.photo_url || cand.photo || cand.photoUrl || "",
    "Slogan": cand.slogan,
    "Election_Date": cand.election_date || cand.electionDate,
    "Election_Time": cand.election_time || cand.electionTime,
    "Mobile": cand.mobile,
    "Allow_Voter_Photo": (cand.allow_voter_photo === true || String(cand.allow_voter_photo).toLowerCase() === "true"),
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
  } else if (targetRows.length > 0 && targetRows[0] > data.length) {
    fieldMap["Password"] = "bhinai2026";
  }

  targetRows.forEach(function(tRow) {
    for (var key in fieldMap) {
      var cIdx = headers.indexOf(key);
      if (cIdx !== -1 && fieldMap[key] !== undefined) {
        sheet.getRange(tRow, cIdx + 1).setValue(fieldMap[key]);
      }
    }
  });

  // Settings शीट में अंतिम सिंक समय अपडेट करें
  try {
    var setSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
    if (setSheet) {
      var sData = setSheet.getDataRange().getValues();
      for (var s = 1; s < sData.length; s++) {
        var sKey = String(sData[s][0]).trim();
        if (sKey === "LAST_CANDIDATE_SYNC") {
          setSheet.getRange(s + 1, 2).setValue(new Date().toISOString());
        }
        if (cand.election_date && sKey === "ELECTION_DATE") {
          setSheet.getRange(s + 1, 2).setValue(cand.election_date);
        }
        if (cand.election_time && sKey === "ELECTION_TIME") {
          setSheet.getRange(s + 1, 2).setValue(cand.election_time);
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
function apiLogin(username, password) {
  return handleLogin(username, password);
}

function apiUpdatePassword(username, newPassword) {
  return handleUpdatePassword(username, newPassword);
}

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
  return "<!DOCTYPE html>\r\n<html lang=\"hi\">\r\n<head>\r\n  <meta charset=\"UTF-8\">\r\n  <base target=\"_top\">\r\n  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\r\n  <title>प्रत्याशी प्रबंधन एवं सेटिंग्स कमान कक्ष - पंचायत चुनाव 2026</title>\r\n  <style>\r\n    :root {\r\n      --primary: #1e3a8a; --primary-light: #3b82f6; --success: #059669;\r\n      --danger: #dc2626; --warning: #d97706; --bg: #f8fafc;\r\n      --card-bg: #ffffff; --text: #0f172a; --text-muted: #64748b;\r\n      --border: #cbd5e1;\r\n    }\r\n    * { box-sizing: border-box; margin:0; padding:0; font-family: -apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, \"Helvetica Neue\", Arial, sans-serif; }\r\n    body { background: var(--bg); color: var(--text); padding: 16px; font-size: 13px; line-height: 1.5; }\r\n    .top-header { background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%); color: #fff; padding: 14px 20px; border-radius: 12px; margin-bottom: 16px; box-shadow: 0 4px 12px rgba(30,58,138,0.2); display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; }\r\n    .top-title { font-size: 1.25rem; font-weight: 800; display:flex; align-items:center; gap:8px; }\r\n    .top-sub { font-size: 0.8rem; opacity: 0.9; margin-top: 3px; }\r\n    .top-actions { display:flex; gap:8px; flex-wrap:wrap; align-items:center; }\r\n    .btn { padding: 7px 14px; border-radius: 6px; font-weight: 700; font-size: 12px; cursor: pointer; border: 1px solid transparent; transition: all 0.15s; display:inline-flex; align-items:center; gap:5px; }\r\n    .btn-primary { background: var(--primary-light); color:#fff; }\r\n    .btn-primary:hover { background: #2563eb; }\r\n    .btn-success { background: var(--success); color:#fff; }\r\n    .btn-success:hover { background: #047857; }\r\n    .btn-warning { background: var(--warning); color:#fff; }\r\n    .btn-danger { background: var(--danger); color:#fff; }\r\n    .btn-outline { background: #fff; border-color: var(--border); color: var(--text); }\r\n    .btn-outline:hover { background: #f1f5f9; }\r\n    .main-grid { display: grid; grid-template-columns: 340px 1fr; gap: 16px; align-items: start; }\r\n    @media(max-width: 900px) { .main-grid { grid-template-columns: 1fr; } }\r\n    .panel-card { background: var(--card-bg); border: 1px solid var(--border); border-radius: 10px; padding: 16px; box-shadow: 0 2px 6px rgba(0,0,0,0.05); }\r\n    .panel-head { font-size: 14px; font-weight: 800; color: #1e3a8a; margin-bottom: 12px; display:flex; justify-content:space-between; align-items:center; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; }\r\n    .cand-search-input { width: 100%; padding: 8px 12px; border: 1px solid var(--border); border-radius: 6px; font-size: 12px; margin-bottom: 10px; outline: none; }\r\n    .cand-search-input:focus { border-color: var(--primary-light); }\r\n    .cand-list { max-height: 680px; overflow-y: auto; display: flex; flex-direction: column; gap: 8px; padding-right: 4px; }\r\n    .cand-item { border: 1px solid var(--border); border-radius: 8px; padding: 10px; background: #fff; cursor: pointer; transition: all 0.15s; display: flex; gap: 10px; align-items: center; position: relative; }\r\n    .cand-item:hover { border-color: var(--primary-light); background: #f0f7ff; }\r\n    .cand-item.active { border-color: #2563eb; background: #eff6ff; box-shadow: 0 0 0 2px rgba(37,99,235,0.2); }\r\n    .cand-thumb { width: 44px; height: 52px; border-radius: 6px; object-fit: cover; background: #e2e8f0; border: 1px solid #cbd5e1; flex-shrink: 0; }\r\n    .cand-item-info { flex: 1; min-width: 0; }\r\n    .cand-item-name { font-weight: 800; font-size: 13px; color: #0f172a; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }\r\n    .cand-item-meta { font-size: 11px; color: var(--text-muted); margin-top: 2px; }\r\n    .cand-badge { display: inline-block; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 4px; margin-top: 3px; }\r\n    .badge-active { background: #dcfce7; color: #15803d; }\r\n    .badge-inactive { background: #fee2e2; color: #b91c1c; }\r\n    .badge-symbol { background: #fef3c7; color: #b45309; font-weight:700; }\r\n    .form-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 12px; margin-bottom: 12px; }\r\n    .form-group { display: flex; flex-direction: column; gap: 4px; }\r\n    .form-group.full-width { grid-column: 1 / -1; }\r\n    .form-label { font-size: 11px; font-weight: 700; color: #334155; text-transform: uppercase; letter-spacing: 0.3px; }\r\n    .form-input, .form-select, .form-textarea { width: 100%; padding: 8px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 13px; font-weight: 600; outline: none; background: #fff; }\r\n    .form-input:focus, .form-select:focus, .form-textarea:focus { border-color: #2563eb; box-shadow: 0 0 0 2px rgba(37,99,235,0.15); }\r\n    .photo-upload-section { background: #f1f5f9; border: 2px dashed #94a3b8; border-radius: 10px; padding: 14px; text-align: center; display: flex; gap: 16px; align-items: center; justify-content: center; flex-wrap: wrap; margin-bottom: 16px; }\r\n    .photo-preview-box { width: 120px; height: 144px; display:flex; align-items:center; justify-content:center; border-radius: 8px; border: 2px solid #3b82f6; overflow: hidden; background: #fff; position: relative; box-shadow: 0 2px 6px rgba(0,0,0,0.1); flex-shrink: 0; }\r\n    .photo-preview-box img { width: 100%; height: 100%; object-fit: cover; object-position: center top; display: block; }\r\n    .photo-desc { text-align: left; max-width: 380px; }\r\n    .specimen-card { background: #fff; border: 2px solid #1e3a8a; border-radius: 8px; overflow: hidden; max-width: 440px; margin: 16px auto; box-shadow: 0 4px 14px rgba(0,0,0,0.12); }\r\n    .specimen-top { background: linear-gradient(to bottom, #ffffff, #fffbeb); border-bottom: 2px solid #1e3a8a; padding: 10px 12px; display: flex; gap: 10px; align-items: center; }\r\n    .specimen-photo { width: 66px; height: 80px; border-radius: 4px; border: 1.5px solid #000; object-fit: cover; object-position: center top; background: #f8fafc; flex-shrink: 0; }\r\n    .specimen-info { flex: 1; text-align: center; }\r\n    .specimen-title { font-size: 10px; font-weight: 700; color: #1e3a8a; }\r\n    .specimen-name { font-size: 15px; font-weight: 900; color: #b91c1c; margin: 2px 0; }\r\n    .specimen-post { font-size: 11px; font-weight: 800; background: #1e3a8a; color: #fff; padding: 2px 6px; border-radius: 3px; display: inline-block; }\r\n    .specimen-gp { font-size: 10px; font-weight: 700; color: #000; margin-top: 3px; }\r\n    .specimen-symbol { width: 56px; text-align: center; flex-shrink: 0; }\r\n    .specimen-symbol-box { width: 52px; height: 52px; border: 1.5px solid #000; border-radius: 4px; padding: 2px; background: #fff; display: flex; align-items: center; justify-content: center; }\r\n    .specimen-symbol-name { font-size: 10px; font-weight: 800; color: #000; margin-top: 2px; }\r\n    .specimen-divider { border-top: 1px dashed #64748b; padding: 4px; text-align: center; font-size: 9px; font-weight: 700; color: #475569; background: #f8fafc; }\r\n    .specimen-bottom { padding: 8px 10px; font-size: 10px; color: #334155; line-height: 1.4; }\r\n    .toast { position: fixed; bottom: 20px; right: 20px; padding: 12px 20px; border-radius: 8px; font-weight: 700; font-size: 13px; color: #fff; box-shadow: 0 4px 12px rgba(0,0,0,0.2); z-index: 9999; display: none; }\r\n    .toast.success { background: #059669; }\r\n    .toast.error { background: #dc2626; }\r\n  </style>\r\n</head>\r\n<body>\r\n  <!-- Top Header Bar -->\r\n  <div class=\"top-header\">\r\n    <div>\r\n      <div class=\"top-title\"><span>🗳️ राजस्थान पंचायती राज चुनाव - 2026</span> <span style=\"font-size:11px; background:rgba(255,255,255,0.2); padding:3px 8px; border-radius:12px;\">SUPER ADMIN COMMAND</span></div>\r\n      <div class=\"top-sub\">सुपर एडमिन प्रत्याशी प्रबंधन, रंगीन फोटो अपलोड एवं एकल मास्टर शीट सिंक कमान कक्ष</div>\r\n    </div>\r\n    <div class=\"top-actions\">\r\n      <button type=\"button\" class=\"btn btn-outline\" onclick=\"openDirectGoogleSheet()\" style=\"background:#fff; font-weight:700; color:#1e3a8a;\">📊 डायरेक्ट गूगल शीट खोलें ↗</button>\n      <button type=\"button\" class=\"btn btn-warning\" onclick=\"createNewCandidate()\">➕ नया प्रत्याशी जोड़ें</button>\r\n      <button type=\"button\" class=\"btn btn-primary\" onclick=\"refreshCandidatesList()\">🔄 रीफ्रेश सूची</button>\r\n      <button type=\"button\" class=\"btn btn-success\" onclick=\"exportPortalJson()\">📥 डाउनलोड portal_users.json</button>\r\n    </div>\r\n  </div>\r\n\r\n  <!-- Master System Settings Bar (Google Sheet Live Sync) -->\r\n  <div style=\"background:#fff; border:1px solid var(--border); border-radius:10px; padding:12px 16px; margin-bottom:16px; box-shadow:0 2px 6px rgba(0,0,0,0.05); display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px;\">\r\n    <div style=\"display:flex; align-items:center; gap:14px; flex-wrap:wrap;\">\r\n      <span style=\"font-weight:800; color:#1e3a8a; font-size:13px;\">⚙️ मास्टर सिस्टम सेटिंग्स (Google Sheet Settings):</span>\r\n      <label style=\"display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-weight:700; color:#0f172a; font-size:12px; background:#eff6ff; padding:5px 10px; border-radius:6px; border:1px solid #bfdbfe;\">\r\n        <input type=\"checkbox\" id=\"masterVoterPhotoCheck\" onchange=\"saveMasterSetting('ALLOW_VOTER_PHOTO_ON_NORMAL_SLIPS', this.checked)\">\r\n        <span>📷 सामान्य मतदाता पर्ची में मतदाता फोटो शामिल करें (Allow Voter Photo on Normal Slips)</span>\r\n      </label>\r\n      <label style=\"display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-weight:700; color:#0f172a; font-size:12px; background:#f0fdf4; padding:5px 10px; border-radius:6px; border:1px solid #bbf7d0;\">\r\n        <input type=\"checkbox\" id=\"masterWithPicCheck\" checked onchange=\"saveMasterSetting('ENABLE_WITH_PIC_SLIPS', this.checked)\">\r\n        <span>🖼️ सचित्र पर्ची (With Pic) मास्टर स्विच</span>\r\n      </label>\r\n    </div>\r\n    <div style=\"font-size:11px; color:#64748b; font-weight:600;\">\r\n      ⚡ यह सेटिंग बदलते ही सभी प्रत्याशियों व डिवाइसों पर गूगल शीट से लाइव अपडेट हो जाएगी\r\n    </div>\r\n  </div>\r\n\r\n  <div class=\"main-grid\">\r\n    <!-- Left Column: Candidates List -->\r\n    <div class=\"panel-card\">\r\n      <div class=\"panel-head\">\r\n        <span>👤 पंजीकृत प्रत्याशी (<span id=\"candCountText\">0</span>)</span>\r\n        <button type=\"button\" class=\"btn btn-xs btn-outline\" onclick=\"createNewCandidate()\" style=\"padding:2px 8px; font-size:11px;\">+ नया</button>\r\n      </div>\r\n      <input type=\"text\" id=\"candSearchInput\" class=\"cand-search-input\" placeholder=\"🔍 प्रत्याशी, पद या पंचायत खोजें...\" oninput=\"filterCandidates(this.value)\">\r\n      <div class=\"cand-list\" id=\"candidatesListContainer\">\r\n        <!-- Candidate Cards rendered here -->\r\n      </div>\r\n    </div>\r\n\r\n    <!-- Right Column: Candidate Editor & Live Specimen -->\r\n    <div class=\"panel-card\">\r\n      <div class=\"panel-head\">\r\n        <span id=\"editorTitleText\">✏️ प्रत्याशी विवरण एवं सेटिंग्स संपादन</span>\r\n        <div>\r\n          <button type=\"button\" class=\"btn btn-danger btn-xs\" id=\"deleteCandBtn\" onclick=\"deleteCurrentCandidate()\" style=\"display:none; padding:3px 10px;\">🗑️ प्रत्याशी हटाएं</button>\r\n        </div>\r\n      </div>\r\n\r\n      <!-- Photo Upload Box (Key Requirement) -->\r\n      <div class=\"photo-upload-section\">\r\n        <div class=\"photo-preview-box\" onclick=\"document.getElementById('photoFileInput').click()\" title=\"फोटो बदलने हेतु क्लिक करें\">\r\n          <img id=\"photoPreviewImg\" src=\"https://api.dicebear.com/7.x/identicon/svg?seed=candidate\" alt=\"Photo\">\r\n        </div>\r\n        <div class=\"photo-desc\">\r\n          <label style=\"font-weight:800; font-size:12px; color:#1e3a8a; display:block; margin-bottom:4px;\">📸 प्रत्याशी की रंगीन फोटो (Color Photo):</label>\r\n          <p style=\"font-size:11px; color:#64748b; margin-bottom:8px;\">मतदाता पर्ची के शीर्ष प्रचार बैनर पर यह फोटो पूरी तरह रंगीन मुद्रित होगी। फोटो का आकार स्वतः ऑप्टिमाइज़ होकर गूगल शीट में सुरक्षित होगा।</p>\r\n          <div style=\"display:flex; gap:8px; align-items:center;\">\r\n            <label for=\"photoFileInput\" class=\"btn btn-primary\" style=\"cursor:pointer; margin:0;\">📁 फोटो चुनें / बदलें</label>\r\n            <input type=\"file\" id=\"photoFileInput\" accept=\"image/*\" style=\"display:none;\" onchange=\"onPhotoFileSelected(event)\">\r\n            <button type=\"button\" id=\"removePhotoBtn\" class=\"btn btn-outline\" onclick=\"removePhoto()\" style=\"display:none;\">❌ हटाएं</button>\r\n          </div>\r\n        </div>\r\n      </div>\r\n\r\n      <!-- Form Details Grid -->\r\n      <form id=\"candidateForm\" onsubmit=\"handleFormSubmit(event)\">\r\n        <div class=\"form-grid\">\r\n          <div class=\"form-group\">\r\n            <label class=\"form-label\" for=\"formUsername\">यूजरनेम (Username) *:</label>\r\n            <input type=\"text\" id=\"formUsername\" class=\"form-input\" required placeholder=\"उदा. kamla / sarpanch01\" oninput=\"updateLiveSpecimen()\">\r\n          </div>\r\n          <div class=\"form-group\">\r\n            <label class=\"form-label\" for=\"formPassword\">पासवर्ड (Password) *:</label>\r\n            <input type=\"text\" id=\"formPassword\" class=\"form-input\" required value=\"bhinai2026\">\r\n          </div>\r\n          <div class=\"form-group\">\r\n            <label class=\"form-label\" for=\"formCandName\">प्रत्याशी का पूरा नाम (Full Name) *:</label>\r\n            <input type=\"text\" id=\"formCandName\" class=\"form-input\" required placeholder=\"उदा. श्रीमती कमला देवी\" oninput=\"updateLiveSpecimen()\" style=\"color:#b91c1c; font-weight:800;\">\r\n          </div>\r\n          <div class=\"form-group\">\r\n            <label class=\"form-label\" for=\"formMobile\">सम्पर्क / मोबाइल नं.:</label>\r\n            <input type=\"tel\" id=\"formMobile\" class=\"form-input\" placeholder=\"उदा. 9876543210\" maxlength=\"10\">\r\n          </div>\r\n        </div>\r\n\r\n        <div class=\"form-grid\">\r\n          <div class=\"form-group\">\r\n            <label class=\"form-label\" for=\"formPost\">चुनाव पद (Contesting Post) *:</label>\r\n            <select id=\"formPost\" class=\"form-select\" onchange=\"onPostChanged(this.value)\">\r\n              <option value=\"सरपंच\">सरपंच पद हेतु</option>\r\n              <option value=\"पंचायत समिति सदस्य\">पंचायत समिति सदस्य हेतु</option>\r\n              <option value=\"जिला परिषद सदस्य\">जिला परिषद सदस्य हेतु</option>\r\n              <option value=\"वार्ड पंच\">वार्ड पंच पद हेतु</option>\r\n              <option value=\"उपसरपंच\">उपसरपंच पद हेतु</option>\r\n            </select>\r\n          </div>\r\n          <div class=\"form-group\">\r\n            <label class=\"form-label\" for=\"formGp\">मुख्य ग्राम पंचायत (Gram Panchayat) *:</label>\r\n            <select id=\"formGp\" class=\"form-select\" onchange=\"updateLiveSpecimen()\"></select>\r\n          </div>\r\n          <div class=\"form-group\">\r\n            <label class=\"form-label\" for=\"formPrintWard\">पर्ची पर छपने वाला वार्ड नं.:</label>\r\n            <input type=\"text\" id=\"formPrintWard\" class=\"form-input\" placeholder=\"उदा. 2 (सरपंच हेतु खाली/ALL)\" oninput=\"updateLiveSpecimen()\">\r\n          </div>\r\n          <div class=\"form-group\">\r\n            <label class=\"form-label\" for=\"formSymbolSelect\">चुनाव चिन्ह (Election Symbol) *:</label>\r\n            <select id=\"formSymbolSelect\" class=\"form-select\" onchange=\"onSymbolChanged(this.value)\"></select>\r\n          </div>\r\n        </div>\r\n\r\n        <div class=\"form-grid\">\r\n          <div class=\"form-group full-width\">\r\n            <label class=\"form-label\" for=\"formAllowedGps\">अनुमति प्राप्त पंचायतें (Allowed Panchayats):</label>\r\n            <input type=\"text\" id=\"formAllowedGps\" class=\"form-input\" placeholder=\"उदा. बांदनवाड़ा या बांदनवाड़ा, देवलियाकलां या ALL\">\r\n          </div>\r\n          <div class=\"form-group full-width\">\r\n            <label class=\"form-label\" for=\"formAllowedWards\">आंशिक वार्ड दायरा (Allowed Panchayat Wards):</label>\r\n            <input type=\"text\" id=\"formAllowedWards\" class=\"form-input\" placeholder=\"उदा. बांदनवाड़ा: ALL या बांदनवाड़ा: 1,2,3 | देवलियाकलां: 1,2,3,4\">\r\n          </div>\r\n        </div>\r\n\r\n        <div class=\"form-grid\">\r\n          <div class=\"form-group full-width\">\r\n            <label class=\"form-label\" for=\"formSlogan\">चुनाव प्रचार अपील / नारा (Campaign Slogan):</label>\r\n            <textarea id=\"formSlogan\" class=\"form-textarea\" rows=\"2\" placeholder=\"क्षेत्र के सर्वांगीण विकास एवं खुशहाली हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।\" oninput=\"updateLiveSpecimen()\"></textarea>\r\n          </div>\r\n        </div>\r\n\r\n        <div class=\"form-grid\">\r\n          <div class=\"form-group\">\r\n            <label class=\"form-label\" for=\"formDate\">मतदान दिनांक (Election Date):</label>\r\n            <input type=\"text\" id=\"formDate\" class=\"form-input\" value=\"15 अक्टूबर 2026\" oninput=\"updateLiveSpecimen()\">\r\n          </div>\r\n          <div class=\"form-group\">\r\n            <label class=\"form-label\" for=\"formTime\">मतदान समय (Voting Time):</label>\r\n            <input type=\"text\" id=\"formTime\" class=\"form-input\" value=\"प्रातः 7:00 बजे से सायं 5:00 बजे तक\" oninput=\"updateLiveSpecimen()\">\r\n          </div>\r\n          <div class=\"form-group\" style=\"justify-content:center;\">\r\n            <label style=\"display:flex; align-items:center; gap:6px; cursor:pointer; font-weight:700; margin-top:14px;\">\r\n              <input type=\"checkbox\" id=\"formStatus\" checked style=\"width:16px; height:16px;\"> 🟢 खाता सक्रिय रखें (ACTIVE)\r\n            </label>\r\n          </div>\r\n        </div>\r\n\r\n        <!-- Live Perforated Slip Specimen Card -->\r\n        <div style=\"margin-top:16px; background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:12px;\">\r\n          <div style=\"font-weight:800; font-size:12px; color:#334155; margin-bottom:6px; text-align:center;\">📋 लाइव पर्ची प्रारूप नमूना (Live Voter Slip Specimen)</div>\r\n          <div class=\"specimen-card\" id=\"liveSpecimenCard\">\r\n            <div class=\"specimen-top\">\r\n              <img id=\"specPhoto\" class=\"specimen-photo\" src=\"https://api.dicebear.com/7.x/identicon/svg?seed=candidate\" alt=\"Photo\">\r\n              <div class=\"specimen-info\">\r\n                <div class=\"specimen-title\">।। श्री गणेशाय नमः ।। 🙏 विनम्र अपील</div>\r\n                <div class=\"specimen-name\" id=\"specName\">कमला देवी</div>\r\n                <div><span class=\"specimen-post\" id=\"specPost\">पंचायत समिति सदस्य प्रत्याशी</span></div>\r\n                <div class=\"specimen-gp\" id=\"specGpRow\">ग्राम पंचायत: बांदनवाड़ा • वार्ड नं.: 2</div>\r\n                <div style=\"font-size:9px; color:#475569; margin-top:2px;\">📅 मतदान: <strong id=\"specDate\">15 अक्टूबर 2026</strong></div>\r\n              </div>\r\n              <div class=\"specimen-symbol\">\r\n                <div class=\"specimen-symbol-box\" id=\"specSymbolBox\">✋</div>\r\n                <div class=\"specimen-symbol-name\" id=\"specSymbolName\">हाथ</div>\r\n              </div>\r\n            </div>\r\n            <div class=\"specimen-divider\">✂️ --- यहाँ से काटकर मतदाता को दें --- ✂️</div>\r\n            <div class=\"specimen-bottom\">\r\n              <div style=\"display:flex; justify-content:space-between; font-weight:700;\"><span>मतदाता: रमेश कुमार</span> <span>क्रम: 104</span></div>\r\n              <div style=\"font-size:9px; color:#64748b;\">मतदान केंद्र: रा.उ.मा.वि. कक्ष-1</div>\r\n            </div>\r\n          </div>\r\n        </div>\r\n\r\n        <!-- Action Buttons -->\r\n        <div style=\"display:flex; gap:10px; margin-top:16px; flex-wrap:wrap;\">\r\n          <button type=\"submit\" id=\"saveSubmitBtn\" class=\"btn btn-success\" style=\"flex:1; padding:12px; font-size:14px; justify-content:center;\">\r\n            💾 गूगल शीट में सुरक्षित करें (Save to Google Sheet)\r\n          </button>\r\n          <button type=\"button\" class=\"btn btn-outline\" onclick=\"createNewCandidate()\" style=\"padding:12px 18px;\">\r\n            ➕ नया प्रत्याशी\r\n          </button>\r\n        </div>\r\n      </form>\r\n    </div>\r\n  </div>\r\n\r\n  <!-- Toast notification -->\r\n  <div id=\"toastBox\" class=\"toast\"></div>\r\n\r\n  <script>\r\n    var candidates = [];\r\n    var officialGps = [\r\n      \"कनईकला\", \"करांटी\", \"कैरोंट\", \"खेडी\", \"कुम्हारिया\",\r\n      \"गुढाखुर्द\", \"घणा\", \"चापानेरी\", \"छछून्दरा\", \"धांतोल\",\r\n      \"देवपुरा\", \"देवलियाकलां\", \"नागोला\", \"नान्दसी\", \"पड़ांगा\",\r\n      \"पाडलिया\", \"बगराई\", \"बड़गांव\", \"बड़ली\", \"बांदनवाड़ा\",\r\n      \"बूबकिया\", \"भिनाय\", \"राममालिया\", \"राताकोट\", \"लामगरा\",\r\n      \"सिंगावल\", \"सोबडी\", \"सोलखुर्द\", \"हियालिया\", \"एकलसिंहा\"\r\n    ];\r\n    var officialSymbols = [\r\n      \"हाथ\", \"कमल\", \"उगता सूरज\", \"पतंग\", \"बोतल\", \"साइकिल\",\r\n      \"गिलास\", \"कार\", \"ताला-चाबी\", \"बल्ला\", \"पंखा\", \"अलमारी\", \"सिलेंडर\", \"मोमबत्ती\",\r\n      \"टेलीविजन\", \"मेज\", \"कुर्सी\", \"कप-प्लेट\", \"कैमरा\", \"किताब\", \"ट्रेक्टर\", \"रेल का इंजन\"\r\n    ];\r\n    var activeCandidate = null;\r\n    var currentPhotoBase64 = \"\";\r\n\r\n    initApp();\r\n\r\n    function initApp() {\r\n      populateDropdowns();\r\n      checkAuthOnLoad();\r\n      refreshCandidatesList();\r\n    }\r\n\r\n    function populateCandidateDropdown() {\n      var sel = document.getElementById(\"topCandSelectDropdown\");\n      if (!sel) return;\n      sel.innerHTML = \"\";\n      candidates.forEach(function(c) {\n        var opt = document.createElement(\"option\");\n        opt.value = c.username;\n        var pText = c.post ? (\" [\" + c.post + \" - \" + (c.panchayat || \"\") + \"]\") : \"\";\n        opt.textContent = (c.candidate_name || c.username) + pText + (c.allow_voter_photo ? \" 📸 [फोटो अनुमत]\" : \"\");\n        sel.appendChild(opt);\n      });\n      if (activeCandidate) sel.value = activeCandidate.username;\n    }\n\n    var liveSpreadsheetUrl = \"https://docs.google.com/spreadsheets/d/1pWouRxbkasaKcexqKdZJGqScqOigv9Hm-L-0skVq6Dc/edit\";\n\n    function openDirectGoogleSheet() {\n      window.open(liveSpreadsheetUrl, \"_blank\");\n    }\n\n    function populateDropdowns() {\r\n      var gpSelect = document.getElementById(\"formGp\");\r\n      gpSelect.innerHTML = \"\";\r\n      officialGps.forEach(function(gp) {\r\n        var opt = document.createElement(\"option\");\r\n        opt.value = gp; opt.textContent = gp;\r\n        gpSelect.appendChild(opt);\r\n      });\r\n\r\n      var symSelect = document.getElementById(\"formSymbolSelect\");\r\n      symSelect.innerHTML = \"<option value=\\\"\\\">-- कोई चुनाव चिन्ह नहीं (केवल यूजर/प्रशासक) --</option>\";\r\n      officialSymbols.forEach(function(sym) {\r\n        var opt = document.createElement(\"option\");\r\n        opt.value = sym; opt.textContent = sym;\r\n        symSelect.appendChild(opt);\r\n      });\r\n    }\r\n\r\n    function onThumbError(el) { if (el) { el.onerror = null; el.src = \"https://api.dicebear.com/7.x/identicon/svg?seed=cand\"; } }\r\n\r\n    function renderCandidatesList(filterText) {\r\n      var container = document.getElementById(\"candidatesListContainer\");\r\n      container.innerHTML = \"\";\r\n      var q = (filterText || \"\").trim().toLowerCase();\r\n      var count = 0;\r\n\r\n      candidates.forEach(function(c) {\r\n        var text = (c.candidate_name + \" \" + c.post + \" \" + c.panchayat + \" \" + c.username).toLowerCase();\r\n        if (q && text.indexOf(q) === -1) return;\r\n        count++;\r\n\r\n        var item = document.createElement(\"div\");\r\n        var isActive = activeCandidate && (activeCandidate.username === c.username);\r\n        item.className = \"cand-item\" + (isActive ? \" active\" : \"\");\r\n        item.onclick = function() { selectCandidate(c.username); };\r\n\r\n        var photoSrc = c.photo_url || (c.username === \"kamla\" || c.username === \"kamala\" ? \"./assets/kamala_photo.jpg\" : \"https://api.dicebear.com/7.x/identicon/svg?seed=\" + encodeURIComponent(c.username));\r\n\r\n        item.innerHTML = \r\n          '<img class=\"cand-thumb\" src=\"' + photoSrc + '\" alt=\"' + (c.candidate_name || c.username) + '\" onerror=\"onThumbError(this)\">' +\r\n          '<div class=\"cand-item-info\">' +\r\n          '  <div class=\"cand-item-name\">' + (c.candidate_name || c.username) + '</div>' +\r\n          '  <div class=\"cand-item-meta\">' + (c.post || \"\") + ' • ' + (c.panchayat || \"\") + '</div>' +\r\n          '  <div>' +\r\n          '    <span class=\"cand-badge ' + (c.status === \"ACTIVE\" ? \"badge-active\" : \"badge-inactive\") + '\">' + (c.status === \"ACTIVE\" ? \"🟢 एक्टिव\" : \"🔴 बंद\") + '</span> ' +\r\n          '    <span class=\"cand-badge badge-symbol\">🚩 ' + (c.symbol_name || \"चिन्ह\") + '</span>' + (c.allow_voter_photo ? ' <span class=\"cand-badge\" style=\"background:#e0f2fe; color:#0369a1; font-weight:700;\">📸 फोटो चालू</span>' : '') +\r\n          '  </div>' +\r\n          '</div>';\r\n\r\n        container.appendChild(item);\r\n      });\r\n\r\n      document.getElementById(\"candCountText\").textContent = count;\n      populateCandidateDropdown();\r\n    }\r\n\r\n    function filterCandidates(q) {\r\n      renderCandidatesList(q);\r\n    }\r\n\r\n    function selectCandidate(uName) {\r\n      var c = candidates.find(function(x) { return x.username === uName; });\r\n      if (!c) return;\r\n      activeCandidate = c;\r\n      currentPhotoBase64 = c.photo_url || \"\";\r\n\r\n      document.getElementById(\"editorTitleText\").textContent = \"✏️ \" + (c.candidate_name || c.username) + \" का विवरण संपादन\";\r\n      document.getElementById(\"deleteCandBtn\").style.display = (c.username === \"admin\" ? \"none\" : \"inline-block\");\r\n\r\n      document.getElementById(\"formUsername\").value = c.username;\r\n      document.getElementById(\"formUsername\").readOnly = (c.username === \"admin\" || c.username === \"kamala\" || c.username === \"kamla\");\r\n      document.getElementById(\"formPassword\").value = c.password || \"bhinai2026\";\r\n      document.getElementById(\"formCandName\").value = c.candidate_name || \"\";\r\n      document.getElementById(\"formMobile\").value = c.mobile || \"\";\r\n      document.getElementById(\"formPost\").value = c.post || \"सरपंच\";\r\n      document.getElementById(\"formGp\").value = c.panchayat || officialGps[0];\r\n      document.getElementById(\"formPrintWard\").value = c.print_ward_no || c.ward || \"\";\r\n      document.getElementById(\"formSymbolSelect\").value = c.symbol_name || \"\";\r\n      document.getElementById(\"formAllowedGps\").value = c.allowed_panchayats || c.panchayat || \"ALL\";\r\n      document.getElementById(\"formAllowedWards\").value = c.allowed_panchayat_wards || c.allowed_wards || \"ALL\";\r\n      document.getElementById(\"formSlogan\").value = c.slogan || \"\";\r\n      document.getElementById(\"formDate\").value = c.election_date || \"15 अक्टूबर 2026\";\r\n      document.getElementById(\"formTime\").value = c.election_time || \"प्रातः 7:00 बजे से सायं 5:00 बजे तक\";\r\n      document.getElementById(\"formStatus\").checked = (c.status === \"ACTIVE\");\n      document.getElementById(\"formAllowVoterPhoto\").checked = !!(c.allow_voter_photo);\n      var topSel = document.getElementById(\"topCandSelectDropdown\"); if (topSel) topSel.value = c.username;\r\n\r\n      var previewImg = document.getElementById(\"photoPreviewImg\");\r\n      if (c.photo_url) {\r\n        previewImg.src = c.photo_url;\r\n        document.getElementById(\"removePhotoBtn\").style.display = \"inline-block\";\r\n      } else if (c.username === \"kamla\" || c.username === \"kamala\") {\r\n        previewImg.src = \"./assets/kamala_photo.jpg\";\r\n        document.getElementById(\"removePhotoBtn\").style.display = \"inline-block\";\r\n      } else {\r\n        previewImg.src = \"https://api.dicebear.com/7.x/identicon/svg?seed=\" + encodeURIComponent(c.username);\r\n        document.getElementById(\"removePhotoBtn\").style.display = \"none\";\r\n      }\r\n\r\n      renderCandidatesList(document.getElementById(\"candSearchInput\").value);\r\n      updateLiveSpecimen();\r\n    }\r\n\r\n    function createNewCandidate() {\r\n      activeCandidate = null;\r\n      currentPhotoBase64 = \"\";\r\n      document.getElementById(\"editorTitleText\").textContent = \"➕ नया प्रत्याशी जोड़ें\";\r\n      document.getElementById(\"deleteCandBtn\").style.display = \"none\";\r\n\r\n      document.getElementById(\"formUsername\").value = \"cand_\" + Math.floor(100 + Math.random() * 900);\r\n      document.getElementById(\"formUsername\").readOnly = false;\r\n      document.getElementById(\"formPassword\").value = \"bhinai2026\";\r\n      document.getElementById(\"formCandName\").value = \"\";\r\n      document.getElementById(\"formMobile\").value = \"\";\r\n      document.getElementById(\"formPost\").value = \"सरपंच\";\r\n      document.getElementById(\"formGp\").value = officialGps[0];\r\n      document.getElementById(\"formPrintWard\").value = \"\";\r\n      document.getElementById(\"formSymbolSelect\").value = \"\";\r\n      document.getElementById(\"formAllowedGps\").value = officialGps[0];\r\n      document.getElementById(\"formAllowedWards\").value = \"ALL\";\r\n      document.getElementById(\"formSlogan\").value = \"क्षेत्र के सर्वांगीण विकास एवं खुशहाली हेतु अपने कर्मठ, ईमानदार एवं सेवाभावी प्रत्याशी को भारी मतों से विजयी बनावें।\";\r\n      document.getElementById(\"formDate\").value = \"15 अक्टूबर 2026\";\r\n      document.getElementById(\"formTime\").value = \"प्रातः 7:00 बजे से सायं 5:00 बजे तक\";\r\n      document.getElementById(\"formStatus\").checked = true;\n      document.getElementById(\"formAllowVoterPhoto\").checked = false;\r\n\r\n      document.getElementById(\"photoPreviewImg\").src = \"https://api.dicebear.com/7.x/identicon/svg?seed=new\";\r\n      document.getElementById(\"removePhotoBtn\").style.display = \"none\";\r\n\r\n      renderCandidatesList();\r\n      updateLiveSpecimen();\r\n    }\r\n\r\n    function onPostChanged(post) {\r\n      var wardInput = document.getElementById(\"formPrintWard\");\r\n      if (post === \"पंचायत समिति सदस्य\") {\r\n        if (!wardInput.value || wardInput.value === \"ALL\") wardInput.value = \"2\";\r\n      } else if (post === \"जिला परिषद सदस्य\") {\r\n        if (!wardInput.value || wardInput.value === \"ALL\") wardInput.value = \"15\";\r\n      } else if (post === \"सरपंच\") {\r\n        wardInput.value = \"ALL\";\r\n      }\r\n      updateLiveSpecimen();\r\n    }\r\n\r\n    function onSymbolChanged(sym) {\r\n      updateLiveSpecimen();\r\n    }\r\n\r\n    // Photo Upload Handler with Client-Side Canvas Compression\r\n    function onPhotoFileSelected(e) {\r\n      var file = e.target.files && e.target.files[0];\r\n      if (!file) return;\r\n      showToast(\"⏳ फोटो प्रोसेस हो रही है...\", \"success\");\r\n\r\n      var reader = new FileReader();\r\n      reader.onload = function(evt) {\r\n        var img = new Image();\r\n        img.onload = function() {\r\n          var maxW = 360, maxH = 450;\r\n          var w = img.width, h = img.height;\r\n          if (w > maxW || h > maxH) {\r\n            if (w/h > maxW/maxH) {\r\n              h = Math.round(h * (maxW / w));\r\n              w = maxW;\r\n            } else {\r\n              w = Math.round(w * (maxH / h));\r\n              h = maxH;\r\n            }\r\n          }\r\n          var canvas = document.createElement(\"canvas\");\r\n          canvas.width = w;\r\n          canvas.height = h;\r\n          var ctx = canvas.getContext(\"2d\");\r\n          ctx.drawImage(img, 0, 0, w, h);\r\n          var base64 = canvas.toDataURL(\"image/jpeg\", 0.82);\r\n          currentPhotoBase64 = base64;\r\n          document.getElementById(\"photoPreviewImg\").src = base64;\r\n          document.getElementById(\"removePhotoBtn\").style.display = \"inline-block\";\r\n          updateLiveSpecimen();\r\n          showToast(\"✅ फोटो तैयार! सुरक्षित करने हेतु नीचे \\\\'सुरक्षित करें\\\\' दबाएं।\", \"success\");\r\n        };\r\n        img.src = evt.target.result;\r\n      };\r\n      reader.readAsDataURL(file);\r\n    }\r\n\r\n    function removePhoto() {\r\n      currentPhotoBase64 = \"\";\r\n      document.getElementById(\"photoFileInput\").value = \"\";\r\n      document.getElementById(\"photoPreviewImg\").src = \"https://api.dicebear.com/7.x/identicon/svg?seed=\" + encodeURIComponent(document.getElementById(\"formUsername\").value);\r\n      document.getElementById(\"removePhotoBtn\").style.display = \"none\";\r\n      updateLiveSpecimen();\r\n    }\r\n\r\n    function updateLiveSpecimen() {\r\n      var name = document.getElementById(\"formCandName\").value.trim() || document.getElementById(\"formUsername\").value;\r\n      var post = document.getElementById(\"formPost\").value;\r\n      var gp = document.getElementById(\"formGp\").value;\r\n      var ward = document.getElementById(\"formPrintWard\").value.trim();\r\n      var sym = document.getElementById(\"formSymbolSelect\").value;\r\n      var date = document.getElementById(\"formDate\").value;\r\n\r\n      document.getElementById(\"specName\").textContent = name;\r\n      document.getElementById(\"specPost\").textContent = post + \" प्रत्याशी\";\r\n\r\n      var gpRowText = \"ग्राम पंचायत: \" + gp;\r\n      if (ward && ward !== \"ALL\" && post !== \"सरपंच\") {\r\n        gpRowText += \" • वार्ड नं.: \" + ward;\r\n      }\r\n      document.getElementById(\"specGpRow\").textContent = gpRowText;\r\n      document.getElementById(\"specDate\").textContent = date;\r\n      if (!sym) {\r\n        document.getElementById(\"specSymbolName\").textContent = \"कोई चिन्ह नहीं\";\r\n        document.getElementById(\"specSymbolBox\").textContent = \"⚪\";\r\n      } else {\r\n        document.getElementById(\"specSymbolName\").textContent = sym;\r\n      }\r\n\r\n      // Symbol Icon\r\n      var symIconMap = { \"हाथ\": \"✋\", \"कमल\": \"🪷\", \"उगता सूरज\": \"☀️\", \"पतंग\": \"🪁\", \"बोतल\": \"🍾\", \"साइकिल\": \"🚲\", \"गिलास\": \"🥛\", \"कार\": \"🚗\", \"ताला-चाबी\": \"🔐\", \"बल्ला\": \"🏏\", \"पंखा\": \"🌀\", \"सिलेंडर\": \"🛢️\", \"मोमबत्ती\": \"🕯️\" };\r\n      document.getElementById(\"specSymbolBox\").textContent = symIconMap[sym] || \"🚩\";\r\n\r\n      // Photo\r\n      var previewImg = document.getElementById(\"photoPreviewImg\");\r\n      document.getElementById(\"specPhoto\").src = previewImg.src;\r\n    }\r\n\r\n    function handleFormSubmit(e) {\r\n      e.preventDefault();\r\n      var btn = document.getElementById(\"saveSubmitBtn\");\r\n      btn.disabled = true; btn.textContent = \"⏳ गूगल शीट में सुरक्षित हो रहा है...\";\r\n\r\n      var candData = {\r\n        username: document.getElementById(\"formUsername\").value.trim(),\r\n        password: document.getElementById(\"formPassword\").value.trim(),\r\n        candidate_name: document.getElementById(\"formCandName\").value.trim(),\r\n        mobile: document.getElementById(\"formMobile\").value.trim(),\r\n        post: document.getElementById(\"formPost\").value,\r\n        panchayat: document.getElementById(\"formGp\").value,\r\n        allowed_panchayats: document.getElementById(\"formAllowedGps\").value.trim(),\r\n        print_ward_no: document.getElementById(\"formPrintWard\").value.trim(),\r\n        ward: document.getElementById(\"formPrintWard\").value.trim(),\r\n        allowed_panchayat_wards: document.getElementById(\"formAllowedWards\").value.trim(),\r\n        symbol_name: document.getElementById(\"formSymbolSelect\").value,\r\n        photo_url: currentPhotoBase64,\r\n        slogan: document.getElementById(\"formSlogan\").value.trim(),\r\n        election_date: document.getElementById(\"formDate\").value.trim(),\r\n        election_time: document.getElementById(\"formTime\").value.trim(),\r\n        status: document.getElementById(\"formStatus\").checked ? \"ACTIVE\" : \"INACTIVE\",\n        allow_voter_photo: document.getElementById(\"formAllowVoterPhoto\").checked\r\n      };\r\n\r\n      function onSaveSuccess(res) {\r\n        btn.disabled = false; btn.textContent = \"💾 गूगल शीट में सुरक्षित करें (Save to Google Sheet)\";\r\n        if (res && res.success) {\r\n          showToast(\"✅ \" + res.message, \"success\");\r\n          if (res.candidates) candidates = res.candidates;\r\n          else refreshCandidatesList();\r\n          selectCandidate(candData.username);\r\n        } else {\r\n          showToast(\"⚠️ त्रुटि: \" + (res.error || \"सुरक्षित नहीं हो सका\"), \"error\");\r\n        }\r\n      }\r\n\r\n      function onSaveError(err) {\r\n        btn.disabled = false; btn.textContent = \"💾 गूगल शीट में सुरक्षित करें (Save to Google Sheet)\";\r\n        showToast(\"⚠️ नेटवर्क विफलता: \" + err, \"error\");\r\n      }\r\n\r\n      function executeSaveViaFetch() {\r\n        var sUrl = \"https://script.google.com/macros/s/AKfycbziw4V3cYSBnqSmUG0FDP4Ge8r-MY_B9CbGkHYcnIKayL5zV39ZRNTL7sclKz9wdRkEwQ/exec\";\r\n        fetch(sUrl, {\r\n          method: \"POST\",\r\n          headers: { \"Content-Type\": \"text/plain;charset=utf-8\" },\r\n          body: JSON.stringify({ action: \"saveCandidate\", candidateData: candData })\r\n        })\r\n        .then(function(r) { return r.json(); })\r\n        .then(onSaveSuccess)\r\n        .catch(function(err) {\r\n          btn.disabled = false; btn.textContent = \"💾 गूगल शीट में सुरक्षित करें (Save to Google Sheet)\";\r\n          showToast(\"⚠️ नेटवर्क विफलता: \" + err, \"error\");\r\n        });\r\n      }\r\n\r\n      if (typeof google !== \"undefined\" && google.script && google.script.run) {\r\n        google.script.run\r\n          .withSuccessHandler(onSaveSuccess)\r\n          .withFailureHandler(function(err) {\r\n            console.warn(\"google.script.run failed, trying POST fetch fallback:\", err);\r\n            executeSaveViaFetch();\r\n          })\r\n          .apiSaveCandidate(candData);\r\n      } else {\r\n        executeSaveViaFetch();\r\n      }\r\n    }\r\n\r\n    function deleteCurrentCandidate() {\r\n      if (!activeCandidate) return;\r\n      var u = activeCandidate.username;\r\n      if (!confirm(\"क्या आप वाकई प्रत्याशी '\" + activeCandidate.candidate_name + \" (\" + u + \")' को हटाना चाहते हैं?\")) return;\r\n\r\n      showToast(\"⏳ हटाया जा रहा है...\", \"success\");\r\n      function onDelSuccess(res) {\r\n        if (res && res.success) {\r\n          showToast(\"✅ \" + res.message, \"success\");\r\n          candidates = candidates.filter(function(x) { return x.username !== u; });\r\n          if (candidates.length > 0) selectCandidate(candidates[0].username);\r\n          else createNewCandidate();\r\n        } else {\r\n          showToast(\"⚠️ त्रुटि: \" + (res.error || \"हट नहीं सका\"), \"error\");\r\n        }\r\n      }\r\n\r\n      function executeDeleteViaFetch() {\r\n        var sUrl = \"https://script.google.com/macros/s/AKfycbziw4V3cYSBnqSmUG0FDP4Ge8r-MY_B9CbGkHYcnIKayL5zV39ZRNTL7sclKz9wdRkEwQ/exec\";\r\n        fetch(sUrl, {\r\n          method: \"POST\",\r\n          headers: { \"Content-Type\": \"text/plain;charset=utf-8\" },\r\n          body: JSON.stringify({ action: \"deleteCandidate\", username: u })\r\n        })\r\n        .then(function(r) { return r.json(); })\r\n        .then(onDelSuccess)\r\n        .catch(function(err) { showToast(\"⚠️ हटा नहीं सका: \" + err, \"error\"); });\r\n      }\r\n\r\n      if (typeof google !== \"undefined\" && google.script && google.script.run) {\r\n        google.script.run\r\n          .withSuccessHandler(onDelSuccess)\r\n          .withFailureHandler(function(err) { executeDeleteViaFetch(); })\r\n          .apiDeleteCandidate(u);\r\n      } else {\r\n        executeDeleteViaFetch();\r\n      }\r\n    }\r\n\r\n    function refreshCandidatesList() {\r\n      showToast(\"🔄 गूगल शीट से प्रत्याशी लोड हो रहे हैं...\", \"success\");\r\n      function onListSuccess(res) {\r\n        if (res && res.candidates) {\r\n          candidates = res.candidates;\r\n          populateAuthCandidatesDropdown();\r\n          renderCandidatesList();\r\n          if (!activeCandidate && candidates.length > 0) { selectCandidate(candidates[0].username); }\r\n          showToast(\"✅ सूची रीफ्रेश हो गई (\" + candidates.length + \" प्रत्याशी)\", \"success\");\r\n        }\r\n      }\r\n      function executeRefreshViaFetch() {\r\n        var sUrl = \"https://script.google.com/macros/s/AKfycbziw4V3cYSBnqSmUG0FDP4Ge8r-MY_B9CbGkHYcnIKayL5zV39ZRNTL7sclKz9wdRkEwQ/exec?action=getCandidates&_t=\" + Date.now();\r\n        fetch(sUrl)\r\n          .then(function(r) { return r.json(); })\r\n          .then(onListSuccess)\r\n          .catch(function(err) { showToast(\"⚠️ प्रत्याशी लोड नहीं हुए: \" + err, \"error\"); });\r\n      }\r\n\r\n      if (typeof google !== \"undefined\" && google.script && google.script.run) {\r\n        google.script.run\r\n          .withSuccessHandler(onListSuccess)\r\n          .withFailureHandler(function(err) { executeRefreshViaFetch(); })\r\n          .apiGetCandidates();\r\n      } else {\r\n        executeRefreshViaFetch();\r\n      }\r\n    }\r\n\r\n    function exportPortalJson() {\r\n      function onJsonSuccess(res) {\r\n        var data = res.portal_users || res;\r\n        var blob = new Blob([JSON.stringify(data, null, 2)], { type: \"application/json\" });\r\n        var a = document.createElement(\"a\");\r\n        a.href = URL.createObjectURL(blob);\r\n        a.download = \"portal_users.json\";\r\n        a.click();\r\n        showToast(\"✅ portal_users.json डाउनलोड हो गया!\", \"success\");\r\n      }\r\n      if (typeof google !== \"undefined\" && google.script && google.script.run) {\r\n        google.script.run.withSuccessHandler(onJsonSuccess).apiGetPortalJson();\r\n      } else {\r\n        fetch(window.location.href + (window.location.href.indexOf(\"?\") === -1 ? \"?\" : \"&\") + \"action=getPortalJson\")\r\n          .then(function(r) { return r.json(); })\r\n          .then(onJsonSuccess);\r\n      }\r\n    }\r\n\r\n    function copyPortalJson() {\r\n      function onJsonSuccess(res) {\r\n        var data = res.portal_users || res;\r\n        var str = JSON.stringify(data, null, 2);\r\n        navigator.clipboard.writeText(str).then(function() {\r\n          showToast(\"📋 portal_users.json क्लिपबोर्ड में कॉपी हो गया!\", \"success\");\r\n        });\r\n      }\r\n      if (typeof google !== \"undefined\" && google.script && google.script.run) {\r\n        google.script.run.withSuccessHandler(onJsonSuccess).apiGetPortalJson();\r\n      } else {\r\n        fetch(window.location.href + (window.location.href.indexOf(\"?\") === -1 ? \"?\" : \"&\") + \"action=getPortalJson\")\r\n          .then(function(r) { return r.json(); })\r\n          .then(onJsonSuccess);\r\n      }\r\n    }\r\n\r\n    function saveMasterSetting(key, val) {\r\n      showToast(\"⏳ गूगल शीट में मास्टर सेटिंग अपडेट हो रही है...\", \"success\");\r\n      function onSettSuccess(res) {\r\n        showToast(\"✅ गूगल शीट में सुरक्षित: \" + key + \" = \" + val, \"success\");\r\n      }\r\n      function onSettError(err) {\r\n        showToast(\"⚠️ सिंक विफलता: \" + err, \"error\");\r\n      }\r\n      if (typeof google !== \"undefined\" && google.script && google.script.run) {\r\n        var obj = {}; obj[key] = val;\r\n        google.script.run.withSuccessHandler(onSettSuccess).withFailureHandler(onSettError).apiSaveSettings(obj);\r\n      } else {\r\n        var base = (window.location.href.indexOf(\"script.google.com\") !== -1) ? window.location.href.split(\"?\")[0] : \"https://script.google.com/macros/s/AKfycbziw4V3cYSBnqSmUG0FDP4Ge8r-MY_B9CbGkHYcnIKayL5zV39ZRNTL7sclKz9wdRkEwQ/exec\";\r\n        var url = base + \"?action=saveSettings&key=\" + encodeURIComponent(key) + \"&value=\" + encodeURIComponent(val) + \"&_t=\" + Date.now();\r\n        fetch(url).then(function(r){ return r.json(); }).then(onSettSuccess).catch(onSettError);\r\n      }\r\n    }\r\n\r\n    function loadMasterSettings() {\r\n      function applySett(res) {\n        if (res && res.spreadsheet_url) { liveSpreadsheetUrl = res.spreadsheet_url; }\r\n        if (res && res.settings) {\r\n          var photoEl = document.getElementById(\"masterVoterPhotoCheck\");\r\n          if (photoEl && res.settings.ALLOW_VOTER_PHOTO_ON_NORMAL_SLIPS !== undefined) {\r\n            photoEl.checked = (res.settings.ALLOW_VOTER_PHOTO_ON_NORMAL_SLIPS === true || String(res.settings.ALLOW_VOTER_PHOTO_ON_NORMAL_SLIPS).toUpperCase() === \"TRUE\");\r\n          }\r\n          var withPicEl = document.getElementById(\"masterWithPicCheck\");\r\n          if (withPicEl && res.settings.ENABLE_WITH_PIC_SLIPS !== undefined) {\r\n            withPicEl.checked = !!res.settings.ENABLE_WITH_PIC_SLIPS;\r\n          }\r\n        }\r\n      }\r\n      if (typeof google !== \"undefined\" && google.script && google.script.run) {\r\n        google.script.run.withSuccessHandler(applySett).handleGetSettings();\r\n      } else {\r\n        var url = window.location.href.split(\"?\")[0] + \"?action=getSettings\";\r\n        fetch(url).then(function(r){ return r.json(); }).then(applySett).catch(function(){});\r\n      }\r\n    }\r\n\r\n    // Load master settings on startup\r\n    window.addEventListener(\"load\", function() {\r\n      loadMasterSettings();\r\n      checkAuthOnLoad();\r\n    });\r\n\r\n    \r\n    // --- 🔐 Security Review & Authentication System ---\r\n    var currentAuthUser = null;\r\n    var currentAuthRole = 'ADMIN';\r\n\r\n    function checkAuthOnLoad() {\r\n      populateAuthCandidatesDropdown();\r\n      var saved = sessionStorage.getItem(\"cs_auth_user\");\r\n      if (saved) {\r\n        try {\r\n          currentAuthUser = JSON.parse(saved);\r\n          applyAuthPermissions(currentAuthUser);\r\n          return;\r\n        } catch(e) {}\r\n      }\r\n      showAuthModal();\r\n    }\r\n\r\n    function populateAuthCandidatesDropdown() {\r\n      var sel = document.getElementById(\"authCandUsernameSelect\");\r\n      if (!sel) return;\r\n      sel.innerHTML = \"\";\r\n      candidates.forEach(function(c) {\r\n        var opt = document.createElement(\"option\");\r\n        opt.value = c.username;\r\n        opt.textContent = (c.candidate_name || c.username) + \" (\" + (c.post || \"प्रत्याशी\") + \" - \" + (c.panchayat || \"ALL\") + \")\";\r\n        sel.appendChild(opt);\r\n      });\r\n    }\r\n\r\n    function showAuthModal() {\r\n      var modal = document.getElementById(\"authReviewModal\");\r\n      if (modal) {\r\n        modal.style.display = \"flex\";\r\n        document.getElementById(\"authPasswordInput\").value = \"\";\r\n        document.getElementById(\"authErrorMsg\").style.display = \"none\";\r\n        populateAuthCandidatesDropdown();\r\n      }\r\n    }\r\n\r\n    function setAuthRole(role) {\r\n      currentAuthRole = role;\r\n      var btnAdmin = document.getElementById(\"authRoleAdminBtn\");\r\n      var btnCand = document.getElementById(\"authRoleCandBtn\");\r\n      var candGroup = document.getElementById(\"authCandSelectGroup\");\r\n      var passInp = document.getElementById(\"authPasswordInput\");\r\n\r\n      if (role === 'ADMIN') {\r\n        btnAdmin.className = \"btn btn-primary\";\r\n        btnCand.className = \"btn btn-outline\";\r\n        if (candGroup) candGroup.style.display = \"none\";\r\n        if (passInp) passInp.placeholder = \"सुपर एडमिन पासवर्ड दर्ज करें (उदा. admin2026)\";\r\n      } else {\r\n        btnAdmin.className = \"btn btn-outline\";\r\n        btnCand.className = \"btn btn-primary\";\r\n        if (candGroup) candGroup.style.display = \"block\";\r\n        if (passInp) passInp.placeholder = \"प्रत्याशी पासवर्ड दर्ज करें\";\r\n      }\r\n    }\r\n\r\n        function verifyAuthLogin() {\r\n      var pass = document.getElementById(\"authPasswordInput\").value.trim();\r\n      var errEl = document.getElementById(\"authErrorMsg\");\r\n      var btn = document.querySelector(\"#authReviewModal .btn-success\");\r\n\r\n      if (!pass) {\r\n        errEl.textContent = \"⚠️ कृपया पासवर्ड दर्ज करें!\";\r\n        errEl.style.display = \"block\";\r\n        return;\r\n      }\r\n\r\n      var uName = (currentAuthRole === 'ADMIN') ? 'admin' : document.getElementById(\"authCandUsernameSelect\").value;\r\n      if (!uName) uName = 'admin';\r\n\r\n      // 1. Instant Verification: Super Admin default passwords\r\n      if (currentAuthRole === 'ADMIN' && (pass === \"admin2026\" || pass === \"bhinai2026\" || pass === \"BHINAI123\")) {\r\n        currentAuthUser = { username: \"admin\", role: \"SUPER_ADMIN\", fullName: \"मुख्य व्यवस्थापक (Super Admin)\" };\r\n        sessionStorage.setItem(\"cs_auth_user\", JSON.stringify(currentAuthUser));\r\n        document.getElementById(\"authReviewModal\").style.display = \"none\";\r\n        applyAuthPermissions(currentAuthUser);\r\n        showToast(\"✅ सुपर एडमिन प्रमाणीकरण सफल!\", \"success\");\r\n        return;\r\n      }\r\n\r\n      // 2. Instant Verification: Candidate from loaded list\r\n      if (currentAuthRole === 'CANDIDATE' && Array.isArray(candidates) && candidates.length > 0) {\r\n        var matchCand = candidates.find(function(x) { return x.username.toLowerCase() === uName.toLowerCase(); });\r\n        if (matchCand && matchCand.password && matchCand.password.trim() === pass) {\r\n          currentAuthUser = { username: uName, role: \"CANDIDATE\", fullName: (matchCand.candidate_name || uName) };\r\n          sessionStorage.setItem(\"cs_auth_user\", JSON.stringify(currentAuthUser));\r\n          document.getElementById(\"authReviewModal\").style.display = \"none\";\r\n          applyAuthPermissions(currentAuthUser);\r\n          showToast(\"✅ प्रत्याशी '\" + (matchCand.candidate_name || uName) + \"' प्रमाणीकरण सफल!\", \"success\");\r\n          return;\r\n        }\r\n      }\r\n\r\n      // 3. Fallback: Online verification via API\r\n      errEl.style.display = \"none\";\r\n      if (btn) { btn.disabled = true; btn.textContent = \"⏳ प्रमाणीकरण हो रहा है...\"; }\r\n\r\n      function onAuthSuccess(res) {\r\n        if (btn) { btn.disabled = false; btn.textContent = \"🔓 प्रवेश करें (Verify & Login)\"; }\r\n        if (res && res.success) {\r\n          currentAuthUser = res.user || { username: uName, role: (currentAuthRole === 'ADMIN' ? 'SUPER_ADMIN' : 'CANDIDATE'), fullName: uName };\r\n          sessionStorage.setItem(\"cs_auth_user\", JSON.stringify(currentAuthUser));\r\n          document.getElementById(\"authReviewModal\").style.display = \"none\";\r\n          applyAuthPermissions(currentAuthUser);\r\n          showToast(\"✅ \" + (currentAuthUser.fullName || currentAuthUser.username) + \" प्रमाणीकरण सफल!\", \"success\");\r\n        } else {\r\n          errEl.textContent = \"❌ \" + ((res && res.error) ? res.error : \"गलत पासवर्ड! कृपया सही पासवर्ड दर्ज करें।\");\r\n          errEl.style.display = \"block\";\r\n        }\r\n      }\r\n\r\n      function onAuthFailure(err) {\r\n        if (btn) { btn.disabled = false; btn.textContent = \"🔓 प्रवेश करें (Verify & Login)\"; }\r\n        errEl.textContent = \"❌ गलत पासवर्ड! (\" + err + \")\";\r\n        errEl.style.display = \"block\";\r\n      }\r\n\r\n      if (typeof google !== \"undefined\" && google.script && google.script.run) {\r\n        google.script.run.withSuccessHandler(onAuthSuccess).withFailureHandler(onAuthFailure).apiLogin(uName, pass);\r\n      } else {\r\n        var base = window.location.href.split(\"?\")[0];\r\n        fetch(base + \"?action=login&username=\" + encodeURIComponent(uName) + \"&password=\" + encodeURIComponent(pass))\r\n          .then(function(r){ return r.json(); })\r\n          .then(onAuthSuccess)\r\n          .catch(onAuthFailure);\r\n      }\r\n    }\r\n\r\n    function applyAuthPermissions(auth) {\r\n      var badge = document.getElementById(\"authRoleBadge\");\r\n      var logoutBtn = document.getElementById(\"logoutBtn\");\r\n      if (logoutBtn) logoutBtn.style.display = \"inline-flex\";\r\n\r\n      if (badge) {\r\n        if (auth.role === 'SUPER_ADMIN') {\r\n          badge.textContent = \"👑 सुपर एडमिन\";\r\n          badge.style.background = \"#10b981\";\r\n        } else {\r\n          badge.textContent = \"👤 \" + auth.name;\r\n          badge.style.background = \"#3b82f6\";\r\n        }\r\n      }\r\n\r\n      var listPanel = document.getElementById(\"candidatesListPanel\") || document.querySelector(\".panel-card:first-child\");\r\n      if (auth.role === 'CANDIDATE') {\r\n        if (listPanel) listPanel.style.display = \"none\";\r\n        var delBtn = document.getElementById(\"deleteCandBtn\");\r\n        if (delBtn) delBtn.style.display = \"none\";\r\n        var newBtn = document.getElementById(\"newCandBtn\");\r\n        if (newBtn) newBtn.style.display = \"none\";\r\n        selectCandidate(auth.username);\r\n      } else {\r\n        if (listPanel) listPanel.style.display = \"block\";\r\n        var newBtn = document.getElementById(\"newCandBtn\");\r\n        if (newBtn) newBtn.style.display = \"inline-flex\";\r\n      }\r\n    }\r\n\r\n    function logoutWebApp() {\r\n      sessionStorage.removeItem(\"cs_auth_user\");\r\n      currentAuthUser = null;\r\n      showToast(\"🚪 लॉगआउट सफल!\");\r\n      showAuthModal();\r\n    }\r\n\r\n    function showToast(msg, type) {\r\n      var t = document.getElementById(\"toastBox\");\r\n      t.textContent = msg;\r\n      t.className = \"toast \" + (type || \"success\");\r\n      t.style.display = \"block\";\r\n      clearTimeout(window._toastTimer);\r\n      window._toastTimer = setTimeout(function() { t.style.display = \"none\"; }, 4000);\r\n    }\r\n  </script>\r\n\r\n  <!-- 🔐 Security Review & Authentication Modal -->\r\n  <div id=\"authReviewModal\" style=\"display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(15,23,42,0.85); z-index:99999; backdrop-filter:blur(5px); align-items:center; justify-content:center;\">\r\n    <div style=\"background:#ffffff; border-radius:14px; width:92%; max-width:440px; padding:24px; box-shadow:0 14px 40px rgba(0,0,0,0.35); border:2px solid #2563eb; text-align:center;\">\r\n      <div style=\"font-size:2.8rem; margin-bottom:8px;\">🔐</div>\r\n      <h3 style=\"font-size:1.25rem; font-weight:800; color:#1e3a8a; margin-bottom:4px;\">निर्वाचन कमान कक्ष प्रमाणीकरण</h3>\r\n      <p style=\"font-size:0.82rem; color:#64748b; margin-bottom:16px;\">सुरक्षित पहुंच हेतु कृपया व्यवस्थापक (Admin) या प्रत्याशी पासवर्ड दर्ज करें:</p>\r\n\r\n      <div style=\"display:flex; gap:8px; margin-bottom:16px;\">\r\n        <button type=\"button\" id=\"authRoleAdminBtn\" class=\"btn btn-primary\" onclick=\"setAuthRole('ADMIN')\" style=\"flex:1; justify-content:center; padding:8px;\">👑 सुपर एडमिन</button>\r\n        <button type=\"button\" id=\"authRoleCandBtn\" class=\"btn btn-outline\" onclick=\"setAuthRole('CANDIDATE')\" style=\"flex:1; justify-content:center; padding:8px;\">👤 प्रत्याशी</button>\r\n      </div>\r\n\r\n      <div id=\"authCandSelectGroup\" style=\"display:none; margin-bottom:12px; text-align:left;\">\r\n        <label style=\"font-size:11px; font-weight:700; color:#334155;\">प्रत्याशी चुनें *:</label>\r\n        <select id=\"authCandUsernameSelect\" class=\"form-select\" style=\"margin-top:4px;\">\r\n        </select>\r\n      </div>\r\n\r\n      <div style=\"margin-bottom:16px; text-align:left;\">\r\n        <label style=\"font-size:11px; font-weight:700; color:#334155;\">लॉगिन पासवर्ड *:</label>\r\n        <input type=\"password\" id=\"authPasswordInput\" class=\"form-input\" placeholder=\"पासवर्ड दर्ज करें (उदा. admin2026)\" style=\"margin-top:4px; font-size:1rem; padding:10px;\" onkeydown=\"if(event.key==='Enter') verifyAuthLogin()\">\r\n        <div id=\"authErrorMsg\" style=\"color:#dc2626; font-size:0.8rem; font-weight:700; margin-top:6px; display:none;\"></div>\r\n      </div>\r\n\r\n      <div style=\"display:flex; gap:10px;\">\r\n        <button type=\"button\" class=\"btn btn-success\" onclick=\"verifyAuthLogin()\" style=\"flex:1; padding:10px; font-size:0.95rem; justify-content:center;\">\r\n          🔓 प्रवेश करें (Verify & Login)\r\n        </button>\r\n      </div>\r\n      <small style=\"display:block; margin-top:12px; color:#059669; font-size:0.75rem; font-weight:600;\">✨ गूगल शीट 'Users' डेटाबेस द्वारा अधिकृत एवं सत्यापित</small>\r\n    </div>\r\n  </div>\r\n\r\n</body>\r\n</html>";
}
