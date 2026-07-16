/**
 * SASCE WIL AFRICA CONFERENCE 2026 — form submission backend.
 * Handles both the Abstract submission form and the Registration form,
 * writing to two separate sheet tabs, supports editing a submission later
 * via a private "edit link" sent to the submitter's email, and serves
 * data to admin.html for the admin dashboard.
 *
 * ===================== SETUP (read this fully) =====================
 * 1. Create a Google Sheet (this will store your submissions).
 * 2. Extensions → Apps Script.
 * 3. Select ALL existing code in the editor and DELETE it, then paste
 *    this whole file in.
 * 4. Change ADMIN_KEY below to your own secret password.
 * 5. Save (Ctrl+S / Cmd+S).
 * 6. Click Deploy → Manage deployments.
 *    - If a deployment already exists, click the pencil (Edit) icon on it,
 *      set "Version" to "New version", and click Deploy. This keeps the
 *      same Web app URL — you do NOT need to update SCRIPT_URL again.
 *    - If none exists yet, click Deploy → New deployment → type "Web app",
 *      set "Execute as: Me" and "Who has access: Anyone", then Deploy and
 *      copy the URL into SCRIPT_URL in abstract-submission.html, index.html,
 *      and admin.html.
 * 7. IMPORTANT: every time you change this file, you must repeat step 6
 *    (Manage deployments → Edit → New version → Deploy) or the live form
 *    keeps running the old code. Saving the file alone is not enough.
 * =====================================================================
 *
 * If "Abstracts" or "Registrations" tabs already exist from earlier testing
 * with an older header layout, delete those tabs before your next test
 * submission — this script recreates them with the correct headers the
 * first time it writes to a sheet that doesn't exist yet.
 */

var ABSTRACTS_SHEET_NAME = 'Abstracts';
var REGISTRATIONS_SHEET_NAME = 'Registrations';
var UPLOADS_FOLDER_NAME = 'WIL Africa 2026 — Proof of Payment uploads';

// Change this to your own secret before deploying. Anyone with this key
// can view all submissions via admin.html, so keep it private.
var ADMIN_KEY = 'CHANGE_ME_ADMIN_KEY_2026';

var ABSTRACT_HEADERS = [
  'Timestamp', 'Submission ID', 'Edit Link', 'Corresponding author', 'Email', 'Phone', 'Country',
  'Institution', 'Role', 'ORCID iD', 'Co-authors',
  'Abstract title', 'Sub-theme', 'Submission type', 'Presentation format',
  'Keywords', 'Abstract text', 'Word count',
  'Ethical clearance', 'Funding / acknowledgements', 'Attendance mode',
  'Journal interest', 'Notes', 'Declaration accepted', 'Raw Data (JSON)'
];

var ABSTRACT_FIELD_MAP = {
  'Timestamp': 'timestamp',
  'Corresponding author': 'corrName',
  'Email': 'corrEmail',
  'Phone': 'corrPhone',
  'Country': 'corrCountry',
  'Institution': 'corrOrg',
  'Role': 'corrRole',
  'ORCID iD': 'corrOrcid',
  'Co-authors': 'coAuthors',
  'Abstract title': 'title',
  'Sub-theme': 'subtheme',
  'Submission type': 'submissionType',
  'Presentation format': 'format',
  'Keywords': 'keywords',
  'Abstract text': 'abstractText',
  'Word count': 'wordCount',
  'Ethical clearance': 'ethics',
  'Funding / acknowledgements': 'funding',
  'Attendance mode': 'attendance',
  'Journal interest': 'journalInterest',
  'Notes': 'notes',
  'Declaration accepted': 'declaration'
};

// NOTE: 'Official Address', 'Research Publishing Workshop Attendance',
// 'Hotel Arrival Time', 'Poster Presentation' and 'Plenary Presentation' are
// kept here (though the form no longer collects them) so older rows already
// in your sheet keep displaying correctly. New submissions will just leave
// these columns blank. New questions were appended at the end, before
// 'Raw Data (JSON)', instead of reordering existing columns — reordering
// would misalign any registrations already recorded in the sheet.
var REGISTRATION_HEADERS = [
  'Timestamp', 'Submission ID', 'Edit Link', 'Email', 'Title', 'Name and Surname', 'Institution Organisation',
  'Designation Position', 'Official Address', 'Office Telephone Number',
  'Official Email', 'Cellphone Number',
  'Research Publishing Workshop Attendance', '4IR CoE WIL Site Visit Attendance',
  'Hotel Arrival Time', 'Hotel Name', 'Conference Registration Arrival Time',
  'Other Arrival Details', 'Poster Presentation', 'Plenary Presentation',
  'Accommodation Booking Method', 'Payment Confirmation Available',
  'Payment Reference Number', 'Payment Date', 'Proof of Payment File URL',
  'Pre-conference Programme Attendance', 'Open Poster/Exhibition',
  'Subtheme Breakaway or Workshop', 'Plenary Theme Paper', 'Registration Fee Category',
  'Raw Data (JSON)'
];

var REGISTRATION_FIELD_MAP = {
  'Timestamp': 'timestamp',
  'Email': 'regEmail',
  'Title': 'title',
  'Name and Surname': 'fullName',
  'Institution Organisation': 'institution',
  'Designation Position': 'designation',
  'Official Address': 'officialAddress',
  'Office Telephone Number': 'officePhone',
  'Official Email': 'officialEmail',
  'Cellphone Number': 'cellphone',
  'Research Publishing Workshop Attendance': 'researchWorkshop',
  '4IR CoE WIL Site Visit Attendance': 'siteVisit',
  'Hotel Arrival Time': 'hotelArrivalTime',
  'Hotel Name': 'hotelName',
  'Conference Registration Arrival Time': 'regDeskArrival',
  'Other Arrival Details': 'otherArrivalDetails',
  'Poster Presentation': 'posterPresentation',
  'Plenary Presentation': 'plenaryPresentation',
  'Accommodation Booking Method': 'accommodationMethod',
  'Payment Confirmation Available': 'paymentConfirmation',
  'Payment Reference Number': 'paymentReference',
  'Payment Date': 'paymentDate',
  'Pre-conference Programme Attendance': 'preConferenceProgramme',
  'Open Poster/Exhibition': 'posterExhibition',
  'Subtheme Breakaway or Workshop': 'subthemeBreakaway',
  'Plenary Theme Paper': 'plenaryPaper',
  'Registration Fee Category': 'feeCategory'
};

function getSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildEditUrl_(pageUrl, id) {
  if (!pageUrl) return '';
  return pageUrl + (pageUrl.indexOf('?') === -1 ? '?' : '&') + 'edit=' + encodeURIComponent(id);
}

function buildRow_(headers, fieldMap, data, id, rawObj, extraHeaderValues) {
  extraHeaderValues = extraHeaderValues || {};
  return headers.map(function (h) {
    if (h === 'Submission ID') return id;
    if (h === 'Raw Data (JSON)') return JSON.stringify(rawObj);
    if (Object.prototype.hasOwnProperty.call(extraHeaderValues, h)) return extraHeaderValues[h];
    var key = fieldMap[h];
    if (!key) return '';
    var v = data[key];
    return v === undefined || v === null ? '' : v;
  });
}

function findRowById_(sheet, headers, id) {
  var idCol = headers.indexOf('Submission ID') + 1;
  if (idCol === 0) return -1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2;
  }
  return -1;
}

function sheetToObjects_(sheet, headers) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function saveProofOfPayment_(base64, fileName, mimeType) {
  if (!base64) return '';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var parentFolders = DriveApp.getFileById(ss.getId()).getParents();
  var searchRoot = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
  var existing = searchRoot.getFoldersByName(UPLOADS_FOLDER_NAME);
  var folder = existing.hasNext() ? existing.next() : searchRoot.createFolder(UPLOADS_FOLDER_NAME);

  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName || 'proof-of-payment');
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function sendEditEmail_(email, editUrl, kind) {
  if (!email || !editUrl) return;
  try {
    var subject = 'SASCE WIL AFRICA CONFERENCE 2026 — your ' + kind + ' submission';
    var body = 'Thank you for your ' + kind + ' submission to the SASCE WIL AFRICA CONFERENCE 2026.\n\n' +
      'If you need to update your details later, use this link:\n' + editUrl + '\n\n' +
      'Keep this link private — anyone with it can edit your submission.';
    MailApp.sendEmail(email, subject, body);
  } catch (err) {
    // Email is a nice-to-have; a failed send should never fail the submission.
  }
}

function upsertAbstract_(data) {
  var sheet = getSheet_(ABSTRACTS_SHEET_NAME, ABSTRACT_HEADERS);
  var isEdit = !!data.editId;
  var id = data.editId || Utilities.getUuid();
  var editUrl = buildEditUrl_(data.pageUrl, id);

  var raw = {};
  for (var k in data) { if (k !== 'editId') raw[k] = data[k]; }

  var row = buildRow_(ABSTRACT_HEADERS, ABSTRACT_FIELD_MAP, data, id, raw, {
    'Edit Link': editUrl
  });

  if (isEdit) {
    var rowNum = findRowById_(sheet, ABSTRACT_HEADERS, data.editId);
    if (rowNum === -1) throw new Error('Submission not found for the given edit link.');
    sheet.getRange(rowNum, 1, 1, ABSTRACT_HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
    sendEditEmail_(data.corrEmail, editUrl, 'abstract');
  }
  return { id: id, editUrl: editUrl };
}

function upsertRegistration_(data) {
  var sheet = getSheet_(REGISTRATIONS_SHEET_NAME, REGISTRATION_HEADERS);
  var isEdit = !!data.editId;
  var id = data.editId || Utilities.getUuid();
  var editUrl = buildEditUrl_(data.pageUrl, id);

  var proofUrl = '';
  if (data.proofOfPaymentBase64) {
    proofUrl = saveProofOfPayment_(data.proofOfPaymentBase64, data.proofOfPaymentName, data.proofOfPaymentType);
  } else if (isEdit) {
    var existingRowNum = findRowById_(sheet, REGISTRATION_HEADERS, data.editId);
    if (existingRowNum !== -1) {
      var idxRaw = REGISTRATION_HEADERS.indexOf('Raw Data (JSON)') + 1;
      var existingRaw = sheet.getRange(existingRowNum, idxRaw).getValue();
      try {
        var parsed = JSON.parse(existingRaw);
        proofUrl = parsed.proofOfPaymentUrl || '';
      } catch (err) { /* ignore malformed legacy row */ }
    }
  }

  var raw = {};
  for (var k in data) { if (k !== 'editId' && k !== 'proofOfPaymentBase64') raw[k] = data[k]; }
  raw.proofOfPaymentUrl = proofUrl;

  var row = buildRow_(REGISTRATION_HEADERS, REGISTRATION_FIELD_MAP, data, id, raw, {
    'Proof of Payment File URL': proofUrl,
    'Edit Link': editUrl
  });

  if (isEdit) {
    var rowNum = findRowById_(sheet, REGISTRATION_HEADERS, data.editId);
    if (rowNum === -1) throw new Error('Submission not found for the given edit link.');
    sheet.getRange(rowNum, 1, 1, REGISTRATION_HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
    sendEditEmail_(data.regEmail, editUrl, 'registration');
  }
  return { id: id, editUrl: editUrl };
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var result;

    if (data.formType === 'registration') {
      result = upsertRegistration_(data);
    } else {
      result = upsertAbstract_(data);
    }

    return jsonOutput_({ result: 'success', submissionId: result.id, editUrl: result.editUrl });

  } catch (err) {
    return jsonOutput_({ result: 'error', message: err.message });
  }
}

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};

  // Admin dashboard data feed.
  if (params.action === 'admin') {
    if (!ADMIN_KEY || params.key !== ADMIN_KEY) {
      return jsonOutput_({ result: 'error', message: 'Unauthorized' });
    }
    var abstractsSheet = getSheet_(ABSTRACTS_SHEET_NAME, ABSTRACT_HEADERS);
    var registrationsSheet = getSheet_(REGISTRATIONS_SHEET_NAME, REGISTRATION_HEADERS);
    return jsonOutput_({
      result: 'success',
      abstracts: sheetToObjects_(abstractsSheet, ABSTRACT_HEADERS),
      registrations: sheetToObjects_(registrationsSheet, REGISTRATION_HEADERS)
    });
  }

  // Single-submission lookup, used by the edit-prefill flow.
  if (params.id && params.formType) {
    var isRegistration = params.formType === 'registration';
    var headers = isRegistration ? REGISTRATION_HEADERS : ABSTRACT_HEADERS;
    var sheet = getSheet_(
      isRegistration ? REGISTRATIONS_SHEET_NAME : ABSTRACTS_SHEET_NAME,
      headers
    );
    var rowNum = findRowById_(sheet, headers, params.id);
    if (rowNum === -1) {
      return jsonOutput_({ result: 'error', message: 'Submission not found.' });
    }
    var idxRaw = headers.indexOf('Raw Data (JSON)') + 1;
    var rawValue = sheet.getRange(rowNum, idxRaw).getValue();
    var data;
    try { data = JSON.parse(rawValue); } catch (err) { data = {}; }
    return jsonOutput_({ result: 'success', data: data });
  }

  return jsonOutput_({ status: 'WIL Africa Conference form endpoint is live' });
}
