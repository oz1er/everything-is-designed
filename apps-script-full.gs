var SHEET_NAME = 'Sheet1';
var MAX_SEATS = 200;

// Column positions (1-based)
var EMAIL_COL = 4;
var UTR_COL = 9;

// Generate unique registration ID
function generateRegistrationId() {
  var uuid = Utilities.getUuid()
    .replace(/-/g, '')
    .substring(0, 8)
    .toUpperCase();

  var year = new Date().getFullYear();

  return 'EVT-' + year + '-' + uuid;
}

// Generate unique QR token
function generateQRToken() {
  return Utilities.getUuid().replace(/-/g, '') +
         Utilities.getUuid().replace(/-/g, '');
}

function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error('Sheet not found: ' + SHEET_NAME);
    }

    // Honeypot protection
    if (e.parameter.website && e.parameter.website.trim() !== '') {
      return jsonOut({ ok: true });
    }

    // Read registration fields
    var name = (e.parameter.name || '').trim();
    var phone = (e.parameter.phone || '').trim();
    var email = (e.parameter.email || '').trim().toLowerCase();
    var college = (e.parameter.college || '').trim();
    var otherCollege = (e.parameter['other-college'] || '').trim();
    var roll = (e.parameter.roll || '').trim();
    var utr = (e.parameter.utr || '').trim();
    var paymentMethod = (e.parameter.paymentMethod || 'UPI').trim();

    // Basic validation
    if (!name || !phone || !email || !college || !roll) {
      return jsonOut({ ok: false, error: 'invalid' });
    }

    var finalCollege = (college === 'Other' && otherCollege)
      ? otherCollege
      : college;

    // Seat capacity check
    var lastRow = sheet.getLastRow();

    if (lastRow >= MAX_SEATS + 1) {
      return jsonOut({ ok: false, error: 'full' });
    }

    // Read existing registrations
    var data = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues()
      : [];

    // Duplicate checks
    for (var i = 0; i < data.length; i++) {

      if (String(data[i][EMAIL_COL - 1]).trim().toLowerCase() === email) {
        return jsonOut({ ok: false, error: 'duplicate_email' });
      }

      if (
        paymentMethod === 'UPI' &&
        utr &&
        String(data[i][UTR_COL - 1]).trim() === utr
      ) {
        return jsonOut({ ok: false, error: 'duplicate_utr' });
      }
    }

    // Generate unique identifiers
    var registrationId = generateRegistrationId();
    var qrToken = generateQRToken();

    // Registration starts as not checked in
    var entryStatus = 'NO';

    // Save registration
    var timestamp = new Date();

    
sheet.appendRow([
  timestamp,
  name,
  phone,
  email,
  finalCollege,
  otherCollege,
  roll,
  paymentMethod,
  utr,
  '',
  registrationId,
  qrToken,
  entryStatus,
  'PENDING'
]);

// Get the newly created row
var newRow = sheet.getLastRow();

var emailSent = false;

try {
  sendRegistrationEmail(
    name,
    email,
    registrationId,
    qrToken
  );

  sheet.getRange(newRow, 14).setValue('SENT');
  emailSent = true;

} catch (mailError) {

  sheet.getRange(newRow, 14).setValue('FAILED');

  console.error('Email error: ' + mailError.message);
}

return jsonOut({
  ok: true,
  registrationId: registrationId,
  emailSent: emailSent
});

    // Return registration ID to frontend
    // QR token remains server-side
    return jsonOut({
      ok: true,
      registrationId: registrationId
    });

  } catch (err) {

    return jsonOut({
      ok: false,
      error: 'server_error',
      msg: err.message
    });

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// JSON response helper
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


function backfillExistingRegistrations() {

  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Sheet1');

  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log('No registrations found.');
    return;
  }

  // Read columns K, L and M
  var range = sheet.getRange(2, 11, lastRow - 1, 3);
  var records = range.getValues();

  var updated = 0;

  for (var i = 0; i < records.length; i++) {

    // Skip empty registration rows
    var email = sheet.getRange(i + 2, 4).getValue();

    if (!email) continue;

    // Generate ID only if missing
    if (!records[i][0]) {
      records[i][0] = generateRegistrationId();
    }

    // Generate QR token only if missing
    if (!records[i][1]) {
      records[i][1] = generateQRToken();
    }

    // Set default entry status only if missing
    if (!records[i][2]) {
      records[i][2] = 'NO';
    }

    updated++;
  }

  // Update K:M without touching other columns
  range.setValues(records);

  SpreadsheetApp.flush();

  Logger.log(updated + ' registrations processed.');
}


/**
 * Generate QR image from the participant's unique token.
 * The QR contains only the token, not personal information.
 */
function generateQRImage(token) {
  var qrUrl = 'https://quickchart.io/qr?size=300&margin=2&text='
    + encodeURIComponent(token);

  var response = UrlFetchApp.fetch(qrUrl);

  if (response.getResponseCode() !== 200) {
    throw new Error('QR image generation failed.');
  }

  return response.getBlob().setName('entry-qr.png');
}


/**
 * Send registration email with an embedded QR code.
 * (UI restyled to match the website: cream background, blue/green/pink
 * accents, chunky offset borders. Email clients ignore web fonts and
 * box-shadow, so this uses tables, inline styles and solid borders.)
 */
function sendRegistrationEmail(name, email, registrationId, token) {

  var qrBlob = generateQRImage(token);

  var subject = 'You\'re registered! | ' + registrationId;

  var FONT = "'Bricolage Grotesque','Trebuchet MS',Arial,sans-serif";
  var MONO = "'Space Mono','Courier New',monospace";

  var htmlBody =
    '<div style="background:#FBF6E9;padding:28px 12px;font-family:' + FONT + ';color:#171717;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;margin:0 auto;">' +

      // Header tag
      '<tr><td style="padding:0 0 18px 4px;">' +
        '<span style="display:inline-block;background:#1FA85A;color:#ffffff;font-family:' + MONO + ';' +
        'font-size:12px;font-weight:700;letter-spacing:1px;padding:6px 14px;">REGISTRATION CONFIRMED</span>' +
      '</td></tr>' +

      // Hero
      '<tr><td style="padding:0 4px 22px 4px;">' +
        '<div style="font-size:42px;line-height:1;font-weight:800;letter-spacing:-2px;color:#171717;">You\'re registered!</div>' +
        '<div style="font-size:42px;line-height:1.1;font-weight:800;letter-spacing:-2px;color:#2F5BEA;">everything is designed.</div>' +
      '</td></tr>' +

      // Main card
      '<tr><td style="background:#ffffff;border-radius:28px;padding:32px 28px;border-right:6px solid #e6e0cf;border-bottom:6px solid #e6e0cf;">' +
        '<div style="font-size:20px;font-weight:800;margin-bottom:8px;">Hello ' + escapeHtml(name) + ',</div>' +
        '<div style="font-size:15px;color:#5b574d;margin-bottom:22px;">Your registration has been received. ' +
        'We\'re looking forward to seeing you at the event.</div>' +

        // Reg ID
        '<div style="background:#FBF6E9;border-radius:16px;padding:16px 18px;margin-bottom:22px;">' +
          '<div style="font-family:' + MONO + ';font-size:11px;letter-spacing:1.5px;color:#5b574d;">REGISTRATION ID</div>' +
          '<div style="font-family:' + MONO + ';font-size:20px;font-weight:700;margin-top:4px;word-break:break-all;">' +
          escapeHtml(registrationId) + '</div>' +
        '</div>' +

        // QR
        '<div style="text-align:center;padding:6px 0 4px 0;">' +
          '<div style="font-size:16px;font-weight:800;margin-bottom:12px;">Your unique entry QR code</div>' +
          '<div style="display:inline-block;background:#ffffff;border:3px solid #171717;border-radius:20px;padding:12px;">' +
            '<img src="cid:entryQR" width="220" height="220" alt="Your entry QR code" style="display:block;">' +
          '</div>' +
          '<div style="font-family:' + MONO + ';font-size:11px;color:#5b574d;margin-top:12px;">' +
          'Show this QR at the registration desk</div>' +
        '</div>' +
      '</td></tr>' +

      '<tr><td style="height:16px;line-height:16px;font-size:0;">&nbsp;</td></tr>' +

      // Event details (blue card)
      '<tr><td style="background:#2F5BEA;color:#ffffff;border-radius:28px;padding:24px 28px;border-right:6px solid #1f43b8;border-bottom:6px solid #1f43b8;">' +
        '<div style="font-family:' + MONO + ';font-size:11px;letter-spacing:1.5px;opacity:.85;">EVENT DETAILS</div>' +
        '<div style="font-size:18px;font-weight:800;margin-top:6px;">Thu, 12 Nov &middot; 10 AM &middot; MJCET</div>' +
        '<div style="font-size:14px;margin-top:6px;opacity:.9;">Carry your college ID to the venue.</div>' +
      '</td></tr>' +

      '<tr><td style="height:16px;line-height:16px;font-size:0;">&nbsp;</td></tr>' +

      // Notice
      '<tr><td style="background:#FFC93C;border-radius:20px;padding:16px 20px;font-size:14px;">' +
        '<strong>Important:</strong> entry is subject to payment verification. Please keep this email for check-in.' +
      '</td></tr>' +

      // Footer
      '<tr><td style="padding:24px 4px 0 4px;text-align:center;font-family:' + MONO + ';font-size:11px;color:#5b574d;">' +
        'Zainab 93930 59590 &middot; Aman 97019 46533 &middot; Ozier 89786 36483<br><br>' +
        'IEEE Computational Intelligence Society &ndash; MJCET' +
      '</td></tr>' +

    '</table></div>';

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: 'Your registration has been received. Registration ID: '
          + registrationId + '. Please retain your QR entry pass.',
    htmlBody: htmlBody,
    inlineImages: {
      entryQR: qrBlob
    },
    name: 'IEEE CIS MJCET'
  });
}


/**
 * Escape user-provided text before inserting into HTML.
 */
function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function authorizeEmailAndQR() {
  Logger.log('Email quota: ' + MailApp.getRemainingDailyQuota());
  var response = UrlFetchApp.fetch(
    'https://quickchart.io/qr?text=TEST&size=100'
  );
  Logger.log('QR service status: ' + response.getResponseCode());
}


function doGet(e) {

  var callback = e.parameter.callback || '';

  // Allow only a simple JavaScript callback name
  if (!/^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(callback)) {
    return ContentService.createTextOutput('Invalid callback');
  }

  var result;

  try {

    var token = (e.parameter.token || '').trim();
    var staffPin = (e.parameter.pin || '').trim();

    // Configure this PIN in Apps Script Script Properties
    var correctPin = PropertiesService
      .getScriptProperties()
      .getProperty('CHECKIN_PIN');

    if (!correctPin || staffPin !== correctPin) {
      throw new Error('unauthorized');
    }

    if (!/^[a-f0-9]{64}$/i.test(token)) {
      throw new Error('invalid_qr');
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {

      var sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName('Sheet1');

      var lastRow = sheet.getLastRow();

      if (lastRow < 2) {
        throw new Error('not_found');
      }

      // L = QR_Token
      var tokens = sheet.getRange(
        2, 12, lastRow - 1, 1
      ).getValues();

      var matchedRow = -1;

      for (var i = 0; i < tokens.length; i++) {
        if (String(tokens[i][0]).trim() === token) {
          matchedRow = i + 2;
          break;
        }
      }

      if (matchedRow === -1) {
        throw new Error('not_found');
      }

      // M = Entry_Status
      var status = sheet.getRange(matchedRow, 13)
        .getValue();

      var name = sheet.getRange(matchedRow, 2)
        .getValue();

      if (String(status).toUpperCase() === 'YES') {

        result = {
          ok: false,
          status: 'already_checked_in',
          name: name,
          entryTime: sheet.getRange(matchedRow, 15).getValue()
        };

      } else {

        var entryTime = new Date();

        sheet.getRange(matchedRow, 13).setValue('YES');
        sheet.getRange(matchedRow, 15).setValue(entryTime);

        SpreadsheetApp.flush();

        result = {
          ok: true,
          status: 'success',
          name: name,
          registrationId: sheet.getRange(matchedRow, 11).getValue(),
          entryTime: entryTime.toISOString()
        };

      }

    } finally {
      lock.releaseLock();
    }

  } catch (err) {

    var errorCode = err.message;

    result = {
      ok: false,
      status: errorCode === 'unauthorized'
        ? 'unauthorized'
        : errorCode === 'invalid_qr'
          ? 'invalid_qr'
          : errorCode === 'not_found'
            ? 'not_found'
            : 'server_error'
    };

  }

  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(result) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}


function doGet(e) {

  var callback = e.parameter.callback || '';

  // Allow only a simple JavaScript callback name
  if (!/^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(callback)) {
    return ContentService.createTextOutput('Invalid callback');
  }

  var result;

  try {

    var token = (e.parameter.token || '').trim();
    var staffPin = (e.parameter.pin || '').trim();

    // Configure this PIN in Apps Script Script Properties
    var correctPin = PropertiesService
      .getScriptProperties()
      .getProperty('CHECKIN_PIN');

    if (!correctPin || staffPin !== correctPin) {
      throw new Error('unauthorized');
    }

    if (!/^[a-f0-9]{64}$/i.test(token)) {
      throw new Error('invalid_qr');
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {

      var sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName('Sheet1');

      var lastRow = sheet.getLastRow();

      if (lastRow < 2) {
        throw new Error('not_found');
      }

      // L = QR_Token
      var tokens = sheet.getRange(
        2, 12, lastRow - 1, 1
      ).getValues();

      var matchedRow = -1;

      for (var i = 0; i < tokens.length; i++) {
        if (String(tokens[i][0]).trim() === token) {
          matchedRow = i + 2;
          break;
        }
      }

      if (matchedRow === -1) {
        throw new Error('not_found');
      }

      // M = Entry_Status
      var status = sheet.getRange(matchedRow, 13)
        .getValue();

      var name = sheet.getRange(matchedRow, 2)
        .getValue();

      if (String(status).toUpperCase() === 'YES') {

        result = {
          ok: false,
          status: 'already_checked_in',
          name: name,
          entryTime: sheet.getRange(matchedRow, 15).getValue()
        };

      } else {

        var entryTime = new Date();

        sheet.getRange(matchedRow, 13).setValue('YES');
        sheet.getRange(matchedRow, 15).setValue(entryTime);

        SpreadsheetApp.flush();

        result = {
          ok: true,
          status: 'success',
          name: name,
          registrationId: sheet.getRange(matchedRow, 11).getValue(),
          entryTime: entryTime.toISOString()
        };

      }

    } finally {
      lock.releaseLock();
    }

  } catch (err) {

    var errorCode = err.message;

    result = {
      ok: false,
      status: errorCode === 'unauthorized'
        ? 'unauthorized'
        : errorCode === 'invalid_qr'
          ? 'invalid_qr'
          : errorCode === 'not_found'
            ? 'not_found'
            : 'server_error'
    };

  }

  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(result) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
