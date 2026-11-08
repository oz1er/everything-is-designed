/**
 * Drop-in replacement for sendRegistrationEmail() in your Apps Script.
 * Styled to match the website (cream background, blue/green/pink accents,
 * chunky offset "shadows"). Email clients ignore web fonts and box-shadow,
 * so this uses table layout, inline styles and solid borders for the offset effect.
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
    inlineImages: { entryQR: qrBlob },
    name: 'IEEE CIS MJCET'
  });
}
