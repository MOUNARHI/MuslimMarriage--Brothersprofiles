/**
 * MATCH REQUEST BACKEND
 * --------------------------------------------------------------
 * Deploy this as a standalone Apps Script (script.google.com),
 * bound to a NEW Google Sheet (or a "Match Requests" tab inside
 * your existing consolidation sheet).
 *
 * SETUP:
 * 1. Create/open the target Google Sheet.
 * 2. Extensions > Apps Script. Paste this file in, replacing Code.gs.
 * 3. Update SHEET_NAME and ADMIN_EMAIL below.
 * 4. Click Deploy > New deployment > type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5. Copy the Web app URL it gives you.
 * 6. Paste that URL into CONFIG.MATCH_ENDPOINT in index.html.
 *
 * The sheet will get one row per request: timestamp, target
 * profile ID, requester's MMID, requester's contact, message,
 * and a "Status" column you update manually (Pending/Contacted/
 * Declined) as the admin team follows up.
 * --------------------------------------------------------------
 */

const SHEET_NAME = 'Match Requests';
const ADMIN_EMAIL = 'muslimmarriageQC@gmail.com'; // get an email each time someone submits

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    sheet.appendRow([
      new Date(),
      data.targetProfileId || '',
      data.requesterMmid || '',
      data.requesterContact || '',
      data.message || '',
      'Pending'
    ]);

    if (ADMIN_EMAIL) {
      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: `New match request — ${data.targetProfileId}`,
        body:
          `A brother has requested an introduction.\n\n` +
          `Target profile: ${data.targetProfileId}\n` +
          `Requester MMID: ${data.requesterMmid}\n` +
          `Requester contact: ${data.requesterContact}\n` +
          `Message: ${data.message || '(none)'}\n\n` +
          `Open the "Match Requests" sheet to manage this.`
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Target Profile ID', 'Requester MMID', 'Requester Contact', 'Message', 'Status']);
  }
  return sheet;
}
