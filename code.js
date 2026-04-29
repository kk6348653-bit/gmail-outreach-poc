const SPREADSHEET_ID = 'YOUR_GOOGLE_SHEET_ID'; // 记得替换为实际的 Google Sheet ID
const SHEET_NAME = 'contacts';

const DAILY_INITIAL_LIMIT = 3;  // 每日限制的首封邮件数
const DAILY_FOLLOWUP_LIMIT = 3;  // 每日限制的 follow-up 邮件数
const FOLLOWUP_AFTER_MINUTES = 3; // Follow-up 时间间隔

function getContacts() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet with name "${SHEET_NAME}" not found.`);
  }

  return sheet.getDataRange().getValues();
}

function sendHelloWorldEmails() {
  const contacts = getContacts();
  let sentCount = 0;
  
  for (let i = 1; i < contacts.length; i++) {
    const email = contacts[i][0];
    const name = contacts[i][1];
    const university = contacts[i][2];
    const role = contacts[i][3];
    const status = contacts[i][4];
    const sentAt = new Date(contacts[i][5]);
    const notes = contacts[i][6];

    if (status !== 'sent') {
      try {
        const subject = `Quick hello from EnrollX`;
        const body = `Hi ${name},\n\nI'm doing a quick test of our Gmail automation workflow. This message was generated from Apps Script.`;
        GmailApp.sendEmail(email, subject, body);

        sheet.getRange(i + 1, statusCol + 1).setValue('sent');
        sheet.getRange(i + 1, sentAtCol + 1).setValue(new Date());
        sheet.getRange(i + 1, notesCol + 1).setValue('Sent by Apps Script hello-world POC');
        
        sentCount++;
      } catch (error) {
        sheet.getRange(i + 1, statusCol + 1).setValue('error');
        sheet.getRange(i + 1, notesCol + 1).setValue(error.message);
      }
    }
  }

  Logger.log(`Done. Sent ${sentCount} email(s).`);
}

function sendFollowupOne() {
  const contacts = getContacts();
  let sentCount = 0;

  for (let i = 1; i < contacts.length; i++) {
    const email = contacts[i][0];
    const status = contacts[i][4];
    const followUpStatus = contacts[i][8]; // Follow-up status
    const lastReplyAt = new Date(contacts[i][10]); // Last reply date

    if (status === 'sent' && followUpStatus !== 'sent' && (new Date() - lastReplyAt) > (FOLLOWUP_AFTER_MINUTES * 60 * 1000)) {
      try {
        const subject = `Follow-up 1 sent by Apps Script`;
        const body = `Hi, just checking back to see if you received my previous email.`;
        GmailApp.sendEmail(email, subject, body);

        sheet.getRange(i + 1, followUpStatusCol + 1).setValue('sent');
        sheet.getRange(i + 1, followUpSentAtCol + 1).setValue(new Date());

        sentCount++;
      } catch (error) {
        sheet.getRange(i + 1, followUpStatusCol + 1).setValue('error');
        sheet.getRange(i + 1, notesCol + 1).setValue(error.message);
      }
    }
  }

  Logger.log(`Done. Sent ${sentCount} follow-up email(s).`);
}