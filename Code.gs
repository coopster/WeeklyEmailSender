function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Weekly Email')
      .addItem('Send All', 'sendWeeklyEmail')
      .addItem('Preview', 'previewWeeklyEmail')
      .addToUi();
}

/*
 * Weekly Email Sender
 *
 * Sends one HTML email for each eligible row in the active sheet.
 *
 * Sheet format:
 * - Row 1 contains column labels.
 * - Data begins on row 2.
 *
 * Reserved columns:
 * - Date:        Literal date text used in the email body.
 * - Send To:      Recipient address placed in BCC.
 * - Send Email:   (Optional) Checkbox; row is processed only when checked.
 * - Email Sent:  If "yes", the row is skipped. After a successful send,
 *                this value is set to "yes".
 *
 * All other labeled columns are included automatically in the email body.
 * The active sheet (tab) name is used as the email subject.
 *
 * "Send All" sends the email and updates the spreadsheet marking Email Sent.
 * "Preview" action preview emails in a Sheets dialog without
 * sending them or marking them as sent.
 */

const RESERVED = {
  DATE: "Date:",
  STUDENT: "Student Name:",
  SEND_TO: "Send To:",
  SEND_TO_NAME: "Send To Name:",
  SEND_EMAIL: "Send Email:",
  EMAIL_SENT: "Email Sent:"
};

const RESERVED_HEADERS = new Set(Object.values(RESERVED));

function previewWeeklyEmail() {
  const result = processWeeklyEmail(() => false);
  showPreviewDialog(result.emails);
}

function sendWeeklyEmail() {
  const result = processWeeklyEmail(sendEmailForRow);

  SpreadsheetApp.getUi().alert(
    `${result.sentCount} email${result.sentCount === 1 ? "" : "s"} sent.`
  );
}

function processWeeklyEmail(sendFn) {
  const senderName = getUserDisplayName();
  const sender = Session.getEffectiveUser().getEmail();
  
  if (!sender) {
    throw new Error(
      "Could not determine the Google account running this script."
    );
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getActiveSheet();

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No emails to process.");
    return { emails: [], sentCount: 0 };
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(value => String(value || "").trim());

  const columns = getReservedColumns(headers);
  const rows = sheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getValues();

  const emailsPrepared = [];
  let sentCount = 0;

  rows.forEach((row, index) => {
    const sheetRow = index + 2;

    const sendEmail =
      columns.SEND_EMAIL === -1 ||
      row[columns.SEND_EMAIL] === true;

    const emailSent = String(row[columns.EMAIL_SENT] || "")
      .trim()
      .toLowerCase();

    if (!sendEmail || emailSent === "yes") {
      return;
    }

    /* Collect up required field values */
    const values = {};
    for (const key of ["SEND_TO", "DATE", "STUDENT"]) {
      values[key] = String(row[columns[key]] || "").trim();

      if (!values[key]) {
        SpreadsheetApp.getUi().alert(
          `Row ${sheetRow}: Missing ${RESERVED[key]}.`
        );
        return;
      }
    }

    const {
      SEND_TO: sendTo,
      DATE: date,
      STUDENT: student
    } = values;
    
    const subject =
      `${sheet.getName()} for ${student} - ${date}`;

    const message =
      buildEmailBody(headers, row, student, senderName);

    const email = {
      sender,
      sendTo,
      subject,
      message
    };

    emailsPrepared.push(email);

    try {
      if (sendFn(email)) {
        sheet
          .getRange(sheetRow, columns.EMAIL_SENT + 1)
          .setValue("yes");

        sentCount++;
      }
    } catch (error) {
      SpreadsheetApp.getUi().alert(`Row ${sheetRow}: Email could not be processed. ${error.message}`);
    }
  });

  return {
    emails: emailsPrepared,
    sentCount
  };
}

function getUserDisplayName() {
  const person = People.People.get("people/me", {
    personFields: "names"
  });

  if (
    !person.names ||
    person.names.length === 0 ||
    !person.names[0].displayName
  ) {
    throw new Error(
      "Could not determine the display name for the current Google account."
    );
  }

  return person.names[0].displayName;
}

function sendEmailForRow({
  sender,
  sendTo,
  subject,
  message
}) {
  MailApp.sendEmail(sender, subject, "", {
    htmlBody: message,
    bcc: sendTo
  });

  /* Indicate an email was successfully sent. Throw if failure. */
  return true;
}

function showPreviewDialog(previews) {
  const html = previews.length === 0
    ? "<p>No emails would be sent.</p>"
    : previews.map((email, index) => `
        <div style="margin-bottom:24px;">
          <h3>Email ${index + 1}</h3>
          <p>
            <b>To:</b> ${escapeHtml(email.sender)}<br>
            <b>BCC:</b> ${escapeHtml(email.sendTo)}<br>
            <b>Subject:</b> ${escapeHtml(email.subject)}
          </p>
          <hr>
          ${email.message}
        </div>
      `).join("");

  const output = HtmlService
    .createHtmlOutput(`
      <div style="font-family:Arial,sans-serif;padding:16px;">
        <p>
          <b>Note:</b> No email was sent and no rows were marked as sent.
        </p>
        ${html}
      </div>
    `)
    .setWidth(700)
    .setHeight(600);

  SpreadsheetApp
    .getUi()
    .showModalDialog(output, "Email Preview");
}

function getReservedColumns(headers) {
  const columns = {};

  for (const [key, label] of Object.entries(RESERVED)) {
    const matches = headers
      .map((header, index) => header === label ? index : -1)
      .filter(index => index !== -1);

    // Send Email is optional.
    if (key === "SEND_EMAIL" && matches.length === 0) {
      columns[key] = -1;
      continue;
    }

    if (matches.length === 0) {
      throw new Error(
        `Required column "${label}" is missing from row 1.`
      );
    }

    if (matches.length > 1) {
      throw new Error(
        `Column "${label}" appears more than once in row 1.`
      );
    }

    columns[key] = matches[0];
  }

  return columns;
}

function buildEmailBody(headers, row, student, senderName) {
  let message =
    "<p>Hello,<br>" +
    "Here is the weekly update for " +
    escapeHtml(student) +
    ":</p>";

  headers.forEach((header, index) => {
    if (!header || RESERVED_HEADERS.has(header)) {
      return;
    }

    message +=
      "<p><b>" +
      escapeHtml(header) +
      "</b> " +
      escapeHtml(row[index]) +
      "</p>";
  });

  message +=
    "<p>Thank you,<br>" + 
    escapeHtml(senderName) +
    "</p>";

  return message;
}


function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}