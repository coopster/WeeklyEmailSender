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
 * - Send To Name:  (Optional) Creates a "Dear <name>" in the message body
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
    const emailsPrepared = [];
  let sentCount = 0;
  
  const context = getEmailContext();
  if (context) {

    context.rows
      .map((row, index) => [row, index + 2])
      .filter(([row]) => shouldProcessRow(row, context.columns))
      .forEach(([row, sheetRow]) => {
        const email = prepareEmail(row, sheetRow, context);

        if (email) {
          emailsPrepared.push(email);

          if (processEmail(email, sendFn, context.sheet, sheetRow, context.columns)) {
            sentCount++;
          }
        }
      });
  }

  return {
    emails: emailsPrepared,
    sentCount
  };
}


function getEmailContext() {
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
    return null;
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(value => String(value || "").trim());

  const columns = getReservedColumns(headers);

  const rows = sheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getValues();

  return { sender, sheet, headers, columns, rows };
}


function shouldProcessRow(row, columns) {
  const sendEmail =
    columns.SEND_EMAIL === -1 ||
    row[columns.SEND_EMAIL] === true;

  const emailSent = String(row[columns.EMAIL_SENT] || "")
    .trim()
    .toLowerCase();

  return sendEmail && emailSent !== "yes";
}


function prepareEmail(row, sheetRow, context) {
  const values = getRequiredValues( row, sheetRow, context.columns );

  if (!values) {
    return null;
  }

  const sendToName =
    String(row[context.columns.SEND_TO_NAME] || "").trim();

  const title = context.sheet.getName();

  const subject =
    `${title} for ${values.STUDENT} - ${values.DATE}`;

  const message = buildEmailBody(
    context.headers,
    row,
    values.STUDENT,
    values.DATE,
    sendToName,
    title,
  );

  return {
    sendTo: values.SEND_TO,
    sender: context.sender,
    subject,
    message
  };
}


function getRequiredValues(row, sheetRow, columns) {
  const values = {};

  for (const key of ["SEND_TO", "DATE", "STUDENT"]) {
    values[key] = String(row[columns[key]] || "").trim();

    if (!values[key]) {
      SpreadsheetApp.getUi().alert(
        `Row ${sheetRow}: Missing ${RESERVED[key]}.`
      );

      return null;
    }
  }

  return values;
}


function processEmail( email, sendFn, sheet, sheetRow, columns ) {
  try {
    if (!sendFn(email)) {
      return false;
    }

    sheet
      .getRange(sheetRow, columns.EMAIL_SENT + 1)
      .setValue("yes");

    return true;

  } catch (error) {
    SpreadsheetApp.getUi().alert(
      `Row ${sheetRow}: Email could not be processed. ${error.message}`
    );

    return false;
  }
}


function sendEmailForRow({ sendTo, sender, subject, message }) {
  MailApp.sendEmail(sendTo, subject, "", {
    htmlBody: message,
    bcc: sender
  });

  // Indicate an email was successfully sent. It will throw if failure.
  return true;
}


function showPreviewDialog(previews) {
  const html = previews.length === 0
    ? "<p>No emails would be sent.</p>"
    : previews.map((email, index) => `
        <div style="margin-bottom:24px;">
          <h3>Email ${index + 1}</h3>
          <p>
            <b>To:</b> ${escapeHtml(email.sendTo)}<br>
            <b>BCC:</b> ${escapeHtml(email.sender)}<br>
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
          <b>Note:</b> This is only a preview. No email was sent and no rows were marked as sent.
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

    // Send Email and Send To Name are optional.
    if ((key === "SEND_EMAIL" || key === "SEND_TO_NAME") && matches.length === 0) {
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

function shouldProcessColumn(header, value) {
  if (!header) {
    return false;
  }

  if (RESERVED_HEADERS.has(header)) {
    return false;
  }

  if (!String(value ?? "").trim()) {
    return false;
  }

  return true;
}


function formatHeader(header) {
  return String(header)
    .replace(/:\s*$/, "")
    .toUpperCase();
}


function buildEmailBody(
  headers,
  row,
  student,
  date,
  sendToName,
  title,
) {
  const introduction = getNamedValue("EmailIntroduction");
  const footer = getNamedValue("EmailFooter");

  const contentBlocks = headers
    .map((header, index) => [header, row[index]])
    .filter(([header, value]) =>
      shouldProcessColumn(header, value)
    )
    .map(([header, value], index) => {
      const headingClass = index % 2 === 0 ? "heading gold" : "heading blue";

      return `
        <tr>
          <td class="section">
            <div class="${headingClass}">
              ${escapeHtml(formatHeader(header))}
            </div>

            <table class="content-box" role="presentation">
              <tr>
                <td>
                  ${escapeHtml(String(value).trim())}
                </td>
              </tr>
            </table>
          </td>
        </tr>
      `;
    })
    .join("");

  const greeting = sendToName
    ? `
      <p>
        Dear ${escapeHtml(sendToName)},
      </p>
    `
    : "";


  return `
    <style>
      .email {
        width: 100%;
        max-width: 700px;
        margin: 0 auto;
        border-collapse: collapse;
        font-family: Arial, Helvetica, sans-serif;
        color: #444444;
        background-color: #ffffff;
      }

      .email td {
        font-size: 15px;
        line-height: 1.5;
      }

      .header {
        padding: 12px 30px 18px 30px;
      }

      .title {
        padding: 10px 24px;
        text-align: center;
        font-size: 28px;
        font-weight: bold;
        color: #ffffff;
        background-color: #275985;
      }

      .intro {
        padding: 6px 30px 18px 30px;
      }

      .intro p {
        margin: 0 0 18px 0;
      }

      .student-info {
        width: 100%;
        border-collapse: collapse;
      }

      .student-info td {
        width: 50%;
        padding: 0;
      }

      .section {
        padding: 12px 30px 10px 30px;
      }

      .heading {
        width: fit-content;
        min-width: 160px;
        margin: 0 auto;
        padding: 7px 20px;
        color: #ffffff;
        font-weight: bold;
        text-align: center;
      }

      .gold {
        background-color: #e7ad20;
      }

      .blue {
        background-color: #275985;
      }

      .content-box {
        width: 100%;
        border-collapse: collapse;
        border: 1px solid #c8d4de;
      }

      .content-box td {
        padding: 20px;
      }

      .footer {
        padding: 12px 30px 20px 30px;
      }

      .footer-box {
        width: 100%;
        border-collapse: collapse;
        background-color: #eef4f7;
      }

      .footer-box td {
        padding: 22px 20px;
        text-align: center;
        font-size: 14px;
      }

      .email-footer {
        padding: 18px;
        text-align: center;
        line-height: 1.5;
      }
    </style>

    <table class="email" role="presentation">

      <tr>
        <td class="header">
          <div class="title">
            ${escapeHtml(title)}
          </div>
        </td>
      </tr>

      <tr>
        <td class="intro">
          ${greeting}

          <p>${introduction}</p>

          <table class="student-info" role="presentation">
            <tr>
              <td>
                <b>Child:</b> ${escapeHtml(student)}
              </td>

              <td align="right">
                <b>Week of:</b> ${escapeHtml(date)}
              </td>
            </tr>
          </table>
        </td>
      </tr>

      ${contentBlocks}

      <tr>
        <td class="footer">
          <table class="footer-box" role="presentation">
            <tr>
              <td class="email-footer">
                ${footer}
              </td>
            </tr>
          </table>
        </td>
      </tr>

    </table>
  `;
}


function getNamedValue(name) {
  const range = SpreadsheetApp
    .getActiveSpreadsheet()
    .getRangeByName(name);

  if (!range) {
    throw new Error(`The range "${name}" was not found. Check your instructions sheet.`);
  }

  return String(range.getValue() || "");
}


function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}