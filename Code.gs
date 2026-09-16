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
 * - Date         Literal date text used in the email body.
 * - Send To      Recipient address placed in BCC.
 * - Send To Name (Optional) Creates a "Dear <name>" in the message body
 * - Send Email   (Optional) Checkbox; row is processed only when checked.
 * - Email Sent   If "yes", the row is skipped. After a successful send,
 *                this value is set to "yes".
 *
 * All other labeled columns are included automatically in the email body.
 * The active sheet (tab) name is used as the email subject.
 *
 * "Send All" sends the email and updates the spreadsheet marking Email Sent.
 * "Preview" action preview emails in a Sheets dialog without
 * sending them or marking them as sent.
 * 
 * Author: Tom Cooper, 9/15/26
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Weekly Email')
    .addItem('Send All', 'sendWeeklyEmail')
    .addItem('Preview', 'previewWeeklyEmail')
    .addSeparator()
    .addItem('About...', 'aboutInfo')
    .addToUi();

  try {
    updateCustomizationValidation();
  } catch (error) {
    console.error(
      `Customization validation failed: ${error.message}`
    );
  }
}

const APP_VERSION = '1.0.0';

function aboutInfo() {
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    'Weekly Parent Emails',
    `Version ${APP_VERSION}

Creates personalized parent emails from spreadsheet rows.

Each worksheet can have its own letterhead, introduction, and footer. Use Preview to review messages before sending.

© Whitinsville Christian School`,
    ui.ButtonSet.OK
  );
}

const RESERVED = {
  DATE: "Date",
  STUDENT: "Student Name",
  SEND_TO: "Send To",
  SEND_TO_NAME: "Send To Name",
  SEND_EMAIL: "Send Email",
  EMAIL_SENT: "Email Sent"
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

  const customization = getEmailCustomization(sheet.getName());

  return { sender, sheet, headers, columns, rows, customization };
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
    context.customization.introduction,
    context.customization.footer
  );

  return {
    sendTo: values.SEND_TO,
    sender: context.sender,
    subject,
    message,
    inlineImages: {
      letterhead: context.customization.letterhead
    }
  };
}

function getRequiredValues(row, sheetRow, columns) {
  // Ignore completely empty rows.
  if (row.every(cell => String(cell || "").trim() === "")) {
    return null;
  }

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


function sendEmailForRow({ sendTo, sender, subject, message, inlineImages }) {
  MailApp.sendEmail(sendTo, subject, "", {
    htmlBody: message,
    bcc: sender,
    inlineImages: inlineImages
  });

  // Indicate an email was successfully sent. It will throw if failure.
  return true;
}


function showPreviewDialog(previews) {
  if (previews.length === 0) {
    return "<p>No emails would be sent.</p>"
  }
  const html = previews.map((email, index) => {
    const previewMessage = email.message.replace(
      "cid:letterhead",
      blobToDataUri(email.inlineImages.letterhead)
    );

    return `
      <div style="margin-bottom:24px;">
        <h3>Email ${index + 1}</h3>
        <p>
          <b>To:</b> ${escapeHtml(email.sendTo)}<br>
          <b>BCC:</b> ${escapeHtml(email.sender)}<br>
          <b>Subject:</b> ${escapeHtml(email.subject)}
        </p>
        <hr>
        ${previewMessage}
      </div>
    `;
  }).join("");

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


function blobToDataUri(blob) {
  const base64 = Utilities.base64Encode(blob.getBytes());

  return `data:${blob.getContentType()};base64,${base64}`;
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
  introduction,
  footer
) {

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

      .letterhead {
        padding: 12px 30px 6px 30px;
        text-align: center;
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
        <td class="letterhead">
          <img
            src="cid:letterhead"
            alt="Whitinsville Christian School"
            style="
              display: block;
              width: 280px;
              max-width: 80%;
              height: auto;
              margin: 0 auto;
            "
          >
        </td>
      </tr>

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


function getImageFromCell(cell) {
  const image = cell.getValue();

  if (
    !image ||
    image.valueType !== SpreadsheetApp.ValueType.IMAGE
  ) {
    throw new Error(
      `Expected an image in ${cell.getSheet().getName()}!${cell.getA1Notation()}.`
    );
  }

  const url = image.getContentUrl();

  if (!url) {
    throw new Error(
      `Could not read the image in ${cell.getA1Notation()}.`
    );
  }

  return UrlFetchApp.fetch(url).getBlob();
}


function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/*
 * Customizations
 */

const CONFIG = {
  SHEET_NAME: "Sheet Name",
  LETTERHEAD: "Letterhead Image",
  INTRODUCTION: "Introduction",
  FOOTER: "Footer"
};

const CUSTOMIZATIONS_RANGE = "Customizations";


function updateCustomizationValidation() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const range = spreadsheet.getRangeByName(CUSTOMIZATIONS_RANGE);

  if (!range) {
    throw new Error(
      `Named range "${CUSTOMIZATIONS_RANGE}" was not found.`
    );
  }

  const headers = range
    .getValues()[0]
    .map(value => String(value || "").trim());

  const columns = getConfigColumns(headers);

  const sheetNames = spreadsheet
    .getSheets()
    .map(sheet => sheet.getName())
    .filter(name => name !== range.getSheet().getName());

  const validation = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(sheetNames, true)
    .setAllowInvalid(false)
    .setHelpText(
      "Select the sheet this customization applies to."
    )
    .build();

  const firstDataRow = range.getRow() + 1;
  const rowCount = range.getNumRows() - 1;
  const sheetNameColumn =
    range.getColumn() + columns.SHEET_NAME;

  range
    .getSheet()
    .getRange(firstDataRow, sheetNameColumn, rowCount, 1)
    .setDataValidation(validation);
}


function getEmailCustomization(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const range = spreadsheet.getRangeByName(CUSTOMIZATIONS_RANGE);

  if (!range) {
    throw new Error(
      `Named range "${CUSTOMIZATIONS_RANGE}" was not found.`
    );
  }

  const values = range.getValues();
  const headers = values[0]
    .map(value => String(value || "").trim());

  const columns = getConfigColumns(headers);

  const matches = [];

  for (let i = 1; i < values.length; i++) {
    const configuredSheetName = String(
      values[i][columns.SHEET_NAME] || ""
    ).trim();

    if (configuredSheetName === sheetName) {
      matches.push(i);
    }
  }

  if (matches.length === 0) {
    throw new Error(
      `No customization is configured for sheet "${sheetName}".`
    );
  }

  if (matches.length > 1) {
    throw new Error(
      `Sheet "${sheetName}" appears more than once in "${CUSTOMIZATIONS_RANGE}".`
    );
  }

  const rowIndex = matches[0];
  const sheet = range.getSheet();
  const absoluteRow = range.getRow() + rowIndex;

  const letterheadCell = sheet.getRange(
    absoluteRow,
    range.getColumn() + columns.LETTERHEAD
  );

  return {
    letterhead: getImageFromCell(letterheadCell),
    introduction: String(
      values[rowIndex][columns.INTRODUCTION] || ""
    ),
    footer: String(
      values[rowIndex][columns.FOOTER] || ""
    )
  };
}


function getConfigColumns(headers) {
  const columns = {};

  for (const [key, label] of Object.entries(CONFIG)) {
    const index = headers.indexOf(label);

    if (index === -1) {
      throw new Error(
        `Configuration column "${label}" is missing from "${CUSTOMIZATIONS_RANGE}".`
      );
    }

    columns[key] = index;
  }

  return columns;
}