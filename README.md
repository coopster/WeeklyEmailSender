# Weekly Parent Emails

A Google Sheets / Google Apps Script tool for creating and sending personalized weekly parent emails from spreadsheet rows.

Each worksheet tab represents one type of parent update. Each row represents one child. The script turns that row into a formatted HTML email using the worksheet's column headings and a configurable letterhead, introduction, and footer.

## Features

- Sends one personalized email per eligible spreadsheet row
- Uses the worksheet tab name as part of the email subject
- Supports a separate email layout/configuration for each worksheet
- Automatically includes non-reserved columns as sections in the email
- Skips blank content cells
- Optional parent/family greeting
- Optional `Send Email` checkbox for selecting which rows should be processed
- `Email Sent` tracking prevents accidental duplicate sends
- Preview dialog shows exactly what will be sent without sending email or updating the sheet
- Automatically validates customization entries against available worksheet names
- Sends a BCC copy to the Google account running the script

## Requirements

This project is intended to be used as a **container-bound Google Apps Script** attached to a Google Sheets spreadsheet.

The spreadsheet should contain:

1. An **Instructions** or configuration sheet containing the named range `Customizations`
2. One or more worksheet tabs containing parent email data
3. The Apps Script code from this repository

## Worksheet Format

Row 1 contains column headings. Data begins in row 2.

Each row represents one child and one email.

### Required Columns

| Column | Purpose |
| --- | --- |
| `Date` | Text displayed as the week/date in the email |
| `Student Name` | Child's name |
| `Send To` | Parent or guardian email address |
| `Email Sent` | Tracks whether the email has already been sent |

Header names must match exactly.

### Optional Reserved Columns

| Column | Purpose |
| --- | --- |
| `Send To Name` | Adds a greeting such as `Dear Smith Family,` |
| `Send Email` | Checkbox controlling whether the row is eligible to be processed |

If `Send Email` is omitted, all otherwise eligible rows are considered.

If it is present, only rows with the checkbox checked are considered.

### Content Columns

Any other labeled column automatically becomes a section in the email.

For example:

| Date | Student Name | Send To | Weekly Glow | Weekly Grow | Teacher Comments | Send Email | Email Sent |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Sep 15 | Emma Collins | parent@example.com | Shared kindly with friends | Working on transitions | Great week! | TRUE | |

The generated email will contain sections titled:

- WEEKLY GLOW
- WEEKLY GROW
- TEACHER COMMENTS

If a child's cell is empty, that section is omitted from that child's email.

A trailing colon in a heading is removed when the heading is displayed in the email.

## Email Customization

Each worksheet can have its own letterhead, introduction, and footer.

The spreadsheet must contain a named range called:

`Customizations`

Its first row must contain these headings:

| Sheet Name | Letterhead Image | Introduction | Footer |
| --- | --- | --- | --- |

### Sheet Name

Select the worksheet tab to which the customization applies.

When the spreadsheet opens, the script updates this column's validation list from the available worksheet names.

The configuration sheet itself is excluded from that list.

### Letterhead Image

The letterhead must be an **image stored in the spreadsheet cell**.

The script reads the image from the cell and embeds it directly into the HTML email.

### Introduction

The introduction appears near the top of the email.

HTML may be used in this field when formatting is desired.

### Footer

The footer appears in the shaded footer area at the bottom of the email.

HTML may also be used in this field.

Each worksheet used for email sending must have exactly one matching row in the `Customizations` range.

## Email Subject

The subject is generated automatically:

```text
<Worksheet Name> for <Student Name> - <Date>
```

For example:

```text
Weekly Preschool Update for Emma Collins - Sep 15
```

## Using the Weekly Email Menu

After the spreadsheet opens, a **Weekly Email** menu is added to Google Sheets.

### Preview

Choose:

**Weekly Email → Preview**

Preview:

- prepares all currently eligible emails
- displays them in a dialog
- shows recipient, BCC, subject, and formatted message
- does **not** send email
- does **not** change `Email Sent`

Preview is the recommended way to verify the messages before sending.

### Send All

Choose:

**Weekly Email → Send All**

The script processes every eligible row.

After an email is successfully sent, the script writes:

```text
yes
```

to that row's `Email Sent` cell.

Rows already containing `yes` in `Email Sent` are skipped, preventing the same row from being sent again accidentally.

The comparison is case-insensitive.

## Which Rows Are Sent?

A row is eligible when:

1. `Email Sent` is not `yes`, and
2. either:
   - there is no `Send Email` column, or
   - the `Send Email` checkbox is checked

A completely empty row is ignored.

If a partially completed row is selected for processing but is missing `Date`, `Student Name`, or `Send To`, the script displays an alert identifying the row and missing field.

## Email Delivery

The value in `Send To` is used as the email's primary recipient.

The Google account running the script is added as a BCC recipient so the sender retains a copy of each message.

Emails are sent using Google Apps Script's `MailApp` service.

## Installation

1. Create or open the Google Sheet that will use the tool.
2. Open **Extensions → Apps Script**.
3. Add the project code from this repository.
4. Save the Apps Script project.
5. Create the `Customizations` named range described above.
6. Add at least one worksheet with the required columns.
7. Reload the spreadsheet.

The **Weekly Email** menu should appear after the spreadsheet opens.

The first time a function requiring Google services is run, Google may ask the user to authorize the script.

## About Dialog

Choose:

**Weekly Email → About...**

to display the application version and a short description.

Current application version:

**1.0.0**

The application version is maintained in the script using:

```javascript
const APP_VERSION = '1.0.0';
```

## Error Handling

The script checks for several common configuration problems, including:

- missing required worksheet columns
- duplicate reserved columns
- missing `Customizations` named range
- missing customization columns
- no customization row for the active worksheet
- multiple customization rows for the same worksheet
- missing or invalid letterhead image
- missing required data in an otherwise active row
- email send failures

Email failures are reported without marking the row as sent.

## Reserved Column Names

The following headings are reserved by the application and are not included as content sections in the email:

```text
Date
Student Name
Send To
Send To Name
Send Email
Email Sent
```

All other non-empty labeled columns are treated as email content.

## Notes for Spreadsheet Designers

- Worksheet names should be meaningful because they are displayed in the email and used in the subject.
- Keep reserved column names unchanged.
- Additional teacher-specific columns can be added freely.
- Columns may appear in any order.
- Empty content cells are simply omitted from that child's email.
- The `Email Sent` column should normally be left for the script to maintain.
- Use `Preview` before `Send All`, especially after changing worksheet columns or email customization.

## Author

Tom Cooper

Initial version: September 15, 2026
