/**
 * CONFIGURATION
 */
const FINANCE_EMAIL = 'abc@x.y';
const REQUIRED_FIELDS = ['Date', 'Category', 'Amount', 'Description'];

/**
 * Utility: Get all report sheets matching the naming pattern
 */
function getReportSheets(sheets, pattern) {
  return sheets.filter(sheet => pattern.test(sheet.getName()));
}

/**
 * Utility: Format date for ICS calendar
 */
function formatICSDate(date) {
  return Utilities.formatDate(date, "UTC", "yyyyMMdd'T'HHmmss'Z'");
}

/**
 * Validation on edit
 * Checks all previous rows for required fields before allowing edits.
 * If incomplete, alerts user and highlights the incomplete row.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = e.range.getRow();
    if (row === 1) return; // Skip header row

    // Check all previous rows (from row 2 up to the row before the current one)
    for (let r = 2; r < row; r++) {
      const prevRowData = sheet.getRange(r, 1, 1, headers.length).getValues()[0];
      let missingFields = [];
      REQUIRED_FIELDS.forEach(field => {
        const colIdx = headers.indexOf(field);
        if (colIdx === -1 || !prevRowData[colIdx]) {
          missingFields.push(field);
        }
      });
      if (missingFields.length > 0) {
        // Highlight the incomplete row for user clarity
        sheet.getRange(r, 1, 1, headers.length).setBackground('#ffe599');
        SpreadsheetApp.getUi().alert(
          `Row ${r} is incomplete. Missing required fields: ${missingFields.join(', ')}.\nPlease complete it before editing new rows.`
        );
        // Optionally, clear the edit to prevent further input
        e.range.setValue('');
        return;
      } else {
        // Remove highlight if row is now complete
        sheet.getRange(r, 1, 1, headers.length).setBackground(null);
      }
    }
  } catch (err) {
    Logger.log('onEdit error: ' + err);
    SpreadsheetApp.getUi().alert('An error occurred during validation. Please try again.');
  }
}

/**
 * Send summary email to finance team for each report sheet
 */
function sendExpenseSummaryEmail() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const reportSheetPattern = /^(.+)_((\d{4}_\d{2}))$/;
    const reportSheets = getReportSheets(sheets, reportSheetPattern);

    reportSheets.forEach(sheet => {
      const match = sheet.getName().match(reportSheetPattern);
      if (!match) return; // Defensive: should always match
      const teamName = match[1];
      const month = match[2];
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return; // No summary

      // Find summary rows (look for 'Category Totals' in first cell)
      let summaryStart = -1;
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === 'Category Totals') {
          summaryStart = i;
          break;
        }
      }
      if (summaryStart === -1) return; // No summary found

      // Build summary for email
      let body = `Dear Finance Team,\n\nPlease find below the summary for ${teamName} (${month}):\n\n`;
      let total = 0;
      for (let i = summaryStart + 1; i < data.length; i++) {
        const cat = data[i][0];
        const amt = data[i][1];
        if (cat === 'Total') {
          total = amt;
          break;
        }
        if (cat && amt != null && cat !== '') {
          body += `- ${cat}: ${amt}\n`;
        }
      }
      body += `\nTotal: ${total}\n`;
      const reportUrl = ss.getUrl() + '#gid=' + sheet.getSheetId();
      body += `\nView the full report: ${reportUrl}\n\nBest regards,\nAuto1 Expense Bot`;
      const subject = `Monthly Expense Report: ${teamName} (${month})`;
      Logger.log('About to send email to: ' + FINANCE_EMAIL);
      MailApp.sendEmail(FINANCE_EMAIL, subject, body);
    });
  } catch (err) {
    Logger.log('Expense summary email error: ' + err);
  }
}

/**
 * Send out a calendar invite on the first working day of the month with the report attached to finance team (.ics attachment)
 */
function sendExpenseCalendarInviteWithAttachment() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const reportSheetPattern = /^(.+)_((\d{4}_\d{2}))$/;
    const reportSheets = getReportSheets(sheets, reportSheetPattern);

    reportSheets.forEach(sheet => {
      const match = sheet.getName().match(reportSheetPattern);
      if (!match) return; // Defensive: should always match
      const teamName = match[1];
      const month = match[2];

      // First working day of the month (Mon-Fri)
      const [year, mon] = month.split('_').map(Number);
      let firstDay = new Date(year, mon - 1, 1);
      // Find the first weekday (not Sat/Sun)
      while (firstDay.getDay() === 0 || firstDay.getDay() === 6) {
        firstDay.setDate(firstDay.getDate() + 1);
      }
      const startTime = firstDay;
      const endTime = new Date(firstDay.getTime() + 60 * 60 * 1000); // 1 hour

      // PDF link for the sheet (may need to be tested for your environment)
      const ssUrl = ss.getUrl();
      const pdfUrl = ssUrl.replace(/edit$/, '') +
        `gviz/tq?tqx=out:pdf&sheet=${encodeURIComponent(sheet.getName())}`;
      const reportUrl = ssUrl + '#gid=' + sheet.getSheetId();

      // Description for the calendar invite
      const description = `Please review the monthly expense report.\n\nSheet: ${reportUrl}`;

      // Create .ics content
      const eventTitle = `Expense Report Review: ${teamName} (${month})`;
      const icsContent =
        "BEGIN:VCALENDAR\n" +
        "VERSION:2.0\n" +
        "PRODID:-//Auto1 Expense Bot//EN\n" +
        "BEGIN:VEVENT\n" +
        "UID:" + Utilities.getUuid() + "\n" +
        "DTSTAMP:" + formatICSDate(new Date()) + "\n" +
        "DTSTART:" + formatICSDate(startTime) + "\n" +
        "DTEND:" + formatICSDate(endTime) + "\n" +
        "SUMMARY:" + eventTitle + "\n" +
        "DESCRIPTION:" + description.replace(/\n/g, "\\n") + "\n" +
        "END:VEVENT\n" +
        "END:VCALENDAR";
      const icsBlob = Utilities.newBlob(icsContent, "text/calendar", `${teamName}_${month}_event.ics`);

      // Send email with .ics attachment
      MailApp.sendEmail({
        to: FINANCE_EMAIL,
        subject: eventTitle,
        body: "Please find attached the calendar invite for the monthly expense report review.\n\n" + description,
        attachments: [icsBlob]
      });
    });
  } catch (err) {
    Logger.log('Expense calendar invite error: ' + err);
  }
}

/**
 * Set up triggers (run once manually)
 */
function setupTriggers() {
  try {
    // Monthly report on last day of month at 23:59
    ScriptApp.newTrigger('sendExpenseSummaryEmail')
      .timeBased()
      .onMonthDay(31)
      .atHour(23)
      .nearMinute(59)
      .create();

    ScriptApp.newTrigger('sendExpenseCalendarInviteWithAttachment')
      .timeBased()
      .onMonthDay(1)
      .atHour(9)
      .nearMinute(0)
      .create();

    Logger.log('Triggers set up successfully.');
  } catch (err) {
    Logger.log('Trigger setup error: ' + err);
  }
}
