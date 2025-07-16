# Google Apps Script: Automated Expense Report Management

## Overview
This script automates validation, summary email notifications, and calendar invites for monthly expense reports in Google Sheets. It is designed for finance teams to ensure data completeness, streamline reporting, and improve workflow efficiency.

## Features
- **Row Validation:** Prevents editing new rows until all required fields in previous rows are complete. Highlights incomplete rows and alerts the user.
- **Summary Email:** Sends a monthly summary of expenses for each report sheet to the finance team.
- **Calendar Invite:** Sends a calendar invite (.ics attachment) for the monthly report review on the first working day of each month.
- **Error Handling & Logging:** All main functions include error handling and logging for easier debugging.
- **Easy Setup:** Includes a function to set up all required time-based triggers.

## Deployment Steps
1. **Open Your Google Sheet**
   - Go to the Google Sheet where you want to use this script.

2. **Open the Script Editor**
   - Click on `Extensions` > `Apps Script`.

3. **Copy and Paste the Script**
   - Replace any existing code with the contents of `Solution.gs`.

4. **Save the Script**
   - Click the save icon and name your project (e.g., "Expense Automation").

5. **Set Up Triggers**
   - In the Apps Script editor, select the `setupTriggers` function from the dropdown and click the "Run" ▶️ button.
   - Grant the required permissions when prompted.

6. **Start Using the Sheet**
   - The script will now validate edits and send emails/invites automatically based on the schedule.

## Testing Approach
- **Validation:**
  - Edited rows to ensure incomplete rows are highlighted and alerts are shown.
  - Confirmed that editing is blocked until all required fields are filled.
- **Summary Email:**
  - Ran `sendExpenseSummaryEmail` manually and verified the email content and recipient.
- **Calendar Invite:**
  - Ran `sendExpenseCalendarInviteWithAttachment` manually and checked the .ics attachment and event details.
- **Error Handling:**
  - Introduced errors to confirm that logs and user alerts are generated.
- **Triggers:**
  - Verified that time-based triggers are created in the Apps Script dashboard.

## Customization
- Update `FINANCE_EMAIL` at the top of `Solution.gs` to your finance team's email address.
- Adjust `REQUIRED_FIELDS` if your sheet uses different required columns.


