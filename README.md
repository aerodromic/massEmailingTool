# Mass Emailing Tool using Google Sheets and Gmail

This tool allows you to send mass emails using data from a Google Sheet and your Gmail account. You can customize the email message by merging it with the data from the spreadsheet.

## Prerequisites

Before using this tool, make sure you have the following:

1. A Google account
2. A Google Sheet with the recipient's data
3. Gmail access enabled for your Google account
4. Google Apps Script enabled for your Google account

## Setup

1. Open the Google Sheet that contains the recipient's data.
2. Create two sheets within the spreadsheet: `Sheet1` and `Sheet2`.
3. In `Sheet1`, populate the following columns:
   - Column 1: Recipient's name
   - Column 2: Company name (optional)
   - Column 3: Recipient's email address
   - (You can add additional columns as needed)
4. In `Sheet2`, enter the following details:
   - Cell A2: Email subject
   - Cell B2: Email message body (can include HTML)
   - Cell C2: BCC email address (optional)
5. Save the spreadsheet.

## Configuring the Script

1. Open the Google Sheet.
2. Click on "Extensions" in the menu bar, then select "Apps Script."
3. In the Apps Script editor, delete any existing code and replace it with the following code:

```javascript
String.prototype.replaceAll = function(search, replacement) {
    var target = this;
    return target.replace(new RegExp(search, 'g'), replacement);
};

function sendEmail() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet1 = ss.getSheetByName('Sheet1');
    var sheet2 = ss.getSheetByName('Sheet2');
    var subject = sheet2.getRange(2, 1).getValue();
    var message = sheet2.getRange(2, 2).getValue();
    var bcc = sheet2.getRange(2, 3).getValue();

    var signature = Gmail.Users.Settings.SendAs.list("me").sendAs.filter(function(account){
        if (account.isDefault) {
            return true;
        }
    })[0].signature;

    message = "<p>" + message + "</p>" + signature;

    var n = sheet1.getLastRow();
    for (var i = 2; i < n + 1; i++) {
        var name = sheet1.getRange(i, 1).getValue();
        var company = sheet1.getRange(i, 2).getValue();
        var emailAddress = sheet1.getRange(i, 3).getValue();

        var newMessage = message.replaceAll("<name>", name).replaceAll("<company>", company);

        GmailApp.sendEmail(emailAddress, subject, newMessage, {
            htmlBody: newMessage,
            bcc: bcc
        });
    }
}
```

4. Save the script by clicking on the floppy disk icon or pressing `Ctrl + S`.
5. Close the Apps Script editor.

## Usage

1. Open the Google Sheet containing the recipient's data.
2. Click on "Add-ons" in the menu bar, then select "Mass Emailing Tool."
3. In the dropdown menu, click on "Send Emails."
4. The script will run and send individual emails to each recipient listed in `Sheet1`, using the subject and message specified in `Sheet2`.
5. The email message will be customized by replacing `<name>` with the recipient's name and `<company>` with the company name (if provided).
6. If you specified a BCC email address in `Sheet2`, it will be included in the BCC field of each email.
7. You will see the progress of the email sending process in the Google Sheets interface. Each row processed will be logged, and any errors encountered will be displayed.

8. Once the script finishes sending all the emails, you can check your Gmail Sent folder to verify that the emails were sent successfully.

Note: It's important to ensure that you have the necessary permissions and quota limits for using GmailApp and SpreadsheetApp services in Google Apps Script. Review and adjust these as needed to avoid any issues.

## Customization

You can modify the code to suit your specific requirements. Here are a few possible customizations:

- Email Formatting: You can customize the email message format by modifying the HTML structure or adding additional placeholders that you can replace using the `replaceAll` function.
- Additional Data Fields: If your spreadsheet contains additional columns, you can retrieve and use that data in the email by modifying the code accordingly.
- Advanced Email Configuration: You can explore the GmailApp documentation to discover additional options for configuring the email, such as attaching files or using different email sending options.

Remember to save the script after making any modifications and test it thoroughly before running it with a large number of recipients.

## Conclusion

This Google Sheets tool provides a convenient way to send mass emails using Gmail and a spreadsheet. By following the setup instructions and customizing the code as needed, you can easily send personalized emails to multiple recipients in a streamlined manner.

Please note that this tool and code sample are provided as-is and Viettonkin Consulting does not take responsibility for any issues or misuse. It's always a good practice to review and test the code thoroughly before using it with important data.

Happy emailing!