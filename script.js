// This script is intended to be used as a Google Apps Script
// Script repo url: https://github.com/VictorieeMan/gdrive-ticket-system
// Check the repo for updates and information on how to use the script.

function processResponses() {
    // Run setup function to ensure the environment is set up
    setupSheet();

    // Get the active spreadsheet and the first sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1'); // Change the sheet name if necessary
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    // Get the indices of the required columns
    var paymentStatusIndex = headers.indexOf('PaymentStatus_manual');
    var sentTicketStatusIndex = headers.indexOf('SentTicketStatus_auto');
    var lastReminderIndex = headers.indexOf('LastReminder');
    var reminderCountIndex = headers.indexOf('ReminderCount');

    // Loop through each row of data (skip the header row)
    for (var i = 1; i < data.length; i++) {
        var row = data[i];

        // Skip the row if the SentTicketStatus_auto column is already set to 1
        if (row[sentTicketStatusIndex] == '1') {
            continue;
        }

        // Skip the row if the PaymentStatus_manual column is not set to 1
        if (row[paymentStatusIndex] != '1') {
            // Check if a reminder needs to be sent
            var lastReminder = row[lastReminderIndex];
            if (shouldSendReminder(lastReminder)) {
                sendReminderEmail(row[1]); // Assuming the email is the second column
                sheet.getRange(i + 1, lastReminderIndex + 1).setValue(new Date());

                // Increment the reminder count
                var reminderCount = row[reminderCountIndex] || 0;
                sheet.getRange(i + 1, reminderCountIndex + 1).setValue(reminderCount + 1);
            }
            continue;
        }

        var email = row[1]; // Assuming the email is the second column
        var status = determineStatus(row); // Custom function to determine status

        var subject = 'Your Form Submission Status';
        var htmlContent = generateHtmlMailContent(status);
        var htmlTicket = generateHtmlTicketContent(status);

        // Convert HTML to PDF
        var pdf = convertHtmlToPdf(htmlTicket);

        // Send email with PDF attachment
        MailApp.sendEmail({
            to: email,
            subject: subject,
            htmlBody: htmlContent,
            attachments: [pdf]
        });

        // Update the SentTicketStatus_auto column in the sheet
        sheet.getRange(i + 1, sentTicketStatusIndex + 1).setValue('1');
    }
}

function setupSheet() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1'); // Change the sheet name if necessary
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Check if the required columns are already present
    if (headers.indexOf('PaymentStatus_manual') === -1) {
        sheet.getRange(1, headers.length + 1).setValue('PaymentStatus_manual');
    }
    if (headers.indexOf('SentTicketStatus_auto') === -1) {
        sheet.getRange(1, headers.length + 2).setValue('SentTicketStatus_auto');
    }
    if (headers.indexOf('LastReminder') === -1) {
        sheet.getRange(1, headers.length + 3).setValue('LastReminder');
    }
    if (headers.indexOf('ReminderCount') === -1) {
        sheet.getRange(1, headers.length + 4).setValue('ReminderCount');
    }
}

function determineStatus(row) {
    // Custom logic to determine status based on row data
    // For example, if the first field is "Approved", return "Approved"
    if (row[0] === 'Approved') {
        return 'Approved';
    } else {
        return 'Pending';
    }
}

function generateHtmlMailContent(status) {
    return `
        <html>
        <body>
            <p>Dear User,</p>
            <p>Your form submission status is: <strong>${status}</strong></p>
            <p>Thank you.</p>
        </body>
        </html>
    `;
}

function generateHtmlTicketContent(status) {
    return `
        <html>
        <body>
            <div style="border: 2px solid #000; padding: 20px; width: 600px; margin: 0 auto; font-family: Arial, sans-serif;">
            <h1 style="text-align: center;">Event Ticket</h1>
            <p style="text-align: center;">Thank you for your submission.</p>
            <hr>
            <p><strong>Status:</strong> ${status}</p>
            <p><strong>Date:</strong> ${new Date().toLocaleDateString()}</p>
            <p><strong>Event:</strong> Your Event Name</p>
            </div>
        </body>
        </html>
    `;
}

function convertHtmlToPdf(htmlContent) {
    var blob = Utilities.newBlob(htmlContent, 'text/html', 'status.html');
    var pdf = DriveApp.createFile(blob).getAs('application/pdf').setName('biljett.pdf');
    return pdf;
}

function shouldSendReminder(lastReminder) {
    if (!lastReminder) {
        return true;
    }
    var lastReminderDate = new Date(lastReminder);
    var currentDate = new Date();
    var diffDays = Math.floor((currentDate - lastReminderDate) / (1000 * 60 * 60 * 24));
    return diffDays >= 3;
}

function sendReminderEmail(email) {
    var subject = 'Reminder: Payment Pending';
    var body = 'Dear User,\n\nThis is a reminder that your payment is still pending. Please complete your payment as soon as possible.\n\nThank you.';
    MailApp.sendEmail(email, subject, body);
}

/*
MIT License

Copyright (c) 2024 VictorieeMan

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Repo URL: https://github.com/VictorieeMan/gdrive-ticket-system
*/