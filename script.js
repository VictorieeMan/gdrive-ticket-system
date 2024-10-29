// This script is intended to be used as a Google Apps Script

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

    // Loop through each row of data (skip the header row)
    for (var i = 1; i < data.length; i++) {
        var row = data[i];

        // Skip the row if the SentTicketStatus_auto column is already set to 1
        if (row[sentTicketStatusIndex] == '1') {
            continue;
        }

        // Skip the row if the PaymentStatus_manual column is not set to 1
        if (row[paymentStatusIndex] != '1') {
            continue;
        }

        var row = data[i];
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
        var sentTicketStatusColumn = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('SentTicketStatus_auto') + 1;
        sheet.getRange(i + 1, sentTicketStatusColumn).setValue('1');
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