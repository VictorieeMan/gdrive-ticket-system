// This script is intended to be used as a Google Apps Script

function processResponses() {
    // Get the active spreadsheet and the first sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1'); // Change the sheet name if necessary
    var data = sheet.getDataRange().getValues();

    // Loop through each row of data (skip the header row)
    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var email = row[1]; // Assuming the email is the second column
        var status = determineStatus(row); // Custom function to determine status

        var subject = 'Your Form Submission Status';
        var htmlContent = generateHtmlMailContent(status);
        var htmlTicket = generateHtmlTicketContent();

        // Convert HTML to PDF
        var pdf = convertHtmlToPdf(htmlTicket);

        // Send email with PDF attachment
        MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: htmlContent,
        attachments: [pdf]
        });
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
            <p><strong>Location:</strong> Event Location</p>
            <hr>
            <p style="text-align: center;">Please bring this ticket to the event.</p>
            <p style="text-align: center;">Best regards,<br>Your Team</p>
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