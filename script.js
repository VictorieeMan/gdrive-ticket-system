// This script is intended to be used as a Google Apps Script

/* Ticket parameters
for easy customization per event */

const EVENT_URL = 'Your Event URL'; // If you have a website for the event
const EVENT_NAME = 'Your Event Name';
const EVENT_DATE = 'Your Event Date';
const EVENT_TIME = 'Your Event Time';
const EVENT_PLACE = 'Your Event Place';

/* Email generation functions
Adjust these functions if you want to adjust the content of the emails sent.*/
function generateHtmlMailContent(status) {
    /* This function generates the HTML content for the email body based on the 
    status of the form submission. */
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

function generateHtmlTicketContent(status, uniqueHash) {
    /* This function generates the HTML content for the event ticket.*/
    var qrCodeUrl = `https://quickchart.io/qr?text=${encodeURIComponent(uniqueHash)}&size=150`;
    var qrCodeBlob = UrlFetchApp.fetch(qrCodeUrl).getBlob();
    var qrCodeBase64 = Utilities.base64Encode(qrCodeBlob.getBytes());
    var qrCodeImage = `data:image/png;base64,${qrCodeBase64}`;

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
            <p><strong>Unique Code:</strong> ${uniqueHash}</p>
            <p style="text-align: center;"><img src="${qrCodeImage}" alt="QR Code"></p>
            </div>
        </body>
        </html>
    `;
}

function sendReminderEmail(email) {
    /* This function sends a payment reminder email to the user. */
    var subject = 'Reminder: Payment Pending';
    var body = 'Dear User,\n\nThis is a reminder that your payment is still pending. Please complete your payment as soon as possible.\n\nThank you.';
    MailApp.sendEmail(email, subject, body);
}

/* Main program logic
The processResponse() function is the driving force of this program that trigger
most of the other code. Unless programming, leave it be.*/

///Main function to process form responses
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
    var ticketIdIndex = headers.indexOf('ticket_id');

    // Create a set to track existing ticket IDs
    var existingTicketIds = new Set(data.slice(1).map(row => row[ticketIdIndex]));

    // Loop through each row of data (skip the header row)
    for (var i = 1; i < data.length; i++) {
        var row = data[i];

        // Generate and store unique ticket_id if not already present
        if (!row[ticketIdIndex]) {
            var timestamp = row[0]; // Use the timestamp from the first column
            var email = row[1]; // Assuming the email is the second column
            var uniqueHash = generateUniqueHash(timestamp, email);

            // Ensure the ticket_id is unique
            var increment = 0;
            var originalHash = uniqueHash;
            while (existingTicketIds.has(uniqueHash)) {
                uniqueHash = `${originalHash}${increment}`;
                increment++;
            }

            existingTicketIds.add(uniqueHash);
            sheet.getRange(i + 1, ticketIdIndex + 1).setValue(uniqueHash);
        } else {
            var uniqueHash = row[ticketIdIndex];
        }

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
        var htmlTicket = generateHtmlTicketContent(status, uniqueHash);

        // Convert HTML to PDF
        var pdf = convertHtmlToPdf(htmlTicket, `biljett_${uniqueHash}.pdf`);

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

/* Utility functions
These help the program preform its tasks.*/
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
    if (headers.indexOf('ticket_id') === -1) {
        sheet.getRange(1, headers.length + 5).setValue('ticket_id');
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

///Converts HTML to PDF
function convertHtmlToPdf(htmlContent, outputFilename) {
    var blob = Utilities.newBlob(htmlContent, 'text/html', 'status.html');
    var pdf = DriveApp.createFile(blob).getAs('application/pdf').setName(outputFilename);
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

function generateUniqueHash(timestamp, email) {
    var baseString = timestamp + email;
    var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, baseString)
        .map(function(byte) {
            return (byte & 0xFF).toString(16).padStart(2, '0');
        })
        .join('');
    return hashToCustomFormat(hash);
}

function hashToCustomFormat(hash) {
    var letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    var numbers = '0123456789';
    var customHash = '';

    // Generate two leading letters
    for (var i = 0; i < 2; i++) {
        var index = parseInt(hash.substr(i * 2, 2), 16) % letters.length;
        customHash += letters.charAt(index);
    }

    customHash += '-'; // Add a hyphen in between

    // Generate two trailing numbers
    for (var i = 2; i < 4; i++) {
        var index = parseInt(hash.substr(i * 2, 2), 16) % numbers.length;
        customHash += numbers.charAt(index);
    }

    return customHash;
}