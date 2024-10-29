function onFormSubmit(e) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Customize these variables according to your form's structure
    var name = data[2];  // Assuming name is in the third column
    var email = data[1];  // Assuming email is in the second column
    var ticketType = data[3];  // Assuming ticket type is in the fourth column
    
    // Create a unique ticket ID
    var ticketId = Utilities.getUuid();

    // Compose the email in HTML
    var htmlBody = '<p>Dear ' + name + ',</p>';
    htmlBody += '<p>Thank you for purchasing a ' + ticketType + ' ticket. Here is your ticket information:</p>';
    htmlBody += '<p><strong>Ticket ID:</strong> ' + ticketId + '</p>';

    // Send the email with the inline image
    MailApp.sendEmail({
        to: email,
        subject: 'Your Ticket',
        htmlBody: htmlBody
    });

    // Store ticket information in the sheet
    // storeTicketInfo(ticketId, email, name, "Valid"); // You can modify the status as needed
}


// function storeTicketInfo(ticketId, email, name, status) {
//     var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RSVP (Responses)2");
//     sheet.appendRow([ticketId, email, name, status, new Date()]);
// }