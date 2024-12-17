var EMAIL_TEMPLATE_DOC_URL = 'https://docs.google.com/document/d/EXAMPLE_DOC_ID/edit?usp=sharing'; // Replace link with your Google Doc link
var SENDER_NAME = "Organization"; // Add your organization name here

function onFormSubmit(e) {
  var responses = e.namedValues;
  
  // Get email and validate it
  // Get email address from the "Email Address" field in the form
  var email = responses['Email Address'][0].trim(); // Change 'Email Address' to the name of your email field
  if (!validateEmail(email)) {
    Logger.log('Invalid email: ' + email);
    return;
  }

  // Get recipient's name (assuming you have a "Full Name" field in your form)
  var recipientName = responses['Full Name'] ? responses['Full Name'][0].trim() : 'New Member'; // Change 'Full Name' to the name of your name field
  
  // Personalized subject line
  var emailSubject = "Thank you " + recipientName + " for your interest in " + SENDER_NAME + "!"; // Change email subject as needed

  // Get the active sheet before the try block
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  var column = e.values.length + 1;

  try {
    MailApp.sendEmail({
      to: email,
      name: SENDER_NAME, // This helps prevent spam classification
      subject: emailSubject,
      htmlBody: createEmailBody(recipientName),
    });
    Logger.log('email sent to: ' + email);

    // Append the status
    sheet.getRange(row, column).setValue('Email Sent');
  } catch (error) {
    Logger.log('Error sending email: ' + error.toString());
    sheet.getRange(row, column).setValue('Email Failed: ' + error.toString());
  }
}

// Validate email address
function validateEmail(email) {
  var emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return emailRegex.test(email);
}

// Create email body from a Google Doc template
function createEmailBody(recipientName) {
  var docId = DocumentApp.openByUrl(EMAIL_TEMPLATE_DOC_URL).getId();
  var emailBody = docToHtml(docId);
  
  // Add personalization to the email body
  emailBody = emailBody.replace('{{Full Name}}', recipientName || 'New Member');
  
  return emailBody;
}

// Convert Google Doc to HTML
function docToHtml(docId) {
  var url = 'https://docs.google.com/feeds/download/documents/export/Export?id=' +
            docId + '&exportFormat=html';
  var param = {
    method: 'get',
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true,
  };
  return UrlFetchApp.fetch(url, param).getContentText();
}
