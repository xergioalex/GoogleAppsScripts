
// This function searches for email threads with a specific subject in the user's Gmail inbox,
// retrieves the responses, and saves the sender's name, email, and the response body to a Google Spreadsheet.

function getEmailResponses() {
  var query = 'subject:"Confirmación de Asistencia para Evento..."'; // Change the subject as needed
  var threads = GmailApp.search(query);
  
  // Try to get the existing spreadsheet by name
  var spreadsheetName = "Evento: Email Responses";
  var spreadsheet = null;
  var files = DriveApp.getFilesByName(spreadsheetName);
  
  if (files.hasNext()) {
    spreadsheet = SpreadsheetApp.open(files.next());
  } else {
    spreadsheet = SpreadsheetApp.create(spreadsheetName);
  }
  
  var sheet = spreadsheet.getActiveSheet();
  
  // Clear the sheet and set headers
  sheet.clear();
  sheet.appendRow(["Name", "Email", "Response"]);

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      if (message.isInInbox()) { // Ensure it is a response
        var from = message.getFrom();
        var body = message.getPlainBody();
        
        // Split the 'from' string to separate name and email
        var nameEmail = from.match(/(.*) <(.*)>/);
        var name = nameEmail ? nameEmail[1] : "";
        var email = nameEmail ? nameEmail[2].toLowerCase() : from.toLowerCase();

        sheet.appendRow([name, email, body]);
      }
    }
  }
  Logger.log("Responses have been saved to the spreadsheet.");
}