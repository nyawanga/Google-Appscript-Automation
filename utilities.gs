/**
 * Lists the labels in the user's account.
 */
function listLabels() {
  var response = Gmail.Users.Labels.list('me');
  if (response.labels.length == 0) {
    Logger.log('No labels found.');
  } else {
    Logger.log('Labels:');
    for (var i = 0; i < response.labels.length; i++) {
      var label = response.labels[i];
      Logger.log('- %s', label.name);
    }
  }
}


function importCSVFromGmail() {
  var threads = GmailApp.search("STUDENT");
  var message = threads[0].getMessages()[0];
  var attachment = message.getAttachments()[0];
  if (attachment.getContentType() === "text/csv") {
    var sheet = SpreadsheetApp.getActiveSheet();
    var csvData = Utilities.parseCsv(attachment.getDataAsString(), ",");
    // Remember to clear the content of the sheet before importing new data
    sheet.clearContents().clearFormats();
    sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData); 
  }
}

function getCsvFromGmail() {
  // Get the newest Gmail thread based on sender and subject
  var gmailThread = GmailApp.search("from:test@example.gmail subject:\'Check out the \"NPS_Score_Verticals\" report\'", 0, 1)[0];
  
  // Get the attachments of the latest mail in the thread.
  var attachments = gmailThread.getMessages()[gmailThread.getMessageCount() - 1].getAttachments();
  
  if (attachments) {
    // Get and and parse the CSV from the first attachment
    var csv = Utilities.parseCsv(attachments[0].getDataAsString());
    Logger.log(csv);
    return csv;
   }
}

function post_data(){
  var dest_sheet = SpreadsheetApp.getActive().getSheetByName("dest");
  var csv = getCsvFromGmail() ;
  dest_sheet.getRange(1,1, csv.length, csv[0].length).clear().setValues(csv) ;
  
}
