
function send_email(recepients, email_title, template_name, file_url){
  
  /* MODULE TO SEND AN EMAIL WITH ATTCHMENTS AND AN OPTION LINK TO A FILE ON GOOGLE DRIVE
    recepients - str     : should be a string comma separated
    email_title - str    : title of the email to be sent
    template_name - str  : this is the name of the email template always an HTML file
    file_url - str       : a url for the attachment on GDRIVE
 */
  
  // var attachment_file = DriveApp.getFileById("file_id").getAs(MimeType.PDF);            // attach the file as a PDF
  var htmlData = [email_title, file_url];
  var html = HtmlService.createTemplateFromFile(template_name);
  html.data = htmlData;
  
  var htmlTemplate = html.evaluate().getContent();
  
  var emailBody = {
    to: recepients, //"email@email.com", //
    subject: email_title, 
    //attachments: [attachment_file],                                                       // incase of an attachement
    htmlBody: htmlTemplate 
  }

  MailApp.sendEmail(emailBody);           //SEND THE EMAIL
  //Logger.log("Success !");
  
}


function getPastWeeks(weeks){
  /* DYNAMICALLY GET THE START DATE OF N NUMBER OF WEEKS AGO
  * @param {int} range : takes a 2d array of a single array values
  */

  var today = Utilities.formatDate(new Date(), 'Africa/Nairobi', 'MMMM dd yyyy 09:00:00');      // get the correct date with specific timezone
  var today = new Date(today)      ;                                                            // convert it back to a date object
  var startOfThisWeek = new Date(today - (today.getDay() * 24 * 3600 * 1000));
  var weeksAgo = Utilities.formatDate(new Date(startOfThisWeek - (weeks * 7 * 24 * 3600 * 1000)) ,'Africa/Nairobi', "YYYY-MM-dd");
//  var weeksAgo = new Date(startOfThisWeek - (1 * 7 * 24 * 3600 * 1000));
  
  return weeksAgo;
}


function MyConnection() {
  var ss = SpreadsheetApp.getActive();
  var sheetDetails = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Details');  // Details is just a tab holding the credentials
  var sheetData = ss.getSheetByName('Data');                                           // Data is the tab that the data is being saved into 
 
  this.host = sheetDetails.getRange("B1").getValue();
  this.databaseName = sheetDetails.getRange("B2").getValue();
  this.userName = sheetDetails.getRange("B3").getValue();
  this.password = sheetDetails.getRange("B4").getValue();
  this.port = sheetDetails.getRange("B5").getValue();
  this.tableName = sheetDetails.getRange("B6").getValue();

  var currentMonth = Utilities.formatDate(new Date(), "GMT+03:00", "YYYY-MM-01");
  
  this.url = 'jdbc:mysql://'+ this.host +':3306/'+ this.databaseName;

  try{
    this.connection = Jdbc.getConnection(this.url, this.userName, this.password);
    return this.connection ;
  }
  catch( err ) {
    SpreadsheetApp.getActive().toast(err.message);                                                         // this give a pop up on the main sheet view 
  }
}


function custom_last_data_cell(range, orientation){
  
 /* Gets the last row or column number based on a selected range values and orientation
 * @param {array} range : takes a 2d array of a single array values
 * @param {string} orientation : if passing data from a row the 'row' else if data from a column then 'col'
 * @returns {number} : the last row number with a value
 https://yagisanatode.com/2019/05/11/google-apps-script-get-the-last-row-of-a-data-range-when-other-columns-have-content-like-hidden-formulas-and-check-boxes/
 https://webapps.stackexchange.com/questions/39494/how-to-distinguish-empty-cells-from-zero
 https://docs.google.com/spreadsheets/d/1l6qq5tb03QnWoyNDT5TWx6QvNsmHp1Bu_Cmv0HxnIQo/edit#gid=6
 https://stackoverflow.com/questions/53538956/how-to-get-range-and-then-set-value-in-google-apps-script
 */
  
  var index_count = 0;
  if ( orientation == "row" ){ 
    for(var index = 0; index < range[0].length; index++){
      if((typeof range[0][index] === "string" || typeof range[0][index] === "number" || (range[0][index] instanceof Date)) && range[0][index] !== "")
      {
        index_count = index + 2;
      }
    }
  }
  
  else if ( orientation == "col" ){
   for( var index = 0; index < range.length; index++ ){
    if((typeof range[index][0] === "string" || typeof range[index][0] === "number" || (range[index][0] instanceof Date)) && range[index][0] !== "")
    {
      index_count = index + 2;
    }
   }
  }
  
  return index_count;
}







