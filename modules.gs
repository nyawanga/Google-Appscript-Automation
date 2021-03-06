
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


function typeCheck(data, type){
  if(type == "str"){
    if(typeof data === "string" && data !== ""){
      return true
      }
    }
  else if(type == "int"){
    if(typeof data === "number" && data !== ""){
      return true
      }
    }
  else if(type == "date"){
    if( ( data instanceof Date) && (data !== "") ){
      return true
      }
    }
  else {
    return false
    }
}

 /**
  * Fuction to unpivot and flatten data that has been set up as a pivot
  * @param {array}   data           : range of data to unpivot
  * @param {string}  control_args   : DEFAULT="0;-1" showing the column index with dimension and no cols to skip.should list dimension columns separated by comma and separate skip rows by semi colon also listed comma separated
  *                                   Example is UnPivotData(range, '0,1;2,3') considers columns in index 0 and 1 for dimensions and skips columns 2 and 3
  @return array of transformed / flattened data
  * @customFunction
  **/
function UnPivotData(data, control_args="0;-1"){
//  var ss = SpreadsheetApp.getActive() ;
//  var test = ss.getSheetByName("tests");
//  var data = test.getRange("A1:E6").getValues() ;
//  var row_dims = row_dims=[0,1] //1
//  var skip_cols = [-1]
  
  var row_dims = control_args.split(";")[0].split(",").map(Number)     // found it hard to pass arrays hacked using strings and split
  var skip_cols = control_args.split(";")[1].split(",").map(Number)
  var records = []
  
  for( var idx=0; idx < data.length ; idx++){
    var header_dim = data[0]
    
    var row = data[idx]
    if( idx > 0 ){                                                  // loop through rows
      
      for( var col_idx=0 ; col_idx < row.length ; col_idx++ ){
        var record = []
        
        var cell_value = row[col_idx]
        for( var index in row_dims ){                              //var row_dims_idx = 0 ; row_dims_idx < row_dims.length ; row_dims_idx++ ){
          var row_dim_item = row_dims[index] 
          var row_dim = row[row_dim_item]
          record.push( row_dim )                                   // loop through the row headers adding to array
        }
        if( row_dims.indexOf(col_idx) < 0 && skip_cols.indexOf(col_idx) < 0 ){                                 // loop through columns                                    
          var col_dim = header_dim[col_idx]
          record.push( col_dim )                                   // add column header
          record.push( cell_value )                                // add cell value
            
          Logger.log( record )
          records.push( record )                                   // add to final array
        }
      }
    }
  }
  Logger.log( records )
  
  return records
}


 /*
 *********************************
 The function changes file owners in your drive (the person running the script) and assings them to the specified new owner.
 It also moves them to a fpecified folder for organization purposes
 Incase of failures the file urls are posted in the respective sheets as i.e move_failures and reassign_failures
 Run from the googlesheets for ease of the last step 
 
 * @param {string}  newOwner
 * @param {string}  destinationFolderId
 * @param {string}  
 */

function changeOwnerAndMove( newOwner,destinationFolderId,moveFiles=true ) {

    var ss = SpreadsheetApp.getActive();
    var newOwner = newOwner ;                                                   // Email of the new owner
    var destinationFolderId = destinationFolderId;
    var reassignFailures = [];
    var moveFailures = [];
    var reassignFailureSheet = ss.getSheetByName("reassign_failures");
    var moveFailureSheet = ss.getSheetByName("move_failures");
    
    var files = DriveApp.searchFiles('"me" in owners');                           // Get all the files in Google Drive for the user running this script
    var destFolder = DriveApp.getFolderById(destinationFolderId)                  // Get destination folder

    //// Loop through every file
    while ( files.hasNext() ) {
      var file = files.next();
  //    if(file.getId() == "1s8CwBjNN8XyffGniS9kYl4oat6aVgnAR0C2n" || file.getId() == "1Cv3ErHcyPgABIOyNvCGJ_KhBfGQaZ8woXq8UQ1") {    
        var file_name = file.getName();                                           // Assign the file name to a variable
        var file_id = file.getId()                                                // Get file ID
        
        Logger.log("Changing ownership of " + file_name + " to " + newOwner);    // Log the name of the file changing ownership
        
        try {
            file.addEditor(newOwner);                                           // Set the owner to be the new owner
        }
        catch(err){
            Logger.log(err.message);
        }
        try { 
            file.setOwner( newOwner );
        }
        catch(err){
          reassignFailures.push( [ file.getUrl() ] );                                 // add failed file to reassingFailures
         Logger.log( err.message ); 
        }

        // Move the file to the specified folder
        if( moveFiles == true || moveFiles == "true" ){
            try { 
                destFolder.addFile( file );
                file.getParents().next().removeFile( file );
            }
            catch(err){
                moveFailures.push( [ file.getUrl() ] ) ;                                   // add failed file to moveFailures
                Logger.log( err.message ); 
            }
        }
    //}
    }
    
    // Kepp a record of failures in the respective sheets 
    if( reassignFailureSheet.length > 0 ){
      reassignFailureSheet.getRange( 1,1,reassignFailures.length,1 ).setValues( reassignFailures ) ;  
    }
    else if( moveFailures.length > 0 ){
      moveFailureSheet.getRange( 1,1,moveFailures.length,1 ).setValues( moveFailures ) ;
    }
}

// function moveFiles(file_id, dest_id){
//   //https://tanaikech.github.io/2019/11/20/moving-file-to-specific-folder-using-google-apps-script/

//   var sourceFileId = file_id;
//   var destinationFolderId = dest_id;

//   var file = DriveApp.getFileById(sourceFileId);
//   DriveApp.getFolderById(destinationFolderId).addFile(file);
//   file
//     .getParents()
//     .next()
//     .removeFile(file);
// }



 /*
 ****************************************************************************************
 A function that assigns a datetime everytime a cell is edited 
 
 //https://webapps.stackexchange.com/questions/54527/onedit-only-updates-top-row-after-pasting-editing-multiple-rows
 //https://developers.google.com/apps-script/guides/triggers/events
 */
function onEdit( e ){

  var ss = SpreadsheetApp.getActiveSheet();
  var range = e.range ;
  var change = range.getValue();
  var row = range.getRow() ;
  var dateValue = ss.getRange(row, 3,1,1).getValue() ;
  var timeCol = 2 ;
  var focusTab = "TabName" ;

  if( ss.getName() == focusTab  && typeCheck(dateValue, "date") ){ 
    //&& (typeCheck(change, "int") || typeCheck(change, "str") || typeCheck(change, "date"))){ 
    var t = Utilities.formatDate( new Date(),"GMT+3", "YYYY-MM-dd HH:mm:ss")
    ss.getRange(row, timeCol,1,1).setValue(t)
  }
  
}
