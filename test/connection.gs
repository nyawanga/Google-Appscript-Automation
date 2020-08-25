function InsertData(){
  
  //https://stackoverflow.com/questions/33342777/insert-data-range-from-google-sheets-to-mysql
  
  var conn = new MyConnection() ;                                        // instantiate a connection object
  var stmt = conn.createStatement();
  var data = data_transform() ;
  var start = new Date();
//  conn.setAutoCommit(false);                                           // this should be set with bulk insert if I remember correctly
  
  stmt.executeUpdate('DELETE FROM table_name');
  for( i=0 ; i < data.length ; i++ ){
     var week = Utilities.formatDate(new Date(data[i][0]),"GMT+3", "YYYY-MM-dd")
     var sql = "INSERT INTO table_name (week, col1, col2, col3) VALUES ('" + week + "','" + data[i][1] + "','" + data[i][2] + "','" + data[i][3] + "')";
     var count = stmt.executeUpdate(sql,1)
  }
  
// FOR SOME REASON THIS OPTION WAS NOT WORKING BUT YOU CAN TRY IT OUT
// var start = new Date();
// var stmt = conn.prepareStatement("INSERT INTO table_name (week, col1, col2, col3) VALUES(?, ?, ?, ?)");
//
// for ( var i = 0; i < data.slice(0,50).length; i++ ) {
//   var week = Utilities.formatDate(new Date(data[i][0]),"GMT+3", "YYYY-MM-dd")
//   stmt.setString( 1 , 'week ' + week );
//   stmt.setString( 2 , 'col1 ' + data[i][1] );
//   stmt.setString( 3 , 'col2 ' + data[i][2] );
//   stmt.setString( 4 , 'col3 ' + data[i][3] );
//   stmt.addBatch();
// }
//  var batch = stmt.executeBatch();
//  //Logger.log(sql) ;
  

  stmt.close();
//  conn.commit();
  

// READING FRON A TABLE
  var ss = SpreadsheetApp.getActive();
  var dest_sheet = ss.getSheetByName('DestSheetName');
  var param_1 = "blah" ;
  
  var sql_statement = "SELECT * FROM "tableName" WHERE vertical = '"+param_1+"' AND param_2 NOT IN ('foo','bar') AND week >= CURDATE() - INTERVAL 10 WEEK" ; 
  try{
    var results = connection.createStatement().executeQuery( sql_statement );
    var metaData = results.getMetaData();
    var columns = metaData.getColumnCount();
    
    // Retrieve metaData to a 2D array
    var values = [];
    var value = [];
    var element = '';
 
    // Get table headers
    for(i = 1; i <= columns; i ++){
      element = metaData.getColumnLabel(i);
      value.push(element);
    }
    values.push( value );
    
    // Get table data row by row
    while(results.next()){
      value = [];
      for(i = 1; i <= columns; i ++){
        element = results.getString(i);
        value.push(element);
      }
      values.push(value);
    }
  
    // Close cursor object
    results.close();
    
    // Write data to sheet Data
    dest_sheet.clear();
    SpreadsheetApp.flush();
    dest_sheet.getRange(1, 1, values.length, value.length).setValues(values);
    SpreadsheetApp.getActive().toast('You data has been refreshed.');
  }catch(err){
    SpreadsheetApp.getActive().toast( err.message );
  } 
  
  // ALWAYS CLOSE CONNECTION OBJECT
  conn.close();
  
  var end = new Date();
//  Logger.log('Time elapsed: %sms for %s rows.', end - start, batch.length);
}
