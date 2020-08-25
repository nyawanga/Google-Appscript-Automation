function InsertData(){
  
  //https://stackoverflow.com/questions/33342777/insert-data-range-from-google-sheets-to-mysql
  
  var conn = new MyConnection() ;                                        // instantiate a connection object
  var stmt = conn.createStatement();
  var data = data_transform() ;
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
  conn.close();
  
  var end = new Date();
//  Logger.log('Time elapsed: %sms for %s rows.', end - start, batch.length);
}
