//*************** MySQL to Google Spreadsheet ***************//

function myMySQLFetchData() { 
  //Connect to database
  //Change the following as per your database credentials
  var conn = Jdbc.getConnection("<database>", "<username>", "<password>"); 
  //Create statement variable
  var stmt = conn.createStatement();
  // Get script starting time
  var start = new Date(); 
  
  //Obtain the currently active spreadsheet
  var doc = SpreadsheetApp.getActiveSpreadsheet(); 
  //Obtain the currently active sheet
  var sheet = doc.getSheetByName('AMDashboard');
  //get start date
  var startdate = sheet.getRange("C1:C1").getDisplayValue() 
  //get end date
  var enddate = sheet.getRange("C2:C2").getDisplayValue() 
  
  //log start date
  Logger.log(startdate) 
  //log end date
  Logger.log(enddate) 
  
  /*try out the query*/
  /*var query = "SELECT AMName AS 'Team', COUNT(listName) AS 'Production', SUM(listStatusID > 0) AS 'Sample Audited', SUM(listStatusID > 0) / COUNT(listName) * 100 AS '% Sample Audited', SUM(listStatusID = 1) AS 'Good Outcome', SUM(listStatusID = 2) AS 'Bad Outcome', SUM(listStatusID = 3) AS 'Needs Improvement' FROM consolidatedDataAll WHERE addedOn >= '"+startdate+"' AND addedOn <= '"+enddate+"' GROUP BY Team;"*/
  /*Logger.log(query)*/
  
  //Relevant query for Resultset object
  var rs = stmt.executeQuery("SELECT AMName AS 'Team', COUNT(listName) AS 'Production', SUM(listStatusID > 0) AS 'Sample Audited', SUM(listStatusID > 0) / COUNT(listName) * 100 AS '% Sample Audited', SUM(listStatusID = 1) AS 'Good Outcome', SUM(listStatusID = 2) AS 'Bad Outcome', SUM(listStatusID = 3) AS 'Needs Improvement', SUM(listStatusID = 2) + SUM(listStatusID = 3) AS 'Total Errors' FROM consolidatedDataAll WHERE addedOn >= '"+startdate+"' AND addedOn <= '"+enddate+"' GROUP BY Team;"); 
  /*May need to correct the query with COUNT DISTINCT() and WHERE clause on isReverted, isNewOrAudited*/  
  
  //Obtain the currently active spreadsheet
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  //Obtain the cell a1
  var cell = doc.getRange('a1');
  //Actual display offset starts at row 4
  var row = 4;
  //Mysql table column name count.
  var getCount = rs.getMetaData().getColumnCount(); 
  
  //Mysql table's column name will be fetched and added in spreadsheet.
  for (var i = 0; i < getCount; i++){  
     cell.offset(row, i).setValue(rs.getMetaData().getColumnName(i+1)); 
  }  
  
  //Mysql table's data will be fetched and added in spreadsheet columnwise.
  var row = 5; 
  while (rs.next()) {
    for (var col = 0; col < rs.getMetaData().getColumnCount(); col++) { 
      cell.offset(row, col).setValue(rs.getString(col + 1)); 
    }
    row++;
  }
  
  //Close everything
  rs.close();
  stmt.close();
  conn.close();
  
  //Get script ending time
  var end = new Date(); 
  //To generate script log.
  Logger.log('Time elapsed: ' + (end.getTime() - start.getTime()));  
  //To view the log click on View -> Logs.
}