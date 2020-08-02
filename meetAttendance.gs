/* Mark attendance against a nominal role v1.0- Subhash Subramanian */

// Menu Options 
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Attendance')
  .addItem('Mark Attendance', 'markAllCols')
  .addToUi();
}



function markAllCols(){
  //get last row of Roll column
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //get lastrow of Roll column
  var lastrowR = lastRowofColumn(ss,1);
  var lastCol = ss.getDataRange().getLastColumn();
  var thisCol = lastCol;
  var firstCol = 2;
  var firstRow = 3;
  var curvalueColA = ss.getRange(1,thisCol).getValue().toString();
  while (curvalueColA.substring(curvalueColA.length-5) !== "_Done" && thisCol > firstCol){
    markthisCol(ss,thisCol,lastrowR);
    thisCol = thisCol-1;
    curvalueColA = ss.getRange(1,thisCol).getValue().toString();
    lastCol = ss.getDataRange().getLastColumn();
  }
  
  ss.getRange(firstRow, firstCol+1, lastrowR-firstRow+1, lastCol-firstCol).activate();
  ss.getActiveRangeList().setHorizontalAlignment('center');
    
  
  Browser.msgBox("Attendance Marked");  
  // Second check if (ss.getRange[3][thisCol]) =="P" or ss.getRange[3][thisCol]) =="A" ) :
  //tell marking attendence is done and go to next col.
}


function markthisCol(ss,thisCol,lastrowR){
  // get lastrow of attendees column
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //get lastrow of attendees column
  var lastCol = ss.getDataRange().getLastColumn();
  var colR = 1 //always in first column
  var lastrowR = lastRowofColumn(ss,colR);
  var colA = thisCol;
  insertCol(colA);
  var lastrowA = lastRowofColumn(ss,thisCol);
  
  // edit first two rows
  // indicate that attendance is marked for the column 
  var curvalueColAR1 = ss.getRange(1,colA).getValue();
  var curvalueColAR2 = ss.getRange(2,colA).getValue();
  //ss.getRange(1,colA).setValue(curvalueColA + "_Done");
  ss.getRange(1,colA+1).setValue(curvalueColAR1 + "_Done");
  ss.getRange(2,colA+1).setValue(curvalueColAR2);  
  
  for (rowR=3; rowR < lastrowR+1; rowR++){ //rowR loop
    lastCol = ss.getDataRange().getLastColumn();
    //Browser.msgBox("I am on Roll row " + rowR + " and Att row " + rowA)
    for (rowA=3; rowA < lastrowA+1; rowA++){ //rowA loop
      if (ss.getRange(rowA, colA).getValue() == ss.getRange(rowR, colR).getValue()){
         ss.getRange(rowR, colA+1).setValue("P");   // Present
         break;
      }else if (rowA == lastrowA) {
          ss.getRange(rowR, colA+1).setValue("A");  // Absent
        }
      else {  // continue to next row on rowA loop
      }
    } // rowA loop
  } // rowR loop
  
  //delete column A
  var spreadsheet = SpreadsheetApp.getActive();  
  spreadsheet.getActiveSheet().deleteColumns(colA);

} // end of function

  
// more elegant solution: use includes function in javascript 
// if attendees rollName n = attendees.includes(rollName[rowR]);



function lastRowofColumn(sheet,column){
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
  // Get the last row with data for the whole sheet.
  var numRows = sheet.getLastRow();
  // Get all data for the given column
  var data = sheet.getRange(1, column, numRows).getValues();
  // Iterate backwards and find first non empty cell
  for(var i = data.length - 1 ; i >= 0 ; i--){
    if (data[i][0] != null && data[i][0] != ""){
      return (i+1);
    }
  }
}

function insertCol(colNum){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().insertColumnsAfter(colNum, 1);
}




