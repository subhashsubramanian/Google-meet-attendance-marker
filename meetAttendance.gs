/* Mark attendance against a nominal role v3.0- Subhash Subramanian */

// Menu Options 
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Attendance')
  .addItem('Mark Attendance', 'markAtt')
  .addItem('Compile Numbers', 'compileStats')
  .addToUi();
}


function compileStats(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
  // count presents
  var lastCol = ss.getDataRange().getLastColumn() ;
  var thisCell = columnToLetter(lastCol, 2);
  var lastrowR = lastRowofColumn(ss,lastCol);
  ss.getRange(1, lastCol+1, lastrowR, 3).activate();
  ss.getActiveRangeList().setHorizontalAlignment('center');
  
  ss.getRange(1,lastCol+1).activate();
  ss.getRange(1,lastCol+1).setValue("Presents");
  ss.getRange(2,lastCol+1).activate();
  
  //Browser.msgBox(thisCell)
  ss.getCurrentCell().setFormula('=countif((C2:' + thisCell + '),"P")');
  ss.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  ss.getActiveRangeList().setHorizontalAlignment('center');

  
  // count absents
  //lastCol = ss.getDataRange().getLastColumn();
  ss.getRange(1,lastCol+2).activate();
  ss.getRange(1,lastCol+2).setValue("Absents");
  ss.getRange(2,lastCol+2).activate();
  //thisCell = columnToLetter(lastCol, 2);
  ss.getCurrentCell().setFormula('=countif((C2:' + thisCell + '),"A")');
  //ss.getRange(2,lastCol+2).activateAsCurrentCell();
  //ss.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  ss.getActiveRange().autoFill(ss.getRange(2,lastCol+2,lastrowR-1,1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  
  
  // count totals
  //lastCol = ss.getDataRange().getLastColumn();
  ss.getRange(1,lastCol+3).activate();
  ss.getRange(1,lastCol+3).setValue("Total");
  ss.getRange(2,lastCol+3).activate();
  //thisCell = columnToLetter(lastCol, 2);
  ss.getCurrentCell().setFormula('=counta(C2:' + thisCell + ')');
  //ss.getRange(2,lastCol+3).activateAsCurrentCell();
  //ss.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  ss.getActiveRange().autoFill(ss.getRange(2,lastCol+3,lastrowR-1,1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}


function markAtt(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var toSheetName = ss.getSheetName();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to mark attendance in the sheet: ' + toSheetName + ' ?', ui.ButtonSet.YES_NO);
  if (response == response.NO) {
     Browser.msgBox("Run the function while on the sheet in which you wish to mark attendance")
   } 
  else {
  // copy from all sheets with names with last 10 charaters same as this sheet
  var lastCol = ss.getDataRange().getLastColumn();
  var colToPaste = lastCol;
  copyFromAllSheets(toSheetName);
  // mark att for each column
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(toSheetName).activate();
  ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  markAllCols(ss);
  }
}

function markAllCols(ss) {  
  //get last row of Roll column
  //get lastrow of Roll column
  var lastrowR = lastRowofColumn(ss,1);
  var lastCol = ss.getDataRange().getLastColumn();
  var thisCol = lastCol;
  var firstCol = 2;
  var firstRow = 2;
  var curvalueColA = ss.getRange(1,thisCol).getValue().toString();
  while (curvalueColA.substring(curvalueColA.length-5) !== "_Done" && thisCol > firstCol){
    markthisCol(ss,firstRow,thisCol,lastrowR);
    thisCol = thisCol-1;
    curvalueColA = ss.getRange(1,thisCol).getValue().toString();
    lastCol = ss.getDataRange().getLastColumn();
  }
  
  ss.getRange(firstRow, firstCol+1, lastrowR-firstRow+1, lastCol-firstCol).activate();
  ss.getActiveRangeList().setHorizontalAlignment('center');
  ss.getRange(1,lastCol).activate();
  ss.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    
  Browser.msgBox("Attendance Marked");  
  // Second check if (ss.getRange[3][thisCol]) =="P" or ss.getRange[3][thisCol]) =="A" ) :
  //tell marking attendence is done and go to next col.
}


function markthisCol(ss,firstRow, thisCol,lastrowR){
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
  //var curvalueColAR2 = ss.getRange(2,colA).getValue();
  //ss.getRange(1,colA).setValue(curvalueColA + "_Done");
  ss.getRange(1,colA+1).setValue(curvalueColAR1 + "_Done");
  //ss.getRange(2,colA+1).setValue(curvalueColAR2);  
  
  /* double loop- at the end- now replaced
  */
  
  //faster method: use array method- check each cell of Roll call column with Attendees array
   var AttArray = flatten(ss.getRange(firstRow,colA,lastrowA-firstRow+1).getValues());
   for (rowR=firstRow; rowR < lastrowR+1; rowR++){ //rowR loop
    lastCol = ss.getDataRange().getLastColumn();
   //Browser.msgBox(AttArray.indexOf(ss.getRange(rowR, colR).getValue())) 
   if(AttArray.indexOf(ss.getRange(rowR, colR).getValue()) > -1) {
     //mark present or Absent
     ss.getRange(rowR, colA+1).setValue("P");   // Present
    }
    else {
    ss.getRange(rowR, colA+1).setValue("A");   // Absent
    }
  }
  
  //delete column A
  var spreadsheet = SpreadsheetApp.getActive();  
  spreadsheet.getActiveSheet().deleteColumns(colA);

} // end of function

  

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


// Takes and array of arrays matrix and return an array of elements.
function flatten(arrayOfArrays){
  return [].concat.apply([], arrayOfArrays);
}


function copyFromAllSheets(toSheetName){
  var file = SpreadsheetApp.getActiveSpreadsheet();
  // get names of all sheets in the file
  var sheetList = getAllSheetNames(file)
  // Browser.msgBox(sheetList)
  // choose relevant sheets = fromSheets
  var copyFromSheets = retainSameSheetNames(sheetList,toSheetName);
  // Browser.msgBox(copyFromSheets)
  // loop to copy required cells from each fromSheet and delete that fromSheet
  for (var i = 0; i < copyFromSheets.length; i++){
    var fromSheetName = copyFromSheets[i];
    var toSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(toSheetName);
    var colToPaste = toSheet.getDataRange().getLastColumn()+1; 
    copyfromSheet(fromSheetName,toSheetName,colToPaste);
    //var colToPaste = toSheet.getDataRange().getLastColumn()
    //delete the fromsheet after copying data
    deleteThisSheet(fromSheetName);
   }
  file.getSheetByName(toSheetName).activate();
  }


function getAllSheetNames(file){  
  // get names of all sheets in the file
  var sheetList = [];
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
    file.getSheets().forEach(function(val){
       sheetList.push(val.getName())
    });
  return flatten(sheetList);
}


function retainSameSheetNames(sheetList,toSheetName){
  // delete sheetnames that dont have name ending with the same 10 characters
  // also delete toSheetName from the array
  // since meet has 10 alphabets
  var i = 0;
  var toSheetLast10 = toSheetName.substring(toSheetName.length-10); 
  var numSheets = sheetList.length
  while (i < numSheets) {
    var fromSheetLast10 = sheetList[i].substring(sheetList[i].length-10); 
    if (sheetList[i] === toSheetName){
      sheetList.splice(i, 1);
      numSheets = sheetList.length
    } else if (fromSheetLast10 === toSheetLast10) {
      ++i;
    } else {
      sheetList.splice(i, 1);
      numSheets = sheetList.length
    }
  }
  // reverse order
  sheetList = sheetList.reverse()
  return sheetList;
}


function copyfromSheet(fromSheetName,toSheetName,colToPaste){
  // copy from relevant cells in source sheets 
  // paste to relevant cells in destination sheets
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(fromSheetName), true);
  var fromSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(fromSheetName);
  var lastAttendeeRow = lastRowofColumn(fromSheet,1);
  spreadsheet.getRange('A2').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(toSheetName), true);
  // convert colNum to alphabet?
  SpreadsheetApp.getActive().getActiveSheet().getRange(2,colToPaste).activate()
  // Browser.msgBox("\'" + fromSheetName + "\'" + "!A2:A" + lastAttendeeRow);
  spreadsheet.getRange("\'" + fromSheetName + "\'" + "!A2:A" + lastAttendeeRow).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  // paste date in first row 10 characters mm/dd/yyyy
  var dateval = fromSheetName.substring(0,10);
  SpreadsheetApp.getActive().getActiveSheet().getRange(1,colToPaste).setValue(dateval);
  
}


function deleteThisSheet(sheetName){
   // delete this source sheet
 var file = SpreadsheetApp.getActiveSpreadsheet();
 var delSheet = file.getSheetByName(sheetName); 
 file.deleteSheet(delSheet);
}


/* double loop
for (rowR=firstRow; rowR < lastrowR+1; rowR++){ //rowR loop
    lastCol = ss.getDataRange().getLastColumn();
    //Browser.msgBox("I am on Roll row " + rowR + " and Att row " + rowA)
    for (rowA=firstRow; rowA < lastrowA+1; rowA++){ //rowA loop
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
 */ 


function columnToLetter(column, row) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter + row;
}


