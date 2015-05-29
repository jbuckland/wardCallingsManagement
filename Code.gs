var callingsSheetName ="Callings";
var callingsArchiveSheetName = "Archived Callings";
var releasesSheetName = "Releases";
var firstFormDataColumn = "E"



var statusColId=1;
var bishopColId = 2;
var firstColId = 3;
var secondColId = 4;
var individualColId = 6;
var auxiliaryColId = 7;
var dateExtendedColId = 12;
var dateSustainedColId = 13;
var dateSetApartColId = 14;
var dateEnteredInMLSColId = 15;

var callingsColumnToSortBy = [{ column : statusColId, ascending: true },{ column : auxiliaryColId, ascending: true } ];
var releasesColumnToSortBy =  [{ column : statusColId, ascending: true }];


var statusNeedsApproval = "1. Needs Approval";
var statusNeedsToBeCalled = "2. Needs To Be Called";
var statusNeedsToBeSustained = "3. Needs To Be Sustained";
var statusNeedsToBeSetApart = "4. Needs To Be Set Apart";
var statusEnterInMLS = "5. Enter in MLS";
var statusComplete = "6. Complete";
var statusDeclined = "7. Declined";

var backgroundNeedsApproval = "white";
var backgroundNeedsToBeCalled = "#f9cb9c";
var backgroundNeedsToBeSustained = "#ffe599";
var backgroundNeedsToBeSetApart = "#9fc5e8";
var backgroundEnterInMLS = "#b4a7d6";
var backgroundComplete = "#b6d7a8";
var backgroundDeclined = "#ea9999";


//sheet events
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("Custom Functions",
             [{ name: "Archive Completed and Declined", functionName: "archive" }]);
}

function onEdit(e){
  Logger.log("onEdit() called");
  var sheet = e.source.getActiveSheet();  
  if(sheet.getName() == callingsSheetName && e.range.getRow()>1){     
    //send the whole row of the changed cell    
    for(var i=0; i< e.range.getNumRows(); i++){
      var currentRow = e.range.getRow()+i;
      Logger.log("currentRow: "+currentRow);
      var row = sheet.getRange(currentRow,1,currentRow,e.source.getLastColumn());    
      setCallingStatus(row);  
    }   
    
    sortSheetWithHeader(sheet, callingsColumnToSortBy);
    
  }else if (sheet.getName() == releasesSheetName) {
    sortSheetWithHeader(sheet, releasesColumnToSortBy);
    
  }  
  else{
    Logger.log("unknown sheet that was edited");
  }
}

function onSubmit(e){
  Logger.log("onSubmit() called");
  var target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(callingsSheetName);
  addRow(target_sheet);
  
  var target_range = target_sheet.getRange(firstFormDataColumn + target_sheet.getLastRow());
  e.range.copyTo(target_range, {contentsOnly:true});
  
  var row = getRow(target_sheet, target_range.getRow());
  setCallingStatus(row);
  sortSheetWithHeader(target_sheet, callingsColumnToSortBy);  
}
//////////////

function archive(){
  Logger.log("archive()");
  var srcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(callingsSheetName);
  var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(callingsArchiveSheetName);
  var statusValues = srcSheet.getRange(1, statusColId, srcSheet.getLastRow(), statusColId).getValues();
  
  Logger.log("found "+statusValues.length+" statuses");
  //go backwards because we'll be removing rows as we go
  for(var i=statusValues.length; i > 0; i--){    
    if(statusValues[i] == statusComplete || statusValues[i] == statusDeclined){
      
      Logger.log("archiving row "+ i+1);
      var row = getRow(srcSheet, i+1);
      addRow(destSheet);      
      var destRange = destSheet.getRange(destSheet.getLastRow(),1);
      row.moveTo(destRange);
      srcSheet.deleteRow(i+1);
    }
    
  }
  
}

function getRow(sheet, rowId){
  return sheet.getRange(rowId,1,rowId,sheet.getLastColumn());    
}

function addRow(target_sheet) {
  var sh = target_sheet, lRow = sh.getLastRow(); 
  var lCol = sh.getLastColumn(), range = sh.getRange(lRow,1);
  sh.insertRowsAfter(lRow, 1);
  range.copyTo(sh.getRange(lRow+1, 1), {contentsOnly:false});
}



function sortSheetWithHeader(sheet, sortOrders){
  var sortRange = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn());
  sortRange.sort( sortOrders );
  
}

function setCallingStatus(range){
  Logger.log("setCallingStatus() called");
  var rowId = range.getRow(); 
  
  var bishopApproval = range.getCell(1,bishopColId).getValue().toUpperCase();
  var firstApproval = range.getCell(1,firstColId).getValue().toUpperCase();
  var secondApproval = range.getCell(1,secondColId).getValue().toUpperCase();
  
  var dateExtended=range.getCell(1,dateExtendedColId).getValue();
  var dateSustained=range.getCell(1,dateSustainedColId).getValue();
  var dateSetApart=range.getCell(1,dateSetApartColId).getValue();
  var dateEnteredInMLS=range.getCell(1,dateEnteredInMLSColId).getValue();
  
  var status = "unknown"; 
  var background = "white";
  if(bishopApproval == "Y" && firstApproval == "Y" && secondApproval == "Y"){   
    if(dateExtended != 0){      
      if(dateSustained != 0){        
        if(dateSetApart != 0){         
          if(dateEnteredInMLS != 0 ){
            status = statusComplete;
            background = backgroundComplete;
          }else{
            status = statusEnterInMLS;
            background = backgroundEnterInMLS;
          }
          
        }else{
          status = statusNeedsToBeSetApart;        
          background = backgroundNeedsToBeSetApart;
        }
        
      }else {
        status = statusNeedsToBeSustained;
        background = backgroundNeedsToBeSustained;
      }      
      
    }else{
      status = statusNeedsToBeCalled;
      background = backgroundNeedsToBeCalled;
    }
  }else if(bishopApproval == "N" && firstApproval == "N" && secondApproval == "N"){
    status = statusDeclined;
    background = backgroundDeclined;
  }
  else {
    status = statusNeedsApproval; 
    background = backgroundNeedsApproval;
  }
  
  var statusCell= range.getCell(1,statusColId);
  statusCell.setValue(status);
  statusCell.setBackground(background);
}




