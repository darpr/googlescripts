//Author: Pradeep Sivakumar
//Date Created: 28 Nov 2010
//Purpose: Publish "active days" for each member, accounting for blank cells and 0 rows.
//Note:
//1. GSpreadsheet bug - Logger doesn't work if focus shifts from editor to spreadsheet, so don't use Browser.msgBox();

var SS = SpreadsheetApp.getActiveSpreadsheet();
var SHEET_LOG = 'log';
var SHEET_LBOARD = 'lboard';
var SHEET_PROFILES = 'db.profiles';
var TODAY = new Date();

function fillActiveDays() {
  var sheetLog = SS.getSheetByName(SHEET_LOG);
  var sheetLboard = SS.getSheetByName(SHEET_LBOARD);
  
  //indicate start of script. fill A1 with "hold"
  sheetLog.setActiveCell("A1").setValue("hold");
  sheetLog.setActiveCell("A1").setBackgroundColor("yellow");
  
  var memberRow;
  var totalMembers = getMembership();
  Logger.log("Total membership: " + getMembership());
  
  try {
  for (memberRow=3;memberRow<=totalMembers+2;memberRow++){
       Logger.log("memberRow: " + memberRow);
  
       //todo: identify dataStartCol (first col with data value).
       var dataStartCol = 3;
       var dataDate = sheetLog.getRange(1, dataStartCol).getValue();
       var daysActive = 0;
            
       while (dataDate <= TODAY){
         Logger.log("dataDate: " + dataDate);
         
         var dataSQ = sheetLog.getRange(memberRow, dataStartCol).getValue();
         var dataPU = sheetLog.getRange(memberRow, dataStartCol + 1).getValue();
         var dataCR = sheetLog.getRange(memberRow, dataStartCol + 2).getValue();
         var dataCA = sheetLog.getRange(memberRow, dataStartCol + 3).getValue();
    
         Logger.log("dataSQ: " + dataSQ + ", dataPU: " + dataPU + ", dataCR: " + dataCR + ", dataCA: " + dataCA);
         
         if ((dataSQ != "" && dataSQ > 0) || (dataPU != "" && dataPU > 0) || (dataCR != "" && dataCR > 0) || (dataCA != "" && dataCA > 0)){
           daysActive++;
         } 
         
         dataStartCol = dataStartCol + 4;
         dataDate = sheetLog.getRange(1, dataStartCol).getValue();  
       }
       Logger.log("Member " + memberRow + " has been active for " + daysActive + " days!\n\n");
  
       sheetLboard.getRange(memberRow-1, 2).setValue(daysActive);
    }
  }
  catch (e) {
    //MailApp.sendEmail(recipient, subject, body);
  }

  //indicate end of script. clear A1
  sheetLog.setActiveCell("A1").setValue("");       
  sheetLog.setActiveCell("A1").setBackgroundColor("white");
}
      
function getMembership(){
   var sheetProfiles = SS.getSheetByName(SHEET_PROFILES);
   var totalMembers = sheetProfiles.getLastRow() - 1;
   return totalMembers; 
}
