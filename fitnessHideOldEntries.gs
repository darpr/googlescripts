//Author: Pradeep Sivakumar
//Date Created: 29 Nov 2010
//Purpose: Hide all past columns in 'log' sheet, except yesterday's and today's data.

var SS = SpreadsheetApp.getActiveSpreadsheet();
var SHEET_LOG = 'log';
var TODAY = setMidnightToday();

function onOpen() {
  var sheetLog = SS.getSheetByName(SHEET_LOG);
  var dataStartCol = 3;
  var dataDate = sheetLog.getRange(1, dataStartCol).getValue();

  Logger.log("Start Date: " + dataDate + ", End Date: " + TODAY);
  
  //4 cols per day's entry
  var numColsToHide = (daysDifference(dataDate, TODAY) - 1) * 4;
  Logger.log("Number of columns to hide: " + numColsToHide);
  
  //clear mesg display (A1)
  sheetLog.hideColumns(3, numColsToHide);
  showMesg("A1", "", "white");
}

function daysDifference(date1, date2){
  var days_diff = (date2.getTime() - date1.getTime())/(1000*60*60*24);
  Logger.log("Difference between two dates in days: " + days_diff);

  //return difference in days based on millisecond calculation
  return ((date2.getTime() - date1.getTime())/(1000*60*60*24));
}

function setMidnightToday(){
  //indicate start of script. fill A1 with mesg
  showMesg("A1", "loading", "yellow");
  
  //get date elements 
  var now = new Date();
  var t_yyyy = now.getFullYear();
  var t_mm = now.getMonth(); //runs 0 to 11
  var t_dd = now.getDate();
  Logger.log("date: " + now + ", " + t_yyyy + ", " + t_mm + ", " + t_dd);
  
  //manipulate current date to set clock at 00:00:00
  var now_edit = new Date(t_yyyy, t_mm, t_dd, 00, 00, 00, 00);
  
  return now_edit;
}

function showMesg(cell, mesg, bgcolor){
  var sheetLog = SS.getSheetByName(SHEET_LOG);

  //set mesg and colorize
  sheetLog.setActiveCell(cell).setValue(mesg);
  sheetLog.setActiveCell(cell).setBackgroundColor(bgcolor);
}
