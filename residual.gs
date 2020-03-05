/*
 * Copyright (c) 2020, salesforce.com, inc.
 * All rights reserved.
 * SPDX-License-Identifier: BSD-3-Clause
 * For full license text, see the LICENSE file in the repo root or https://opensource.org/licenses/BSD-3-Clause
 */
var residualRiskSheet = "Residual Risk"
var residualDataSheet = "Raw Residual Data"
var residualRiskFactorSheet = "ResidualRiskFactors"
var inherentRiskSheet = "Inherent Risk";

function getResidualFactors() {
  var result = [];
  var seen = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualRiskFactorSheet);
  var range = sheet.getRange(1,1,sheet.getLastRow(),1);
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    if(values[i][0] != '') {
      if(seen.indexOf(values[i][0]) < 0) {
        seen.push(values[i][0]);
        result.push(values[i][0]);
      }
    }
  }
  return result;
}

function getResidualOptions() {
  var result = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualRiskFactorSheet);
  var range = sheet.getRange(1,1,sheet.getLastRow(),3);
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    if(values[i][0] != '') {
      result.push({factor: values[i][0], option: values[i][1], helpText: values[i][2]});
    }
  }
  return result;
}

function getHeaderResidualValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualDataSheet);
  var headerRange = sheet.getRange(1,1,1,sheet.getLastColumn());
  var headerValues = headerRange.getValues();
  return headerValues[0];
}

function getResidualTooltips() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualDataSheet);
  var ret = [];
  for(var i = 5; i < sheet.getLastColumn(); i++) {
    var r = sheet.getRange(1, i);
    if(r.getValue()) {
      ret[r.getValue().toString()] = r.getNote();
    }
  }
  return ret;
}

function loadResidualFromResidualSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualRiskSheet);
  var cell = sheet.getCurrentCell();
  if(null != cell) {
    var rowId = sheet.getRange(cell.getRow(), 1).getValue();
    var sysName = sheet.getRange(cell.getRow(), 2).getValue();
    var release = sheet.getRange(cell.getRow(), 3).getValue();
    if(!(rowId === "RowId" || rowId === '')) { 
      promptToLoadResidual(rowId, sysName, release);
    }
  }
}

function getResidualRowId(rowId) {
  var ret = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualDataSheet);
  var headerValues = getHeaderResidualValues();
  var range = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  var values = range.getValues();
  var flt = values.filter(function(x) { return x[0] === rowId; } );
  if(flt.length == 1) {
    ret = mapRowToSystem(flt[0], headerValues);
  }
  return ret;
}

function newResidualSystem() {
  var headers = getHeaderResidualValues();
  var ret = {};
  for(var i = 0; i < headers.length; i++) {
    ret[headers[i]] = "";
  }
  ret["RowId"] = Utilities.getUuid();
  return ret;
}

function newReleaseFromResidualSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualRiskSheet);
  var cell = sheet.getCurrentCell();
  if(null != cell) {
    var rowId = sheet.getRange(cell.getRow(), 1).getValue();
    var sysName = sheet.getRange(cell.getRow(), 2).getValue();
    var release = sheet.getRange(cell.getRow(), 3).getValue();
    if(!(rowId === "RowId" || rowId === '')) { 
      promptForResidualNewRelease(rowId, sysName, release);
    }
  }
}

function getNewResidualReleaseByRowId(rowId) {
  var ret = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualDataSheet);
  var headerValues = getHeaderResidualValues();
  var range = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  var values = range.getValues();
  var flt = values.filter(function(x) { return x[0] === rowId; } );
  if(flt.length == 1) {
    ret = mapRowToSystem(flt[0], headerValues);
  }
  ret["RowId"] = Utilities.getUuid();
  ret["Release"] = "";
  return ret;
}

function findResidualRow(rowId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualDataSheet);
  var values = sheet.getRange(1,1,sheet.getLastRow()).getValues();
  for(var i = 0; i < values.length; i++) {
    if(rowId === values[i][0]) {
      return (i + 1);
    }
  }
  return -1;
}

function getLatestResidualName(name) {
  var ret = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualRiskSheet);
  var headerValues = getHeaderResidualValues();
  var range = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  var values = range.getValues();
  var flt = values.filter(function(x) { 
    return x[1] === name; 
  }).sort(function (a,b) { if(a[1] < b[1]) return 1; if(a[1] > b[1]) return -1; return 0; });
  if(flt.length > 0 ) {
    ret = mapRowToSystem(flt[0], headerValues);
  }
  return ret;
}

function processResidualForm(formObject) {
  var sys = mapFormToSystem(formObject, getHeaderResidualValues() );
  saveResidualData(sys);
}


function saveResidualData(systemData) {
  if(systemData["RowId"] && systemData["RowId"].toLowerCase() != "rowid") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(residualDataSheet);
    var rowNum = findResidualRow(systemData["RowId"]);
    if(rowNum == -1) { // we did not find the RowId so we need to append a new row
      rowNum = sheet.getLastRow() + 1;
    }
    if(rowNum > 1) { // save the data here
      var headers = getHeaderResidualValues();
      // go through headers and read each matching value from the system data
      for(var i = 0; i < headers.length; i++) {
        if(systemData[headers[i]] != undefined) { // is this a valid value
          var range = sheet.getRange(rowNum, (i + 1));
          range.setValue(systemData[headers[i]])
        }
      }
    }
    // set timestamp to current
    var tsCol = headers.indexOf("Timestamp");
    if(tsCol >= 0) {
      var range = sheet.getRange(rowNum, (tsCol + 1));
      range.setValue(Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd HH:mm:ss"));      
    }
  }
  refreshResidualRow(systemData);
}

function refreshResidualRow(systemData) {
  var rowNum = findResidualRowByName(systemData["System Identifier"]);
  populateResidualRiskRow(rowNum,systemData);
}

function findResidualRowByName(systemIdentifier) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualRiskSheet);
  var values = sheet.getRange(2,2,sheet.getLastRow()).getValues();
  for(var i = 0; i < values.length; i++) {
    if(systemIdentifier === values[i][0]) {
      return (i + 2);
    }
  }
  return -1;
}

function populateResidualRisk() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualRiskSheet);
  // clear existing data
  var range = sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn());
  range.clear({formatOnly: true, contentsOnly: true});
  var values = getMostRecentResiduals();
  var keys = Object.keys(values).sort(function (a,b) { var na = a.toUpperCase(); nb = b.toUpperCase(); if(na < nb) { return -1;} if(na > nb) { return 1; }; return 0; } );
  var rowNum = 2;
  for(var i = 0; i < keys.length; i++) {
    if( values[keys[i]].rowId && keys[i] && !( values[keys[i]].rowId === ""  || keys[i] === "")) {
      var sys = getResidualRowId(values[keys[i]].rowId);
      populateResidualRiskRow(rowNum, sys);
      rowNum++;
    }    
  }
}

function getMostRecentResiduals() {
  var result = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualDataSheet);
  var values = sheet.getRange(2,1,sheet.getLastRow(),7).getValues();
  for(var i = 0; i < values.length; i++) {
    if(result[values[i][4]] === undefined) {
      result[values[i][4]] = { release: values[i][1], rowId: values[i][0]  };
    } else {
      if(values[i][1] > result[values[i][3]].release) {
        result[values[i][4]] = { release: values[i][1], rowId: values[i][0]  };
      }
    }
  }
  return result;
}

function populateResidualRiskRow(rowNum, system) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualRiskSheet);
  var headerValues = getHeaderResidualValues();
  var residualFactors = getResidualFactors();
  if(rowNum < 0) {
    rowNum = sheet.getLastRow() + 1;
  }
  // clear existing data
  var range = sheet.getRange(rowNum,1,1,sheet.getLastColumn());
  var r = sheet.getRange(rowNum,1,1,1);
  r.setValues([ [ system["RowId"] ] ]);//, system["System Identifier"], system["Release"]] ]]);
  sheet.getRange(rowNum, 2).setFormula('=VLOOKUP(VLOOKUP(A' + rowNum + ',RawResidualData,4,FALSE),RawInherentData,4,FALSE)');
  sheet.getRange(rowNum, 3).setFormula('=VLOOKUP(A' + rowNum + ',RawResidualData,2,false)');
  sheet.getRange(rowNum,4).setFormula(inherentRiskScoreFormula(rowNum));
  sheet.getRange(rowNum,5).setFormula(inherentRiskLevelFormula(rowNum));
  sheet.getRange(rowNum,6).setFormula(residualMitigationValueFormula(rowNum, system, residualFactors, headerValues ));
  sheet.getRange(rowNum,7).setFormula(residualRiskScoreFormula(rowNum));
  sheet.getRange(rowNum,8).setFormula(residualRiskLevelFormula(rowNum));
  //var targetCell = sheet.getRange(rowNum,11);
  //targetCell.setNote(residualToDataOnlyString(getResidualRowId(system["RowId"])));
  setResidualConditionalFormat();
}

function inherentRiskScoreFormula(rowNum) {
  return '=VALUE(RIGHT(VLOOKUP(B' + rowNum + ',InherentSystemIdentifier,7,false),1))'
}

function inherentRiskLevelFormula(rowNum) {
  return '=TRIM(LEFT(VLOOKUP(B' + rowNum + ',InherentSystemIdentifier,7,false),FIND(" -",VLOOKUP(B' + rowNum + ',InherentSystemIdentifier,7,false))))';
}

function residualRiskLevelFormula(rowNum) {
  return '=IF(G' + rowNum + ' < 1.01, "Very Low",IF(G' + rowNum + ' < 2.01, "Low", IF(G' + rowNum + ' < 3.01, "Moderate", IF(G' + rowNum + ' < 4.01, "High", "Very High"))))';
}

function residualRiskScoreFormula(rowNum) {
  return '=IF(F' + rowNum + ' > D' + rowNum + ', 1, MIN(5,D' + rowNum + ' - F' + rowNum + '))';
}

function residualMitigationValueFormula(rowNum, system, residualFactors, headerValues) {
  var sumSection = "";
  var ret = "";
  var sysKeys = Object.keys(system);
  for(var i = 0; i < residualFactors.length; i++) {
    if(sysKeys.indexOf(residualFactors[i]) != -1) {
      var col = headerValues.indexOf(residualFactors[i]) + 1;
      if(col > 0) {
        // formula here
        sumSection += 'ROUND( ( VLOOKUP(CONCATENATE(INDEX(RawResidualData,1,' + col + '),"~",INDEX(RawResidualData,MATCH(A' + rowNum + ',ResidualSystemIDs,0),' + col + 
          ')),ResidualFactors,2,false) * VLOOKUP(INDEX(RawResidualData,1,' + col + '),ResidualFactorsAll,7,false) ) / Denominator * 5,2)';
        if(i < residualFactors.length - 1) {
          sumSection += " , ";
        }
      }
    }
  }
  ret = '=SUM(' + sumSection + ')';
  return ret;
}

function setResidualConditionalFormat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualRiskSheet);
  var range1 = sheet.getRange("E:E");
  var range2 = sheet.getRange("H:H");
  var veryHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Very High")
    .setBackground("red")
    .setRanges([range1, range2])
    .build();
  var high = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("High")
    .setBackground("orange")
    .setRanges([range1, range2])
    .build();
  var moderate = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Moderate")
    .setBackground("yellow")
    .setRanges([range1, range2])
    .build();
  var low = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Low")
    .setBackground("#b2dfee")
    .setRanges([range1, range2])
    .build();
  var veryLow = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Very Low")
    .setBackground("green")
    .setRanges([range1, range2])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(veryHigh);
  rules.push(high);
  rules.push(moderate);
  rules.push(low);
  rules.push(veryLow);
  sheet.setConditionalFormatRules(rules);
}
