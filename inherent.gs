/*
 * Copyright (c) 2020, salesforce.com, inc.
 * All rights reserved.
 * SPDX-License-Identifier: BSD-3-Clause
 * For full license text, see the LICENSE file in the repo root or https://opensource.org/licenses/BSD-3-Clause
 */
var inherentDataSheet = "Raw Inherent Data";
var likelihoodSheet = "Likelihood";
var impactSheet = "Impact";
var inherentRiskSheet = "Inherent Risk";

function getFactorValues() {
  var likelihood = getLikelihoodFactors();
  var impact = getImpactFactors();
  return likelihood.concat(impact);
}

function getLikelihoodFactors() {
  var ret = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(likelihoodSheet);
  sheet.getRange(1,1,sheet.getLastRow(),1).getValues().map(function(row) { row.map(function(col) { if(ret.indexOf(col) == -1) { ret.push(col); } }); });
  return ret;
}

function getLikelihoodOptions() {
  var result = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(likelihoodSheet);
  var range = sheet.getRange(1,1,sheet.getLastRow(),3);
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    if(values[i][0] != '') {
      result.push({factor: values[i][0], option: values[i][1], helpText: values[i][2]});
    }
  }
  return result;
}

function getImpactFactors() {
  var ret = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(impactSheet);
  sheet.getRange(1,1,sheet.getLastRow(),1).getValues().map(function(row) { row.map(function(col) { if(ret.indexOf(col) == -1) { ret.push(col); } }); });
  return ret;
}

function getImpactOptions() {
  var result = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(impactSheet);
  var range = sheet.getRange(1,1,sheet.getLastRow(),3);
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    if(values[i][0] != '') {
      result.push({factor: values[i][0], option: values[i][1], helpText: values[i][2]});
    }
  }
  return result;
}

function getInherentTooltips() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentDataSheet);
  var ret = [];
  for(var i = 5; i < sheet.getLastColumn(); i++) {
    var r = sheet.getRange(1, i);
    if(r.getValue()) {
      ret[r.getValue().toString()] = r.getNote();
    }
  }
  return ret;
}

function getNewSystem() {
  var factorValues = getFactorValues();
  var ret = { };
  ret['RowId'] = Utilities.getUuid();
  for(var i = 0; i < factorValues.length; i++) {
    ret[factorValues[i].toString()] = null;
  }
  return ret;
}

function getNewReleaseByRowId(rowId) {
  var ret = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentDataSheet);
  var headerValues = getHeaderInherentValues();
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


function getHeaderInherentValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentDataSheet);
  var headerRange = sheet.getRange(1,1,1,sheet.getLastColumn());
  var headerValues = headerRange.getValues();
  return headerValues[0];
}

function findInherentDataRow(rowId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentDataSheet);
  var values = sheet.getRange(1,1,sheet.getLastRow()).getValues();
  for(var i = 0; i < values.length; i++) {
    if(rowId === values[i][0]) {
      return (i + 1);
    }
  }
  return -1;
}

function findInherentRow(rowId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentRiskSheet);
  var values = sheet.getRange(1,1,sheet.getLastRow()).getValues();
  for(var i = 0; i < values.length; i++) {
    if(rowId === values[i][0]) {
      return (i + 1);
    }
  }
  return -1;
}

function findInherentRowByName(systemIdentifier) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentRiskSheet);
  var values = sheet.getRange(1,2,sheet.getLastRow()).getValues();
  for(var i = 0; i < values.length; i++) {
    if(systemIdentifier === values[i][0]) {
      return (i + 1);
    }
  }
  return -1;
}

function getMostRecentInherentSystems() {
  var result = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentDataSheet);
  var values = sheet.getRange(2,1,sheet.getLastRow(),4).getValues();
  for(var i = 0; i < values.length; i++) {
    if(result[values[i][3]] === undefined) {
      result[values[i][3]] = { release: values[i][1], rowId: values[i][0]  };
    } else {
      if(values[i][1] > result[values[i][3]].release) {
        result[values[i][3]] = { release: values[i][1], rowId: values[i][0]  };
      }
    }
  }
  return result;
}

function getInherentSystemByRowId(rowId) {
  var ret = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentDataSheet);
  var headerValues = getHeaderInherentValues();
  var values = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  var flt = values.filter(function(x) { return x[0] === rowId; } );
  if(flt.length == 1) {
    ret = mapRowToSystem(flt[0], headerValues, flt[0]);
  }
  return ret;
}

function processForm(formObject) {
  Logger.log(formObject);
  var ui = SpreadsheetApp.getUi();
  var headerValues = getHeaderInherentValues();
  var sys = mapFormToSystem(formObject, headerValues);
  Logger.log(systemToString(sys));
  saveInherentData(sys);
}

function saveInherentData(systemData) {
  if(systemData["RowId"] && systemData["RowId"].toLowerCase() != "rowid") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(inherentDataSheet);
    var rowNum = findInherentDataRow(systemData["RowId"]);
    if(rowNum == -1) { // we did not find the RowId so we need to append a new row
      rowNum = sheet.getLastRow() + 1;
    }
    if(rowNum > 1) { // save the data here
      var headers = getHeaderInherentValues();
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
  refreshInherentRisk(systemData);
}

function refreshInherentRisk(systemData) {
  var sys = getLatestSystemByName(systemData["System Identifier"]);
  var row = findInherentRowByName(sys["System Identifier"])
  populateInherentRiskRow(row, sys);
}

function testGetLatestSystemByName() {
getLatestSystemByName("baz");
}

function getLatestSystemByName(name) {
  var ret = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentDataSheet);
  var headerValues = getHeaderInherentValues();
  var values = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  var flt = values.filter(x =>  { 
    return x[3] === name; 
  }).sort( (a,b) => { if(a[1] < b[1]) return 1; if(a[1] > b[1]) return -1; return 0; });
  if(flt.length > 0 ) {
    ret = mapRowToSystem(flt[0], headerValues);
  }
  return ret;
}

function getMostRecentSystems() {
  var result = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentDataSheet);
  var values = sheet.getRange(2,1,sheet.getLastRow(),4).getValues();
  for(var i = 0; i < values.length; i++) {
    if(result[values[i][3]] === undefined) {
      result[values[i][3]] = { release: values[i][1], rowId: values[i][0]  };
    } else {
      if(values[i][1] > result[values[i][3]].release) {
        result[values[i][3]] = { release: values[i][1], rowId: values[i][0]  };
      }
    }
  }
  return result;
}


function populateInherentRisk() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentRiskSheet);
  // clear existing data
  var range = sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn());
  range.clear({formatOnly: true, contentsOnly: true});
  var values = getMostRecentSystems();
  var keys = Object.keys(values).sort(function (a,b) { var na = a.toUpperCase(); nb = b.toUpperCase(); if(na < nb) { return -1;} if(na > nb) { return 1; }; return 0; } );
  var rowNum = 2;
  for(var i = 0; i < keys.length; i++) {
    if( values[keys[i]].rowId && keys[i] && !( values[keys[i]].rowId === ""  || keys[i] === "")) {
      var sys = getInherentSystemByRowId(values[keys[i]].rowId);
      populateInherentRiskRow(rowNum, sys);
      rowNum++;
    }    
  } 
}

function populateInherentRiskRow(rowNum, system) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentRiskSheet);
  var headerValues = getHeaderInherentValues();
  var likelihoodFactors = getLikelihoodFactors();
  var impactFactors = getImpactFactors();
  if(rowNum < 0) {
    rowNum = ss.getLastRow() + 1;
  }
  // clear existing data
  var range = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn());
  range.clear({formatOnly: true, contentsOnly: true});
  var keys = Object.keys(system);
  var r = sheet.getRange(rowNum,1,1,3);
  var foo = system["RowId"];
  r.setValues([[system["RowId"], system["System Identifier"], system["Release"]]]);
  var targetCell = sheet.getRange(rowNum,2);
  //targetCell.setNote(systemNote(system["RowId"]));
  sheet.getRange(rowNum,4).setFormula(likelihoodFormula(rowNum, system, likelihoodFactors, headerValues));
  sheet.getRange(rowNum,5).setFormula(ratingFormula(rowNum,"D"));
  sheet.getRange(rowNum,6).setFormula(impactFormula(rowNum, system, impactFactors, headerValues));
  sheet.getRange(rowNum,7).setFormula(ratingFormula(rowNum,"F"));
  sheet.getRange(rowNum,8).setFormula(riskLevelFormula(rowNum, "E", "G"));
  setConditionalFormat(); 
}

function likelihoodFormula(rowNum, system, likelihoodFactors, headerValues) {
  var sumSection = "";
  var countSection = "";
  var maxSection = "";
  var ret = "";
  var sysKeys = Object.keys(system);
  for(var i = 0; i < likelihoodFactors.length; i++) {
    if(sysKeys.indexOf(likelihoodFactors[i]) != -1) {
      var col = headerValues.indexOf(likelihoodFactors[i]) + 1;
      if(col > 0) {
        // formula here
        sumSection += 'SUMIF(LikelihoodTerms,CONCATENATE(INDEX(RawInherentData,1,' + col + '),"~~",INDEX(RawInherentData,MATCH(A' + rowNum + ',SystemIDs,0),' + col + ')),LikelihoodValues)';
        countSection += 'COUNTIF(LikelihoodTerms,CONCATENATE(INDEX(RawInherentData,1,' + col + '),"~~",INDEX(RawInherentData,MATCH(A' + rowNum + ',SystemIDs,0),' + col + ')))';
        maxSection += 'MAXIFS(LikelihoodValues,LikelihoodPinned,CONCATENATE(INDEX(RawInherentData,1,' + col + '),"~~",INDEX(RawInherentData,MATCH(A' + rowNum + ',SystemIDs,0),' + col + ')))';
        if(i < likelihoodFactors.length - 1) {
          sumSection += " + ";
          countSection += " + ";
          maxSection += ", ";
        }
      }
    }
  }
  ret = '=MAX(ROUND((' + sumSection + ') / (' + countSection + '),2), MAX(' + maxSection + '))';
  return ret;
}

function impactFormula(rowNum, system, impactFactors, headerValues) {
  var sumSection = "";
  var countSection = "";
  var maxSection = "";
  var ret = "";
  var sysKeys = Object.keys(system);
  for(var i = 0; i < impactFactors.length; i++) {
    if(sysKeys.indexOf(impactFactors[i]) != -1) {
      var col = headerValues.indexOf(impactFactors[i]) + 1;
      if(col > 0) {
        // formula here
        sumSection += 'SUMIF(ImpactTerms,CONCATENATE(INDEX(RawInherentData,1,' + col + '),"~~",INDEX(RawInherentData,MATCH(A' + rowNum + ',SystemIDs,0),' + col + ')),ImpactValues)';
        countSection += 'COUNTIF(ImpactTerms,CONCATENATE(INDEX(RawInherentData,1,' + col + '),"~~",INDEX(RawInherentData,MATCH(A' + rowNum + ',SystemIDs,0),' + col + ')))';
        maxSection += 'MAXIFS(ImpactValues,ImpactPinned,CONCATENATE(INDEX(RawInherentData,1,' + col + '),"~~",INDEX(RawInherentData,MATCH(A' + rowNum + ',SystemIDs,0),' + col + ')))';
        if(i < impactFactors.length - 1) {
          sumSection += " + ";
          countSection += " + ";
          maxSection += ", ";
        }
      }
    }
  }
  ret = '=MAX(ROUND((' + sumSection + ') / (' + countSection + '),2), MAX(' + maxSection + '))';
  return ret;
}

function setConditionalFormat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentRiskSheet);
  var range = sheet.getRange("H:H");
  var veryHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Very High - 5")
    .setBackground("red")
    .setRanges([range])
    .build();
  var high = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("High - 4")
    .setBackground("orange")
    .setRanges([range])
    .build();
  var moderate = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Moderate - 3")
    .setBackground("yellow")
    .setRanges([range])
    .build();
  var low = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Low - 2")
    .setBackground("#b2dfee")
    .setRanges([range])
    .build();
  var veryLow = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Very Low - 1")
    .setBackground("green")
    .setRanges([range])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(veryHigh);
  rules.push(high);
  rules.push(moderate);
  rules.push(low);
  rules.push(veryLow);
  sheet.setConditionalFormatRules(rules);
}

