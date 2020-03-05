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

function mapRowToSystem(rowValues, headerValues) {
  var ret = {};
  for(var i = 0; i < headerValues.length; i++) {
    if(i < rowValues.length) {
      ret[headerValues[i]] = rowValues[i];
    } else {
      ret[headerValues[i]] = null;
    }
  }
  return ret;
}

function mapFormToSystem(systemForm, headerValues) {
  var ret = {};
  for(var i = 0; i < headerValues.length; i++) {
    if(systemForm[headerValues[i]]) {
      ret[headerValues[i].toString()] = systemForm[headerValues[i].toString()];
    } else {
      ret[headerValues[i].toString()] = null;
    }
  }
  return ret;
}


function systemToString(systemData) {
  var st = "systemData:";
  for(var i = 0; i < Object.keys(systemData).length; i++) {
    st += Object.keys(systemData)[i] + " / " + systemData[Object.keys(systemData)[i]] + "\n";
  }
  return st;
}

function ratingFormula(rowNum, colLetter) {
  return '=VLOOKUP(' + colLetter + rowNum + ',{1,"Very Low - 1";1.50,"Low - 2";2.50,"Moderate - 3";3.50,"High - 4";4.50,"Very High - 5"},2,1)'
}

function riskLevelFormula(rowNum, likelihoodRatingCol, impactRatingCol) {
  return '=INDEX(RatingTable,MATCH(' + likelihoodRatingCol + rowNum + ',RatingLikelihood,0), MATCH(' + impactRatingCol + rowNum + ',RatingImpact,0))';
}

function inherentWithNoResidual() {
  var inh = listInherentSystemNames();
  var res = listResidualSystemNames();
  var ret = inh.filter(function (name) { return res.indexOf(name) == -1 }).map(function (value, index, array) { return value.replace("'","").replace('"','')}).sort(); 
  return ret;
}

function listInherentSystemNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentRiskSheet);
  var values = sheet.getRange(2,1,sheet.getLastRow(),2).getValues();
  var ret = [];
  for(var i = 0; i < values.length; i++) {
    if(values[i][1] && values[i][1] != "" && ret.indexOf(values[i][1]) == -1) {
       ret.push(values[i][1]);
    }
  }
  return ret;
}

function listResidualSystemNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(residualRiskSheet);
  var values = sheet.getRange(2,1,sheet.getLastRow(),2).getValues();
  var ret = [];
  for(var i = 0; i < values.length; i++) {
    if(values[i][1] && values[i][1] != "" && ret.indexOf(values[i][1]) == -1) {
       ret.push(values[i][1]);
    }
  }
  return ret;
}

