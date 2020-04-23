/*
 * Copyright (c) 2020, salesforce.com, inc.
 * All rights reserved.
 * SPDX-License-Identifier: BSD-3-Clause
 * For full license text, see the LICENSE file in the repo root or https://opensource.org/licenses/BSD-3-Clause
 */
var _version = "1.0.0";
var inherentDataSheet = "Raw Inherent Data";
var likelihoodSheet = "Likelihood";
var impactSheet = "Impact";
var inherentRiskSheet = "Inherent Risk";
var residualRiskSheet = "Residual Risk";
var residualDataSheet = "Raw Residual Data";
var residualRiskFactorsSheet = "ResidualRiskFactors";
var ratingLookupSheet = "RatingLookup";
var loadImage = "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";
var newReleaseImage = "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";

function Initialize() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  verifyTabs();
  namedRanges();  
}
 
function verifyTabs() {
  inherentRiskTab();
  likelihoodTab();
  impactTab();
  inherentDataTab();
  ratingLookupTab();
  residualFactorsTab();
  residualDataTab();
  residualRiskTab();
}

function addTab(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if(null == sheet) {
    sheet = ss.insertSheet();
    sheet.setName(name)
    sheet.addDeveloperMetadata("ID",name);
  }
  return sheet;
}

function likelihoodTab() {
  var sheet = addTab(likelihoodSheet);
  var c = sheet.getRange(1,1);
  c.setNote("Likelihood Factor");
  c = sheet.getRange(1,2);
  c.setNote("Factor Option");
  c = sheet.getRange(1,3);
  c.setNote("Option Help Text");
  c = sheet.getRange(1,4);
  c.setNote("Match String");
  c.setFormula('=CONCATENATE(A1,"~",B1)');
  c = sheet.getRange(1,5);
  c.setNote("Option Risk Value");
  c = sheet.getRange(1,6);
  c.setNote("Option Risk Pinned Match String");  
}

function impactTab() {
  var sheet = addTab(impactSheet);
  var c = sheet.getRange(1,1);
  c.setNote("Impact Factor");
  c = sheet.getRange(1,2);
  c.setNote("Factor Option");
  c = sheet.getRange(1,3);
  c.setNote("Option Help Text");
  c = sheet.getRange(1,4);
  c.setNote("Match String");
  c.setFormula('=CONCATENATE(A1,"~",B1)');
  c = sheet.getRange(1,5);
  c.setNote("Option Risk Value");
  c = sheet.getRange(1,6);
  c.setNote("Option Risk Pinned Match String");  
}

function inherentDataTab() {
  var sheet = addTab(inherentDataSheet);
  var c = sheet.getRange(1,1);
  c.setNote("RowId");
  c.setValue([["RowId"]]);
  var c = sheet.getRange(1,2);
  c.setNote("Release");
  c.setValue([["Release"]]);
  var c = sheet.getRange(1,3);
  c.setNote("Timestamp");
  c.setValue([["Timestamp"]]);
  var c = sheet.getRange(1,4);
  c.setNote("System Identifier");
  c.setValue([["System Identifier"]]);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);
  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("B:B").setNumberFormat("@");
}

function inherentRiskTab() {
  var sheet = addTab(inherentRiskSheet);
  var c = sheet.getRange(1,1);
  c.setValue([["RowId"]]);
  c.setNote("RowId");
  var c = sheet.getRange(1,2);
  c.setValue([["System Identifier"]]);
  c.setNote("System Identifier");
  var c = sheet.getRange(1,3);
  c.setValue([["Latest Release"]]);
  c.setNote("Latest Release");
  var c = sheet.getRange(1,4);
  c.setValue([["Likelihood"]]);
  c.setNote("Likelihood");
  var c = sheet.getRange(1,5);
  c.setValue([["Likelihood Rating"]]);
  c.setNote("Likelihood Rating");
  var c = sheet.getRange(1,6);
  c.setValue([["Impact"]]);
  c.setNote("Impact");
  var c = sheet.getRange(1,7);
  c.setValue([["Impact Rating"]]);
  c.setNote("Impact Rating");
  var c = sheet.getRange(1,8);
  c.setValue([["Risk Level"]]);
  c.setNote("Risk Level");
  sheet.setFrozenRows(1);
  sheet.hideColumns(1);
  sheet.getRange("C:C").setNumberFormat("@");
  var imageBlob = Utilities.newBlob(Utilities.base64Decode(loadImage), 'image/png', 'loadImage');
  var loadImg = sheet.insertImage(loadImage, 9,1);
  loadImg.assignScript("loadSystemFromInherentSheet");
  imageBlob = Utilities.newBlob(Utilities.base64Decode(newReleaseImage), 'image/png', 'newReleaseImage');
  var newRelImg = sheet.insertImage(newReleaseImage, 10, 1);
  newRelImg.assignScript("newReleaseFromInherentSheet");  
}

function residualFactorsTab() {
  var sheet = addTab(residualRiskFactorsSheet);
  var c = sheet.getRange(1,1);
  c.setNote("Residual Factor");
  c = sheet.getRange(1,2);
  c.setNote("Factor Option");
  c = sheet.getRange(1,3);
  c.setNote("Option Help Text");
  c = sheet.getRange(1,4);
  c.setNote("Match String");
  c.setFormula('=CONCATENATE(A1,"~",B1)');
  c = sheet.getRange(1,5);
  c.setNote("Option Risk Value");
  c = sheet.getRange(1,6);
  c.setNote("Option Risk Pinned Match String");
  c = sheet.getRange(1,7);
  c.setNote("Factor Weightage Value (only for the first instance of factor)");
  c = sheet.getRange(1,9,1,2);
  c.setValues([["Residual Risk Denominator", "80"]]);
}

function residualDataTab() {
  var sheet = addTab(residualDataSheet);
  var c = sheet.getRange(1,1);
  c.setNote("RowId");
  c.setValue([["RowId"]]);
  var c = sheet.getRange(1,2);
  c.setNote("Release");
  c.setValue([["Release"]]);
  var c = sheet.getRange(1,3);
  c.setNote("Timestamp");
  c.setValue([["Timestamp"]]);
  var c = sheet.getRange(1,4);
  c.setNote("InherentId");
  c.setValue([["InherentId"]]);  
  var c = sheet.getRange(1,5);
  c.setNote("System Identifier");
  c.setValue([["System Identifier"]]);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(5);
  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("B:B").setNumberFormat("@");
}

function residualRiskTab() {
  var sheet = addTab(residualRiskSheet);
  var c = sheet.getRange(1,1);
  c.setValue([["RowId"]]);
  c.setNote("RowId");
  var c = sheet.getRange(1,2);
  c.setValue([["System Identifier"]]);
  c.setNote("System Identifier");
  var c = sheet.getRange(1,3);
  c.setValue([["Latest Release"]]);
  c.setNote("Latest Release");
  var c = sheet.getRange(1,4);
  c.setValue([["Inherent Risk Score"]]);
  c.setNote("Iherent Risk Score");
  var c = sheet.getRange(1,5);
  c.setValue([["Inherent Risk Level"]]);
  c.setNote("Inherent Risk Level");
  var c = sheet.getRange(1,6);
  c.setValue([["Mitigation Value "]]);
  c.setNote("Mitigation Value");
  var c = sheet.getRange(1,7);
  c.setValue([["Residual Risk Score"]]);
  c.setNote("Residual Risk Score");
  var c = sheet.getRange(1,8);
  c.setValue([["Residual Risk Level"]]);
  c.setNote("Residual Risk Level");
  sheet.setFrozenRows(1);
  sheet.hideColumns(1);
  sheet.getRange("C:C").setNumberFormat("@");
  var imageBlob = Utilities.newBlob(Utilities.base64Decode(loadImage), 'image/png', 'loadImage');
  var loadImg = sheet.insertImage(loadImage, 9,1);
  loadImg.assignScript("loadResidualFromResidualSheet");
  imageBlob = Utilities.newBlob(Utilities.base64Decode(newReleaseImage), 'image/png', 'newReleaseImage');
  var newRelImg = sheet.insertImage(newReleaseImage, 10, 1);
  newRelImg.assignScript("newReleaseFromResidualSheet");  
}


var ratingLookup = [ ["", "Very Low - 1", "Low - 2", "Moderate - 3", "High - 4", "Very High - 5" ],
                     ["Very High - 5", "Low - 2", "Moderate - 3", "High - 4", "Very High - 5", "Very High - 5" ],
                     ["High - 4", "Very Low - 1", "Low - 2", "Moderate - 3", "High - 4", "Very High - 5" ],
                     ["Moderate - 3", "Very Low - 1", "Very Low - 1", "Low - 2", "Moderate - 3", "High - 4" ],
                     ["Low - 2", "Very Low - 1", "Very Low - 1", "Very Low - 1", "Low - 2", "Moderate - 3" ],
                     ["Very Low - 1", "Very Low - 1", "Very Low - 1", "Very Low - 1", "Very Low - 1", "Low - 2" ],
                     ["Likelihood", "", "", "Impact", "", ""]
                   ];

function ratingLookupTab() {
  var sheet = addTab(ratingLookupSheet);
  sheet.clear({formatOnly: true, contentsOnly: true});
  var range = sheet.getRange(1,1,7,6);
  range.setValues(ratingLookup);
  sheet.hideSheet();
}

function namedRanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ids = ss.getSheetByName(inherentDataSheet);
  var ranges = ss.getNamedRanges();
  ss.setNamedRange("RawInherentData", ids.getRange("A:Z"));
  ss.setNamedRange('SystemIDs', ids.getRange("A:A"))
  var rls = ss.getSheetByName(ratingLookupSheet);
  ss.setNamedRange("RatingImpact", rls.getRange("1:1"));
  ss.setNamedRange("RatingLikelihood", rls.getRange("A:A"));
  ss.setNamedRange("RatingTable", rls.getRange("A1:F6"));
  var ls = ss.getSheetByName(likelihoodSheet);
  ss.setNamedRange("LikelihoodPinned", ls.getRange("F:F"));
  ss.setNamedRange("LikelihoodTerms", ls.getRange("D:D"));
  ss.setNamedRange("LikelihoodValues", ls.getRange("E:E"));
  var is = ss.getSheetByName(impactSheet)
  ss.setNamedRange("ImpactPinned", is.getRange("F:F"));
  ss.setNamedRange("ImpactTerms", is.getRange("D:D"));
  ss.setNamedRange("ImpactValues", is.getRange("E:E"));
  var irs = ss.getSheetByName(inherentRiskSheet);
  ss.setNamedRange("InherentSystemIdentifier", irs.getRange("B:H"));
  var rds = ss.getSheetByName(residualDataSheet);
  ss.setNamedRange("RawResidualData", rds.getRange("A:Z"));
  ss.setNamedRange("ResidualSystemIDs", rds.getRange("A:A"));
  var rfs = ss.getSheetByName(residualRiskFactorsSheet);
  ss.setNamedRange("ResidualFactors", rfs.getRange("D:G"));
  ss.setNamedRange("ResidualFactorsAll", rfs.getRange("A:G"));
  ss.setNamedRange("Denominator", rfs.getRange("J1"));
}

var sampleResidualFactors = [
  ["Technical debt", "Very Low", "The system has a very low level of technical debt", "", "-1", "", "8"],
  ["Technical debt", "Low", "The system has a low level of technical debt", "", "-2", "", ""],
  ["Technical debt", "Moderate", "The system has a moderate level of technical debt", "", "-3", "", ""],
  ["Technical debt", "High", "The system has a high  level of technical debt", "", "-4", "", ""],
  ["Technical debt", "Very High", "The system has a very high level of technical debt", "", "-5", "", ""],
  ["Security Coverage", "None", "Coverage is unavailable or ineffective", "", "0", "", "10"],
  ["Security Coverage", "Some", "Coverage is present and effective", "", "0.5", "", ""],
  ["Security Coverage", "Full", "Coverage is present and highly effective/proactive", "", "1", "", ""]
];

var sampleLikelihoodFactors = [
  ["Who has access to the system?", "Restricted Internal Only", "Only a limited set of Salesforce personnel have access to the system", "", "1"],
  ["Who has access to the system?", "Internal only", "Only internal  personnel have access to the system", "", "2"],
  ["Who has access to the system?", "Partner users", "Internal personnel and a limited set of external entities have access to the system (third party, contractors,  submitters, etc.)", "", "3"],
  ["Who has access to the system?", "Authenticated customers", "Identified entities that have an authorized use of the system", "", "4"],
  ["Who has access to the system?", "Anonymous Internet users", "Open to unauthenticated access", "", "5"],
  ["How exposed is the knowledge of the system?", "Confidential", "Only a limited number of internal personnel have information about the system", "", "1"],
  ["How exposed is the knowledge of the system?", "Internal", "Only internal personnel have information about the system", "", "2"],
  ["How exposed is the knowledge of the system?", "Partner", "Internal personnel and a limited number of external entities have information about the system", "", "3"],
  ["How exposed is the knowledge of the system?", "Public knowledge (public documentation, open source, etc.)", "Some level of detailed information about the sytem is available to the general public", "", "5"]
];

var sampleImpactFactors = [
  ["What types of information could be disclosed?", "Public data", "Data that is generally available to non-authenticated users", "", "1"],
  ["What types of information could be disclosed?", "User/partner data or metadata", "Data used by authenticated non-adminstrative users to perform business requirements", "", "3"],
  ["What types of information could be disclosed?", "Administrative data or metadata", "Data used by authenticated adminstrative users to perform administrative tasks", "", "4"],
  ["What types of information could be disclosed?", "Authentication secrets", "Sensitive and compartmantilized data used for resticted purposes", "", "5"],
  ["What types of information could be disclosed?", "Compliance Data", "Sensitive data with restricted access used to perform business requirements, such as HIPAA, PCI, etc.", "", "5"],
  ["What is the use case of the system?", "Internal use only", "Systems that are only used by internal personnel and have access restrictions in place", "", "1"],
  ["What is the use case of the system?", "Internal and Partner use only", "Systems that have restrcited access to only internal personnel and/or partners", "", "2"],
  ["What is the use case of the system?", "Deprecated customer use", "Authenticated systems that are no longer under active development or are being retired", "", "4"],
  ["What is the use case of the system?", "Deprecated public use", "Publicly available systems that are no longer under active development or are being retired", "", "4"],
  ["What is the use case of the system?", "Standard customer use", "Authenticated systems that are in general availability", "", "3"],
  ["What is the use case of the system?", "Standard public use", "Publicly available systems that are in general availability", "", "3"],
  ["What is the use case of the system?", "Strategic customer use", "Systems that are highly publicised or marketeted with higher impact potential", "", "5"],
  ["What is the use case of the system?", "Strategic public use", "Publicly available systems that are highly publicised or marketeted with higher impact potential", "", "5"]
];

var sampleInherent = {
  "RowId" : "1893cdb0-7347-4abb-b2f9-c691b2d4cd6e",
  "Release" : "1.0",
  "Timestamp" : "2020-03-05 19:16:20",
  "System Identifier" : "foo",
  "Who has access to the system?" : "Restricted Internal Only",
  "How exposed is the knowledge of the system?" : "Partner",
  "What types of information could be disclosed?" : "Authentication secrets",
  "What is the use case of the system?" : "Standard customer use"
};

var sampleResidual = {
  "RowId" : "c89fca69-2c15-44a2-94c5-2bd7435fce0f",
  "Release" : "1.0",
  "Timestamp" : "2020-03-05 19:16:20",
  "InherentId" : "1893cdb0-7347-4abb-b2f9-c691b2d4cd6e",
  "System Identifier" : "foo",
  "Technical debt" : "Moderate",
  "Security Coverage" : "Some"
}

function populateSampleData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName(likelihoodSheet).getRange(1,1,sampleLikelihoodFactors.length, sampleLikelihoodFactors[0].length).setValues(sampleLikelihoodFactors);
  ss.getSheetByName(impactSheet).getRange(1,1,sampleImpactFactors.length, sampleImpactFactors[0].length).setValues(sampleImpactFactors);
  ss.getSheetByName(residualRiskFactorsSheet).getRange(1,1,sampleResidualFactors.length, sampleResidualFactors[0].length).setValues(sampleResidualFactors);
  for(var i = 0; i < sampleLikelihoodFactors.length; i++) {
    ss.getSheetByName(likelihoodSheet).getRange(i + 1, 4, 1, 1).setFormula('=CONCATENATE(A' + (i + 1) + ',"~",B' + (i + 1) + ')');
  }
  for(var i = 0; i < sampleImpactFactors.length; i++) {
    ss.getSheetByName(impactSheet).getRange(i + 1, 4, 1, 1).setFormula('=CONCATENATE(A' + (i + 1) + ',"~",B' + (i + 1) + ')');
  }
  for(var i = 0; i < sampleResidualFactors.length; i++) {
    ss.getSheetByName(residualRiskFactorsSheet).getRange(i + 1, 4, 1, 1).setFormula('=CONCATENATE(A' + (i + 1) + ',"~",B' + (i + 1) + ')');
  }
  var icol = 5;
  var ik = Object.keys(sampleInherent);
  for(i = 4; i < ik.length; i++) {
    ss.getSheetByName(inherentDataSheet).getRange(1, icol, 1, 1).setValue([[ik[i]]]);
    icol++;
  }
  var rcol = 6;
  var rk = Object.keys(sampleResidual);
  for(i = 5; i < rk.length; i++) {
    ss.getSheetByName(residualDataSheet).getRange(1, rcol, 1, 1).setValue([[rk[i]]]);  
    rcol++;
  }
  // populate inherent and residual data cells
  saveInherentData(sampleInherent);
  saveResidualData(sampleResidual);
}
