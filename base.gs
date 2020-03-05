/*
 * Copyright (c) 2020, salesforce.com, inc.
 * All rights reserved.
 * SPDX-License-Identifier: BSD-3-Clause
 * For full license text, see the LICENSE file in the repo root or https://opensource.org/licenses/BSD-3-Clause
 */

var _CURRENT_VERSION_;
var inherentDataSheet = "Raw Inherent Data";

function onOpen() {
  SpreadsheetApp.getUi() 
      .createMenu('Risk Management')
      .addItem('New Inherent Risk System', 'newSystem')
      .addItem("New Residual Risk System", 'showNewResidualForm')
      .addSeparator()
      .addItem('Recalculate Inherent Risk', 'populateInherentRisk')
      .addSeparator()
      .addItem('Recalculate Residual Risk', 'populateResidualRisk')
      .addToUi();
}

function newSystem() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Create new risk entry?", ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
      showInherentRiskDialog(getNewSystem());
  }
}

function loadSystemFromRawSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentDataSheet);
  var cell = sheet.getCurrentCell();
  if(null != cell) {
    var rowId = sheet.getRange(cell.getRow(), 1).getValue();
    var sysName = sheet.getRange(cell.getRow(), 4).getValue();
    var release = sheet.getRange(cell.getRow(), 2).getValue();
    if(!(rowId === "RowId" || rowId === '')) { 
      promptToLoad(rowId, sysName, release);
    }
  }
}

function loadSystemFromInherentSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentRiskSheet);
  var cell = sheet.getCurrentCell();
  if(null != cell) {
    var rowId = sheet.getRange(cell.getRow(), 1).getValue();
    var sysName = sheet.getRange(cell.getRow(), 2).getValue();
    var release = sheet.getRange(cell.getRow(), 3).getValue();
    if(!(rowId === "RowId" || rowId === '')) { 
      promptToLoad(rowId, sysName, release);
    }
  }
}

function newReleaseFromInherentSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(inherentRiskSheet);
  var cell = sheet.getCurrentCell();
  if(null != cell) {
    var rowId = sheet.getRange(cell.getRow(), 1).getValue();
    var sysName = sheet.getRange(cell.getRow(), 2).getValue();
    var release = sheet.getRange(cell.getRow(), 3).getValue();
    if(!(rowId === "RowId" || rowId === '')) { 
      promptForNewRelease(rowId, sysName, release);
    }
  }
}

function promptToLoad(rowId, sysName, release) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert("Load options for " + sysName + ' / ' + release + "?", ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      Logger.log('Loading data for %s.', rowId);
      workOnSystem(rowId);
    } else if (response == ui.Button.NO) {
      Logger.log('The user didn\'t want to load the system %s.', rowId);
    } else {
      Logger.log('The user clicked the close button in the dialog\'s title bar.');
    }
}

function promptForNewRelease(rowId, sysName, release) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert("Create a new release for " + sysName + ' / ' + release + "?", ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      Logger.log('Creating new release for %s.', rowId);
      showInherentRiskDialog(getNewReleaseByRowId(rowId));
    } else if (response == ui.Button.NO) {
      Logger.log('The user didn\'t want to create a new release %s.', rowId);
    } else {
      Logger.log('The user clicked the close button in the dialog\'s title bar.');
    }
}

function promptToLoadResidual(rowId, sysName, release) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert("Load residual options for " + sysName + ' / ' + release + "?", ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      Logger.log('Loading data for %s.', rowId);
      workOnResidual(rowId);
    } else if (response == ui.Button.NO) {
      Logger.log('The user didn\'t want to load the system %s.', rowId);
    } else {
      Logger.log('The user clicked the close button in the dialog\'s title bar.');
    }
}

function showInherentRiskDialog(systemData) {
  var page = HtmlService.createTemplateFromFile('inherentPage');
  page.data = systemData;
  var html = page.evaluate().setWidth(800).setHeight(600).setTitle(systemData['System Identifier'] ? systemData['System Identifier'] : ""  + ' Inherent Risk');
  Logger.log(page.getCodeWithComments());
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showResidualRiskDialog(systemData) {
  var page = HtmlService.createTemplateFromFile('residualPage');
  page.data = systemData;
  var html = page.evaluate().setWidth(800).setHeight(600).setTitle(systemData['System Identifier'] ? systemData['System Identifier'] : ""  + ' Residual Risk');
  Logger.log(page.getCodeWithComments());
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function workOnSystem(rowId) {
  showInherentRiskDialog(getInherentSystemByRowId(rowId));
}

function workOnResidual(rowId) {
  showResidualRiskDialog(getResidualRowId(rowId));
}

function showNewResidualForm() {
  var page = HtmlService.createTemplateFromFile('newResidualPage');
  page.data = inherentWithNoResidual();
  var html = page.evaluate().setWidth(400).setHeight(200);
  SpreadsheetApp.getUi()
    .showModalDialog(html, 'New Residual Risk');
}

function processNewResidualForm(formObject) {
  var name = formObject["systemName"];
  var inhSystem = getLatestSystemByName(name);
  var resSystem = newResidualSystem();
  resSystem["System Identifier"] = name;
  resSystem["InherentId"] = inhSystem["RowId"];
  showResidualRiskDialog(resSystem);
}

function promptForResidualNewRelease(rowId, sysName, release) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert("Create a new residual release for " + sysName + ' / ' + release + "?", ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      showResidualRiskDialog(getNewResidualReleaseByRowId(rowId));
    } else if (response == ui.Button.NO) {
      Logger.log('The user didn\'t want to create a new release %s.', rowId);
    } else {
      Logger.log('The user clicked the close button in the dialog\'s title bar.');
    }
}
