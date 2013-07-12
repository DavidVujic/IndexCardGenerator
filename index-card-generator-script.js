// START: Template functions
// You may need to update these functions if the template is changed.

function getTemplateArea() {
  return "A1:F10";
}

function setCardId(backlogItem, card) {
 card.getCell(2, 3).setValue(backlogItem['Id']);
}

function setCardName(backlogItem, card) {
  var max = 19;
  var storyName = backlogItem['Name'];
  
  if (storyName && storyName.length > max) {
    storyName = storyName.substring(0, max) + '...';
  }
  
  card.getCell(3, 3).setValue(storyName);
}

function setUserStory(backlogItem, card) {
 card.getCell(5, 3).setValue(backlogItem['User story']);
}

function setHowToTest(backlogItem, card) {
 card.getCell(8, 3).setValue(backlogItem['How to test']);
}

function setImportance(backlogItem, card) {
 card.getCell(5, 5).setValue(backlogItem['Importance']);
}

function setEstimate(backlogItem, card) {
 card.getCell(8, 5).setValue(backlogItem['Estimate']);
}

function getTemplateStartColumn() {
 return getTemplateArea().substring(0,1);
}

function getTemplateStartRow() {
  return parseInt(getTemplateArea().substring(1,2), 10);
}

function getTemplateLastColumn() {
  return getTemplateArea().substring(3,4);
}

function getTemplateLastRow() {
  return parseInt(getTemplateArea().substring(4), 10);
}
// END: Template functions

// START: Get sheets
function getSpreadsheet() {
 return SpreadsheetApp.getActiveSpreadsheet(); 
}

function getBacklogSheet() {
  return getSpreadsheet().getSheetByName("Backlog");
}

function getTemplateSheet() {
  return getSpreadsheet().getSheetByName("Template");
}

function getCardSheet() {
  return getSpreadsheet().getSheetByName("Cards");
}

function getPreparedCardSheet(template, numberOfItems, numberOfRows) {
  var rowsNeeded = numberOfItems * numberOfRows;
  
  var sheet = getCardSheet();
  sheet.clear();
  
  setColumnWidthTo(sheet, template);
  
  var rows = sheet.getMaxRows();
  
  if (rows < rowsNeeded) {
    sheet.insertRows(1, (rowsNeeded - rows));
  }
  
  setRowHeightTo(sheet, numberOfRows, numberOfItems);
  
  return sheet;
}
// END: Get sheets

// START: Get range within sheets
function getTemplateRange() {
 return getTemplateSheet().getRange(getTemplateArea());
}

function getHeadersRange(backlog) {
  return backlog.getRange(1, 1, 1, backlog.getLastColumn());
}

function getItemsRange(backlog) {
  var numRows = backlog.getLastRow() - 1;
  
  return backlog.getRange(2, 1, numRows, backlog.getLastColumn());
}

function getSelectedItemsRange(backlog) {
  var range = getSpreadsheet().getActiveRange();
  var startRow = range.getRowIndex();
  var rows = range.getNumRows();
  
  if (startRow < 2 ) { 
    startRow = 2; 
    rows = (rows > 1 ? rows-1 : rows);
  }
  
  return backlog.getRange(startRow, 1, rows, backlog.getLastColumn());
}
// END: Get range within sheets

function setRowHeightTo(cardSheet, numberOfRows, numberOfItems) {
  var templateSheet = getTemplateSheet();
  
  for (var i = 0; i < numberOfItems; i++) {
    for (var j = 1; j < (numberOfRows+1); j++) {
      var currentRow = (i*numberOfRows)+j;
      var currentHeight = templateSheet.getRowHeight(j);
      cardSheet.setRowHeight(currentRow, currentHeight);
    }
  }
}

function setColumnWidthTo(cardSheet, templateRange) {
  var templateSheet = getTemplateSheet();
  var max = templateRange.getLastColumn() + 1;
  
  for (var i = 1; i < max; i++) {
    var currentWidth = templateSheet.getColumnWidth(i);
    cardSheet.setColumnWidth(i, currentWidth);
  }
}

/* Get backlog items as objects with property name and values from the backlog. */
function getBacklogItems(selectedOnly) {
  var backlog = getBacklogSheet();
  
  var rowsRange = (selectedOnly ? getSelectedItemsRange(backlog) : getItemsRange(backlog));
  var rows = rowsRange.getValues();
  var headers = getHeadersRange(backlog).getValues()[0];
  
  var backlogItems = [];
  
  for (var i = 0; i < rows.length; i++) {
    var backlogItem = {};
    
    for (var j = 0; j < rows[i].length; j++) {
      backlogItem[headers[j]] = rows[i][j];
    }
    
    backlogItems.push(backlogItem);
  }
  
  return backlogItems;
}

function assertCardSheetExists() {
  if (getCardSheet() == null) {
    getSpreadsheet().insertSheet("Cards", 2);
    Browser.msgBox("The 'Cards' sheet was missing and has now been added. Please try again.");
    return false;
  }
  
  return true;
}

function createCardsFromBacklog() {
  if (!assertCardSheetExists()) {
    return;
  }
  
  var backlogItems = getBacklogItems(false);
  createCards(backlogItems);
}

function createCardsFromSelectedRowsInBacklog() {
  if (!assertCardSheetExists()) {
    return;
  }
  
  if (getBacklogSheet().getName() != SpreadsheetApp.getActiveSheet().getName()) {
    Browser.msgBox("The Backlog sheet need to be active when creating cards from selected rows. Please try again.");
    return;
  }
  
  var backlogItems = getBacklogItems(true);
  createCards(backlogItems);
}

function createCards(backlogItems) {
  var numberOfRows = getTemplateLastRow();
  var template = getTemplateRange();
  var cardSheet = getPreparedCardSheet(template, backlogItems.length, numberOfRows);
  
  var startRow = getTemplateStartRow();  
  var lastRow = getTemplateLastRow();
  var startColumn = getTemplateStartColumn();
  var lastColumn = getTemplateLastColumn();
   
  for (var i = 0; i < backlogItems.length; i++) {
    var rangeVal = startColumn + startRow + ":" + lastColumn + lastRow;
    
    var card = cardSheet.getRange(rangeVal);
    template.copyTo(card);
    
    setCardId(backlogItems[i], card);
    setCardName(backlogItems[i], card);
    setUserStory(backlogItems[i], card);
    setImportance(backlogItems[i], card);
    setHowToTest(backlogItems[i], card);
    setEstimate(backlogItems[i], card);
    
    startRow += numberOfRows;
    lastRow += numberOfRows;
  }
  
  Browser.msgBox("Done!");
}

/* Will add a Cards menu. Runs when the spreadsheet is loaded. */
function onOpen() {
  var sheet = getSpreadsheet();
  var menuEntries = [ {name: "Create cards", functionName: "createCardsFromBacklog"}, {name: "Create cards from selected rows", functionName: "createCardsFromSelectedRowsInBacklog"} ];
   
  sheet.addMenu("Story Cards", menuEntries);
}