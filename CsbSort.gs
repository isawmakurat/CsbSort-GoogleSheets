function onOpen() {
//  var submenu = [{name:"Sortuj pò kaszëbskù wedle wëbróny kòlumnë", functionName:"sortCsbByColumn"}];
//  SpreadsheetApp.getActiveSpreadsheet().addMenu('CsbSort', submenu);  
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem("Sortuj pò kaszëbskù...", "csbSortInit")
  .addToUi();
};

function csbSortInit() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isSelectedManyRows()) {
    ui.alert("Wëbierzë nôprzód zakres do pòsortowaniô!");
    return;
  }
  var sortParametersForm = HtmlService.createTemplateFromFile("CsbSortParametersForm").evaluate().setWidth(300).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(sortParametersForm, "Sortowanié pò kaszëbskù")
};

function csbSortByColumn(sortColumn) {
  var ui = SpreadsheetApp.getUi();

  var sortColumnIdx = 1 + sortColumn.toUpperCase().charCodeAt(0) - "A".charCodeAt(0);
  var data = getSelectionRange();
  if((sortColumnIdx < data.getColumn()) || (sortColumnIdx >= (data.getColumn() + data.getWidth()))) {
    ui.alert("Pòdónô kòlumna nie je w zaznaczonym zakresu. Spóbùjë jesz rôz.");
    return;
  }
  data = prepareTmpColumn(data, sortColumnIdx);
  data.sort(sortColumnIdx);
  data = removeTmpColumn(data, sortColumnIdx);
};

function prepareTmpColumn(cellRange, sortColumnIdx) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.insertColumnBefore(sortColumnIdx);
//  sheet.hideColumns(sortColumnIdx);
  var originalSortRange = sheet.getRange(cellRange.getRow(), sortColumnIdx + 1, cellRange.getLastRow());
  var tmpColumnRange = sheet.getRange(cellRange.getRow(), sortColumnIdx, cellRange.getLastRow())
  
  tmpColumnRange.setValues(translateAll(originalSortRange.getValues()));
  return cellRange.offset(0, 0, cellRange.getHeight(), cellRange.getWidth() + 1).activate();
};

function translateAll(valuesToTranslate) {
  return valuesToTranslate.map(translateOne);
};

function translateOne(valueToTranslate) {
  var letters = (valueToTranslate[0]+"").toUpperCase().split("");
  var translated = letters.map(translateLetter).join("");
  return [translated];
};

function translateLetter(letterToTranslate) {
  const CSB_ALPHABET = "0123456789AĄÃBCĆDEĘÉËFGHIJKLŁMNŃOÒÓÔPQRSŚTUÙVWXYZŹŻ -,;:!?.()/%";
  const TMP_ALPHABET = "0123456789_,;:!?.()[]{}@*/\&#%^<>|~$ABCDEFGHIJKLMNOPQRSTUVWXYZ ";

  var charIdx = CSB_ALPHABET.indexOf(letterToTranslate);
  if (charIdx >= 0) {
    return TMP_ALPHABET.charAt(charIdx);
  } else {
    return ""
  }
};

function removeTmpColumn(cellRange, tmpColumnIdx) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.deleteColumn(tmpColumnIdx);
  return cellRange.offset(0, 0, cellRange.getHeight(), cellRange.getWidth() - 1).activate();
};

function getSelectionRange() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getSelection().getActiveRange();
  return data;
};

function getSelectionRangeDescription() {
  return getSelectionRange().getA1Notation();
};

function isSelectedManyRows() {
  return getSelectionRange().getHeight() > 1;
};

function getSelectionColumns() {
  var selectedRange = getSelectionRange();
  var columns = [];
  var activeCoursorColumn = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getCurrentCell().getColumn();
  for(var i = selectedRange.getColumn(); i <= selectedRange.getLastColumn(); i++) {
    var rangeColumn = {"column": String.fromCharCode("A".charCodeAt(0) + i - 1)};
    if(i == activeCoursorColumn) {
      rangeColumn.selected = true;
    }
    columns.push(rangeColumn);
  }
  return columns;
};
