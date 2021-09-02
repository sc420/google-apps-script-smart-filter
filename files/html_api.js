// Script: html_api.gs

/* global formatter, getHeaderNamesInActiveSheet, HtmlService, SpreadsheetApp */
/* global showUpdateSuccessToast */

//------------------------------------------------------------------------------
// Used in HTML
//------------------------------------------------------------------------------

// Reference: https://developers.google.com/apps-script/guides/html/best-practices#separate_html_css_and_javascript
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

//------------------------------------------------------------------------------
// Called by Scripts
//------------------------------------------------------------------------------

function getDataForHTML() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  const sheetSettings = formatter.getSheetSettings(sheetName);

  const headerNames = getHeaderNamesInActiveSheet();
  const { opList } = formatter;

  return { sheetSettings, headerNames, opList };
}

function updateSheetSettings(newSheetSettings) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  const info = formatter.updateSheetSettings(sheetName, newSheetSettings);
  showUpdateSuccessToast(info);
  return info;
}

function resetSheetSettings() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('重設', '重設此工作表的設定?', ui.ButtonSet.YES_NO);
  if (result !== ui.Button.YES) return false;

  ui.alert(
    '提醒',
    '重設後將會關閉設定側邊欄位，如需繼續使用請重新開啟',
    ui.ButtonSet.OK,
  );

  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  formatter.resetSheetSettings(sheetName);
  return true;
}
