// Script (指令碼): main.gs

/* global Formatter, HtmlService, SettingsManager, SpreadsheetApp */

//------------------------------------------------------------------------------
// Global Variables
//------------------------------------------------------------------------------

const formatter = new Formatter();

//------------------------------------------------------------------------------
// Used Functions
//------------------------------------------------------------------------------

function showUpdateSuccessToast(info) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.toast(
    `包含數量: ${info.numInclude}, 排除數量: ${info.numExclude},\
衝突數量: ${info.numConflict}`,
    'Smart 篩選',
  );
}

function updateActiveSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  const info = formatter.updateSheet(sheetName);
  showUpdateSuccessToast(info);
}

function getHeaderNamesInActiveSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const headerRange = dataRange.offset(0, 0, 1);
  const values = headerRange.getValues();
  return values[0];
}

function resetAllSettings() {
  const settingsManager = new SettingsManager();
  settingsManager.resetAllSettings();
}

function loadHtml(filename, title) {
  const template = HtmlService.createTemplateFromFile(filename);
  const html = template.evaluate();
  html.setTitle(title);
  return html;
}

//------------------------------------------------------------------------------
// Menu
//------------------------------------------------------------------------------

function showSettingsSidebar() {
  const ui = SpreadsheetApp.getUi();
  const html = loadHtml('settings_sidebar', '設定');
  ui.showSidebar(html);
}

function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  // Reference: https://developers.google.com/apps-script/guides/dialogs#custom_sidebars
  ui.createMenu('Smart 篩選').addItem('設定', 'showSettingsSidebar').addToUi();
}

//------------------------------------------------------------------------------
// Triggers
//------------------------------------------------------------------------------

function onEdit() {
  updateActiveSheet();
}

function onOpen() {
  updateActiveSheet();
  createCustomMenu();
}
