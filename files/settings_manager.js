// Script (指令碼): settings_manager.gs

/* global PropertiesService */

class SettingsManager {
  constructor() {
    this.propertyKey = 'IPO_BOT_SETTINGS';
    this.documentProperties = PropertiesService.getDocumentProperties();

    this.settings = this.readSettings();
  }

  getSheetSettings(sheetName) {
    if (!(sheetName in this.settings)) {
      this.settings[sheetName] = SettingsManager.getDefaultSettingsPerSheet();
    }
    return this.settings[sheetName];
  }

  setSheetSettings(sheetName, newSettings) {
    this.settings[sheetName] = newSettings;
    this.saveSettings();
  }

  resetSheetSettings(sheetName) {
    this.settings[sheetName] = SettingsManager.getDefaultSettingsPerSheet();
    this.saveSettings();
  }

  resetAllSettings() {
    this.settings = {};
    this.saveSettings();
  }

  static getDefaultSettingsPerSheet() {
    return {
      includeFilter: {
        rules: [],
        formattingStyle: {
          backgroundColor: '#b7e1cd',
        },
      },
      excludeFilter: {
        rules: [],
        formattingStyle: {
          backgroundColor: '#f4c7c3',
        },
      },
    };
  }

  readSettings() {
    let dataAsStr = this.documentProperties.getProperty(this.propertyKey);
    if (!dataAsStr) dataAsStr = '{}';
    try {
      return JSON.parse(dataAsStr);
    } catch (e) {
      return {};
    }
  }

  saveSettings() {
    const dataAsStr = JSON.stringify(this.settings);
    this.documentProperties.setProperty(this.propertyKey, dataAsStr);
  }
}
