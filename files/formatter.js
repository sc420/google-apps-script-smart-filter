// Script: formatter.gs

/* global SpreadsheetApp, SettingsManager */

class Formatter {
  constructor() {
    this.opList = ['>=', '=', '<='];
    this.defaultBackgroundColor = 'white';
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!this.spreadsheet) throw new Error('Should have an active spreadsheet');

    this.settingsManager = new SettingsManager();
  }

  updateSheet(sheetName) {
    const sheetSettings = this.settingsManager.getSheetSettings(sheetName);
    const sheet = this.findSheetByName(sheetName);
    const headerNames = Formatter.getHeaderNames(sheet);
    const dataRange = sheet.getDataRange();

    const [bgColors, info] = Formatter.computeBackgroundColors(
      headerNames,
      dataRange,
      sheetSettings,
    );
    Formatter.updateSheetBackgroundColors(bgColors, dataRange);
    return info;
  }

  getSheetSettings(sheetName) {
    return this.settingsManager.getSheetSettings(sheetName);
  }

  updateSheetSettings(sheetName, newSettings) {
    this.settingsManager.setSheetSettings(sheetName, newSettings);
    return this.updateSheet(sheetName);
  }

  resetSheetSettings(sheetName) {
    this.settingsManager.resetSheetSettings(sheetName);
    this.updateSheet(sheetName);
  }

  static computeBackgroundColors(headerNames, dataRange, sheetSettings) {
    const { includeFilter, excludeFilter } = sheetSettings;
    const includeRows = Formatter.checkIncludeRows(
      headerNames,
      dataRange,
      includeFilter.rules,
    );
    const excludeRows = Formatter.checkExcludeRows(
      headerNames,
      dataRange,
      excludeFilter.rules,
    );
    const bgColors = Array(includeRows.length);
    const info = {
      numInclude: 0,
      numExclude: 0,
      numConflict: 0,
    };
    for (let row = 0; row < includeRows.length; row += 1) {
      if (includeRows[row]) {
        bgColors[row] = includeFilter.formattingStyle.backgroundColor;
      }
    }
    for (let row = 0; row < excludeRows.length; row += 1) {
      if (excludeRows[row]) {
        bgColors[row] = excludeFilter.formattingStyle.backgroundColor;
        info.numExclude += 1;
      }
    }
    for (let row = 0; row < includeRows.length; row += 1) {
      if (includeRows[row] && !excludeRows[row]) {
        info.numInclude += 1;
      }
      if (includeRows[row] && excludeRows[row]) {
        info.numConflict += 1;
      }
    }
    return [bgColors, info];
  }

  static updateSheetBackgroundColors(newBgColors, dataRange) {
    const sheetBgColors = dataRange.getBackgrounds();
    for (let row = 1; row < sheetBgColors.length; row += 1) {
      const newBgColor = newBgColors[row];
      sheetBgColors[row] = sheetBgColors[row].fill(newBgColor);
    }
    dataRange.setBackgrounds(sheetBgColors);
  }

  static checkIncludeRows(headerNames, dataRange, rules) {
    const sheetValues = dataRange.getValues();
    const output = [];
    output.push(false); // First row
    for (let row = 1; row < sheetValues.length; row += 1) {
      const sheetRow = sheetValues[row];

      let numMatches = 0;
      for (let ruleIndex = 0; ruleIndex < rules.length; ruleIndex += 1) {
        const rule = rules[ruleIndex];
        const col = headerNames.indexOf(rule.key);
        if (col < 0) throw new Error(`找不到欄位 "${rule.key}"`);
        const sheetValue = sheetRow[col];
        if (
          !Formatter.matchRule(dataRange, row, col, sheetValue, ruleIndex, rule)
        ) {
          break;
        }
        numMatches += 1;
      }

      const isAnyMatch = numMatches >= 1;
      const isAllMatch = numMatches === rules.length;
      output.push(isAnyMatch && isAllMatch);
    }
    return output;
  }

  static checkExcludeRows(headerNames, dataRange, rules) {
    const sheetValues = dataRange.getValues();
    const output = [];
    output.push(false); // First row
    for (let row = 1; row < sheetValues.length; row += 1) {
      const sheetRow = sheetValues[row];

      let isAnyMatch = false;
      for (let ruleIndex = 0; ruleIndex < rules.length; ruleIndex += 1) {
        const rule = rules[ruleIndex];
        const col = headerNames.indexOf(rule.key);
        if (col < 0) throw new Error(`找不到欄位 "${rule.key}"`);
        const sheetValue = sheetRow[col];
        if (
          Formatter.matchRule(dataRange, row, col, sheetValue, ruleIndex, rule)
        ) {
          isAnyMatch = true;
          break;
        }
      }

      output.push(isAnyMatch);
    }
    return output;
  }

  static matchRule(dataRange, row, col, sheetValue, ruleIndex, rule) {
    switch (rule.op) {
      case '>=': {
        const sheetValueAsFloat = Formatter.parseSheetValueAsFloat(
          dataRange,
          row,
          col,
          sheetValue,
        );
        const ruleValueAsFloat = Formatter.parseRuleValueAsFloat(
          ruleIndex,
          rule,
        );
        return sheetValueAsFloat >= ruleValueAsFloat;
      }
      case '=': {
        const sheetValueAsFloat = Formatter.parseSheetValueAsFloat(
          dataRange,
          row,
          col,
          sheetValue,
        );
        const ruleValueAsFloat = Formatter.parseRuleValueAsFloat(
          ruleIndex,
          rule,
        );
        return Math.abs(sheetValueAsFloat - ruleValueAsFloat) < Number.EPSILON;
      }
      case '<=': {
        const sheetValueAsFloat = Formatter.parseSheetValueAsFloat(
          dataRange,
          row,
          col,
          sheetValue,
        );
        const ruleValueAsFloat = Formatter.parseRuleValueAsFloat(
          ruleIndex,
          rule,
        );
        return sheetValueAsFloat <= ruleValueAsFloat;
      }
      default:
        throw new Error(`Unknown operation ${rule.op}`);
    }
  }

  static parseSheetValueAsFloat(dataRange, row, col, sheetValue) {
    const sheetValueAsFloat = Number.parseFloat(sheetValue);
    if (Number.isNaN(sheetValueAsFloat)) {
      const cellRange = dataRange.offset(row, col, 1, 1);
      const a1Notation = cellRange.getA1Notation();
      throw new Error(`無法將 "${a1Notation}" 轉成數字: "${sheetValue}"`);
    }
    return sheetValueAsFloat;
  }

  static parseRuleValueAsFloat(ruleIndex, rule) {
    const ruleValueAsFloat = Number.parseFloat(rule.value);
    if (Number.isNaN(ruleValueAsFloat)) {
      throw new Error(
        `無法將第${ruleIndex + 1}個條件轉成數字: "${rule.value}"`,
      );
    }
    return ruleValueAsFloat;
  }

  findSheetByName(sheetName) {
    const sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(
        `Could not find sheet "${sheetName}" in the active spreadsheet`,
      );
    }
    return sheet;
  }

  static getHeaderNames(sheet) {
    const dataRange = sheet.getDataRange();
    const headerRange = dataRange.offset(0, 0, 1);
    const values = headerRange.getValues();
    return values[0];
  }
}
