// Script (指令碼): formatter.gs

/* global SpreadsheetApp, SettingsManager */

class Formatter {
  constructor() {
    this.opList = [
      '=',
      '!=',
      '>',
      '>=',
      '<',
      '<=',
      '介於 [x,y]',
      '不介於 [x,y]',
      '文字包含',
      '文字不包含',
      '文字開頭',
      '文字結尾',
      'regex',
    ];
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
        if (!Formatter.matchRule(sheetValue, ruleIndex, rule)) {
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
        if (Formatter.matchRule(sheetValue, ruleIndex, rule)) {
          isAnyMatch = true;
          break;
        }
      }

      output.push(isAnyMatch);
    }
    return output;
  }

  static matchRule(sheetValue, ruleIndex, rule) {
    const sheetStr = `${sheetValue}`;
    // Reference: https://support.google.com/docs/table/25273
    switch (rule.op) {
      case '=': {
        const sheetNum = Formatter.parseStrToNum(sheetStr);
        const ruleNum = Formatter.parseStrToNum(rule.value);
        if (Number.isNaN(sheetNum) || Number.isNaN(ruleNum)) {
          return sheetStr === rule.value;
        }
        return Math.abs(sheetNum - ruleNum) < Number.EPSILON;
      }
      case '!=': {
        const sheetNum = Formatter.parseStrToNum(sheetStr);
        const ruleNum = Formatter.parseStrToNum(rule.value);
        if (Number.isNaN(sheetNum) || Number.isNaN(ruleNum)) {
          return sheetStr !== rule.value;
        }
        return Math.abs(sheetNum - ruleNum) >= Number.EPSILON;
      }
      case '>': {
        const sheetNum = Formatter.parseStrToNum(sheetStr);
        const ruleNum = Formatter.parseStrToNum(rule.value);
        if (Number.isNaN(sheetNum) || Number.isNaN(ruleNum)) {
          return sheetStr > rule.value;
        }
        return sheetNum > ruleNum;
      }
      case '>=': {
        const sheetNum = Formatter.parseStrToNum(sheetStr);
        const ruleNum = Formatter.parseStrToNum(rule.value);
        if (Number.isNaN(sheetNum) || Number.isNaN(ruleNum)) {
          return sheetStr >= rule.value;
        }
        return sheetNum >= ruleNum;
      }
      case '<': {
        const sheetNum = Formatter.parseStrToNum(sheetStr);
        const ruleNum = Formatter.parseStrToNum(rule.value);
        if (Number.isNaN(sheetNum) || Number.isNaN(ruleNum)) {
          return sheetStr < rule.value;
        }
        return sheetNum < ruleNum;
      }
      case '<=': {
        const sheetNum = Formatter.parseStrToNum(sheetStr);
        const ruleNum = Formatter.parseStrToNum(rule.value);
        if (Number.isNaN(sheetNum) || Number.isNaN(ruleNum)) {
          return sheetStr <= rule.value;
        }
        return sheetNum <= ruleNum;
      }
      case '介於 [x,y]': {
        const sheetNum = Formatter.parseStrToNum(sheetStr);
        if (Number.isNaN(sheetNum)) return false;
        const [start, end] = Formatter.tryParseRangeText(ruleIndex, rule);
        if (
          Number.isNaN(sheetNum)
          || Number.isNaN(start)
          || Number.isNaN(end)
        ) {
          return false;
        }
        return sheetNum >= start && sheetNum <= end;
      }
      case '不介於 [x,y]': {
        const sheetNum = Formatter.parseStrToNum(sheetStr);
        const [start, end] = Formatter.tryParseRangeText(ruleIndex, rule);
        if (
          Number.isNaN(sheetNum)
          || Number.isNaN(start)
          || Number.isNaN(end)
        ) {
          return false;
        }
        return !(sheetNum >= start && sheetNum <= end);
      }
      case '文字包含': {
        return sheetStr.indexOf(rule.value) >= 0;
      }
      case '文字不包含': {
        return sheetStr.indexOf(rule.value) < 0;
      }
      case '文字開頭': {
        return sheetStr.startsWith(rule.value);
      }
      case '文字結尾': {
        return sheetStr.endsWith(rule.value);
      }
      case 'regex': {
        const regex = new RegExp(rule.value);
        return regex.test(sheetStr);
      }
      default:
        throw new Error(`Unknown operation ${rule.op}`);
    }
  }

  static tryParseRangeText(ruleIndex, rule) {
    const str = rule.value;
    let array = [];
    try {
      array = JSON.parse(str);
    } catch (err) {
      const tokens = str.split(',');
      array = tokens.map((token) => Formatter.parseStrToNum(token));
    }
    if (array.length !== 2) {
      throw new Error(
        `無法判別第${ruleIndex + 1}個條件: "${str}"\
請輸入兩個數字，範例: "100,200" (不包含雙引號)`,
      );
    }

    return array;
  }

  static parseStrToNum(str) {
    return Number.parseFloat(str);
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
