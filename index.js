/**
 * Добавляет пункт меню в таблицу
 */
function onOpen() {
    var spreadsheet = SpreadsheetApp.getActive();
    var menuItems = [
      {name: 'Добавить новые переводы', functionName: 'addTranslations'}
    ];
    spreadsheet.addMenu('BuzzGuru', menuItems);
  }
  
  var TRANSLATE_ID_NAME = '_concat_id';
  var HEAD_RANGE = 'A1:P2';
  var TABLE_RANGE = 'A1:J1958';
  var BASE_TABLE_NAME = 'Интерфейс 1.0';
  // var UPDATES_TABLE_NAME = 'Первая итерация до 1900 строки';
  var UPDATES_TABLE_NAME = 'Вычитка английского';
  var ROW_LENGTH = 15;
  var COLOR_SUCCESS = "lightgreen";
  var COLOR_ALERT = "red";
  
  /** 
   * Запускает заполнение пустых значений в основной таблице из текущей таблицы с обновлениями
   */
  
  function addTranslations() {
    var spreadsheet = SpreadsheetApp.getActive();
    
    
    // UPDATES TABLE
    
    // Ask table name with updates  
    // var updateTableName = UPDATES_TABLE_NAME;
    var updateTableName = askData('Укажите таблицу переводов', 'Укажите название таблицы с новыми переводами' + ' (например, "' + UPDATES_TABLE_NAME + '"):');
    if (!updateTableName || updateTableName === 'cancel') {
      return;
    }
    
    var updatesTable = spreadsheet.getSheetByName(updateTableName);
    if (!updatesTable) throw new Error('addTranslations: updatesTable ' + updateTableName + ' not found');
    
    // var updatesDataRange = TABLE_RANGE;
    // var updatesDataRange = askData('Укажите диапазон', 'Укажите диапазон данных в таблице новых переводов' + ' (например, "A1:J1958"):');
    
    
  
    // BASE TABLE
    var baseTable = spreadsheet.getSheetByName(BASE_TABLE_NAME);
    if (!baseTable) throw new Error('addTranslations: baseTable ' + BASE_TABLE_NAME + ' not found');
  
    baseTable.activate();
    
    // translateIdResults(baseTable, translateId);
    
    // var startRowIdx = 2;
    var startRowIdx = askData('Укажите начальную строку перевода', 'Укажите номер первой строки в таблице новых переводов ' + updateTableName + ' (например, "2"):');
    if (!startRowIdx || startRowIdx === 'cancel') {
      return;
    }
    
    // var endRowIdx = 3;
    var endRowIdx = askData('Укажите конечную строку перевода', 'Укажите номер последней строки в таблице новых переводов ' + updateTableName + ' (например, "100"):');
    if (!endRowIdx || endRowIdx === 'cancel') {
      return;
    }
    
    Logger.log('Базовая таблица "' + BASE_TABLE_NAME + '". Таблица с обновлениями "' + updateTableName + '". Первая строка: ' + startRowIdx + '. Последняя строка: ' + endRowIdx);  
    
    var statFoundLinesIdx = [];
    var statNotFoundLineIds = [];
    
    var updatesTranslateIdIdx = findColIndexByColName(updatesTable, TRANSLATE_ID_NAME);
      
    var baseRuIdx = findColIndexByColName(baseTable, 'ru');
    if (!baseRuIdx) throw new Error('addTranslations: baseRuIdx is not set');
    
    var updatesRuIdx = findColIndexByColName(baseTable, 'ru');
    if (!updatesRuIdx) throw new Error('addTranslations: updatesRuIdx is not set');
    
    updatesTable.activate();
    
    for(var idx = startRowIdx; idx <= endRowIdx; idx++) {
      var replacedIdx = replaceRow(baseTable, updatesTable, idx, baseRuIdx, updateTableName);
      
      if (!replacedIdx) {
        statNotFoundLineIds.push(idx);
        continue;
      }
  
      statFoundLinesIdx.push(replacedIdx + 1); 
    }
    
    var results = 'Обновлены строки (' + statFoundLinesIdx.length + '): ' + statFoundLinesIdx.join(',') + ' | ' +
      'Не найдены ключи (' + statNotFoundLineIds.length + '): ' + statNotFoundLineIds.join(',');
    var title = 'Результы обновления переводов в "' + BASE_TABLE_NAME + '"';
    
    Logger.log(title + ': ' + results);
    
    Browser.msgBox(title, results, Browser.Buttons.OK);
  }
  
  function replaceRow(baseTable, updatesTable, rowIdx, baseRuIdx, updateTableName) {
    var context = 'replaceRow: '; 
    if (!baseTable) throw new Error(context + 'baseTable is not set');
    if (!updatesTable) throw new Error(context + 'updatesTable is not set');
    if (!rowIdx) throw new Error(context + 'rowIdx is not set');
    if (!baseRuIdx) throw new Error(context + 'baseRuIdx is not set');
    if (!updateTableName) throw new Error(context + 'updateTableName is not set');
    
    // Get translate id 
    var translateId = findTranslateId(updatesTable, TABLE_RANGE, rowIdx);
    if (!translateId) {
      setTableCellBg(updatesTable, rowIdx, 1);
      return;
    }
    
    // var updateColName = 'de';
    // var updateColName = askData('Укажите язык', 'Укажите язык' + ' (например, "de"):');
    
    // var updatesColIdx = findColIndexByColName(updatesTable, updateColName);
    // Logger.log('RU ' + updatesRuId + ' DE ' + updatesDeId);
    
    // Found updates table row values for translateId
    // var foundUpdatesRow = getRowByTranslateId(updatesTable, TABLE_RANGE, translateId).data;
    // Logger.log('foundUpdatesRow ' + foundUpdatesRow);
    
    var updatesRow = getRowByTranslateId(updatesTable, TABLE_RANGE, translateId);
    if (!updatesRow) throw new Error(context + 'updatesRow is not found'); 
    
    var foundUpdatesTranslations = updatesRow.translations;
    Logger.log(updateTableName + ' - ' + rowIdx + ': ' + translateId + ' ' + foundUpdatesTranslations);
    
    // var updatedValue = foundUpdatesRow[updatesColIdx];
    // Logger.log('updateColName ' + updateColName + ' updatedValue ' + updatedValue);
    
    // Found base table row with translateId
    var baseRow = getRowByTranslateId(baseTable, TABLE_RANGE, translateId);
    if (!baseRow) {
      setTableCellBg(updatesTable, rowIdx, 1);
      Logger.log('ВНИМАНИЕ! Ключ ' + translateId + ' не найден: ряд ' + rowIdx + ' колонка ' + 1);
      return;
    } else {
      var baseRowIdx = baseRow.index;
      var foundBaseTranslations = baseRow.translations;
      
      var updatesRuIdx = findColIndexByColName(updatesTable, 'ru');
      
      Logger.log(BASE_TABLE_NAME + ' - ' + baseRowIdx + ': ' + translateId + ' ' + foundBaseTranslations);
      
      if (!isSameValue(foundBaseTranslations, foundUpdatesTranslations)) {
        setTableCellBg(updatesTable, rowIdx, updatesRuIdx + 1);
        Logger.log('ВНИМАНИЕ! Значение ключа ' + translateId + ' в обновлениях "' + foundUpdatesTranslations[0] + '" и в базовой таблице "' + foundBaseTranslations[0] + '" не совпадает: ряд ' + rowIdx + ' колонка ' + (updatesRuIdx + 1) );
        return;
      }
      
      setTableCellBg(updatesTable, rowIdx, 1, COLOR_SUCCESS, ROW_LENGTH);
      
      replaceTableRow(baseTable, baseRowIdx + 1, baseRuIdx + 1, foundUpdatesTranslations);
      return baseRowIdx;
    } 
  }
  
  function setTableCellBg(table, rowIdx, colIdx, color, length) {
    var context = 'setTableCellBg: '; 
    if (!table) throw new Error(context + 'table is not set');
    if (!rowIdx) throw new Error(context + 'rowIdx is not set');
    if (!colIdx) throw new Error(context + 'colIdx is not set');
    if (!color) color = COLOR_ALERT;
    if (!length) length = 1;
    
    var range = table.getRange(rowIdx, colIdx, 1, length);
    range.setBackground(color);
  }
  
  function isSameValue(firstArray, secondArray, index) {
    var context = 'isSameValue: '; 
    if (!firstArray) throw new Error(context + 'firstArray is not set');
    if (!secondArray) throw new Error(context + 'secondArray is not set');
    if (!index) index = 0;
    
    return firstArray[index].trim() === secondArray[index].trim();
  }
  
  
  function translateIdResults(table, translateId) {
    var context = 'translateIdResults: '; 
    if (!table) throw new Error(context + 'table is not set');
    if (!translateId) throw new Error(context + 'translateId is not set');
    
    // var dataRange = 'A119:P120';
    // var dataRange = 'A119:P120';
    // var dataRange = askData('Укажите диапазон', 'Укажите диапазон данных' + ' (на пример, "A119:P120"):');
    
    var foundBaseRow = getRowByTranslateId(table, dataRange, translateId).data;
    Logger.log('foundBaseRow ' + foundBaseRow);
    
    var tableHead = getTableHead(table);
    var resultsRows = [tableHead, foundBaseRow];
    
    resultsTable(resultsRows);
    // Browser.msgBox('foundUpdatesRow', foundUpdatesRow, Browser.Buttons.OK);
  }
  
  function replaceTableRow(table, rowIdx, colId, resultsRow) { 
    var context = 'replaceTableRow: '; 
    if (!table) throw new Error(context + 'table is not set');
    if (!rowIdx) throw new Error(context + 'rowIdx is not set');
    if (!colId) throw new Error(context + 'colId is not set');
    if (!resultsRow) throw new Error(context + 'colId is not set');
    
    table.getRange(rowIdx, colId, 1, resultsRow.length).setValues([resultsRow]);
  }
  
  function resultsTable(resultsRows) {
    var context = 'resultsTable: ';  
    if (!resultsRows) throw new Error(context + 'resultsRows is not set');
    
    var spreadsheet = SpreadsheetApp.getActive();
    
    var sheetName = BASE_TABLE_NAME + ' results';
    
    var resultsSheet = spreadsheet.getSheetByName(sheetName);
    if (resultsSheet) {
      resultsSheet.clear();
      resultsSheet.activate();
    } else {
      resultsSheet = spreadsheet.insertSheet(sheetName, spreadsheet.getNumSheets());
    }
    
    resultsSheet.getRange(1, 1, resultsRows.length, resultsRows[0].length).setValues(resultsRows);
  }
  
  function askData(title, msg) {
    var context = 'askData: '; 
    if (!title) throw new Error(context + 'title is not set');
    if (!msg) throw new Error(context + 'msg is not set');
    
    // Promth for data range
    var data = Browser.inputBox(title, msg,
        Browser.Buttons.OK_CANCEL);
    if (data == 'cancel') {
      return;
    }
    return data;
  }
  
  function findTranslateId(table, dataRange, rowIdx) {
    var context = 'findTranslateId: '; 
    if (!table) throw new Error(context + 'table is not set');
    if (!dataRange) throw new Error(context + 'dataRange is not set');
    if (!rowIdx) throw new Error(context + 'rowIdx is not set');
    
    var row = table.getRange(dataRange);
    var rowValues = row.getValues();
    var foundId = rowValues[rowIdx - 1][findColIndexByColName(table, TRANSLATE_ID_NAME)];
    // var en = rowValues[rowIdx - 1][findColIndexByColName(table, 'en')];
    // var de = rowValues[rowIdx - 1][findColIndexByColName(table, 'de')];
    return foundId;
  }
  
  function getRowByTranslateId(table, dataRange, translateId) {
    var context = 'getRowByTranslateId: '; 
    if (!table) throw new Error(context + 'table is not set');
    if (!dataRange) throw new Error(context + 'dataRange is not set');
    if (!translateId) throw new Error(context + 'translateId is not set');
    
    var tableRange = table.getRange(dataRange);
    var tableValues = tableRange.getValues();
  
    var translateIdx = findColIndexByColName(table, TRANSLATE_ID_NAME);
    // Logger.log('translateIdx ' + translateIdx);
    
    var foundRow;
    var foundIdx;
    tableValues.forEach(function (rowValues, idx) {
      if (rowValues[translateIdx].trim() === translateId.trim()) {
        foundRow = rowValues;
        foundIdx = idx;
        // Logger.log('foundRow ' + foundRow);
        // Logger.log('foundIdx ' + foundIdx);
      }
    });
    
    var updatesRuIdx = findColIndexByColName(table, 'ru');
    
    if (!foundRow) return;
    
    return {
      index: foundIdx,
      data: foundRow,
      translations: trimArray(foundRow.slice(updatesRuIdx))
    }
  }
  
  function trimArray(a) {
    if (!a) throw new Error('trimArray: array is not set');
    return a.filter(function(item) {
      return item && item.length > 0;
    })
  }
  
  function findColIndexByColName(table, colName) {
    var context = 'findColIndexByColName: '; 
    if (!table) throw new Error(context + 'table is not set');
    if (!colName) throw new Error(context + 'colName is not set');
    
    var headValues = getTableHead(table);
    var foundIndex;
    headValues.forEach(function (headName, idx) {
      if (headName === colName) foundIndex = idx;
    })
    return foundIndex;
  }
  
  function getTableHead(table) {
    var context = 'getTableHead: '; 
    if (!table) throw new Error(context + 'table is not set');
    
    var head = table.getRange(HEAD_RANGE);
    var headValues = head.getValues();
    return headValues[0];
  }
  