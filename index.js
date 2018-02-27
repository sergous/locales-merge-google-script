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
  var UPDATES_TABLE_NAME = 'Первая итерация до 1900 строки';
  
  /** 
   * Запускает заполнение пустых значений в основной таблице из текущей таблицы с обновлениями
   */
  
  function addTranslations() {
    
    var spreadsheet = SpreadsheetApp.getActive();
    
    
    // UPDATES TABLE
    
    // Ask table name with updates  
    var updateTableName = UPDATES_TABLE_NAME;
    // var updateTableName = askData('Укажите таблицу переводов', 'Укажите название таблицы с новыми переводами' + ' (например, "' + UPDATES_TABLE_NAME + '"):');
    
    var updatesTable = spreadsheet.getSheetByName(updateTableName);
    
    var updatesDataRange = TABLE_RANGE;
    // var updatesDataRange = askData('Укажите диапазон', 'Укажите диапазон данных в таблице новых переводов' + ' (например, "A1:J1958"):');
    
    
  
    // BASE TABLE
    var baseTable = spreadsheet.getSheetByName(BASE_TABLE_NAME);
    baseTable.activate();
    
    // translateIdResults(baseTable, translateId);
    
    var startRowIdx = 2;
    // var startRowIdx = askData('Укажите начальную строку перевода', 'Укажите номер первой строки в таблице новых переводов ' + UPDATES_TABLE_NAME + ' (например, "2"):');
    
    var endRowIdx = 10;
    // var endRowIdx = askData('Укажите конечную строку перевода', 'Укажите номер последней строки в таблице новых переводов ' + UPDATES_TABLE_NAME + ' (например, "100"):');
    
    var statFoundLinesIdx = [];
    var statNotFoundLineIds = [];
    
    var updatesTranslateIdIdx = findColIndexByColName(updatesTable, TRANSLATE_ID_NAME);
      
    var baseRuIdx = findColIndexByColName(baseTable, 'ru');
    if (!baseRuIdx) throw new Error('addTranslations: baseRuIdx is not set');
    
    updatesTable.activate();
    
    for(var idx = startRowIdx; idx <= endRowIdx; idx++) {
      var replacedIdx = replaceRow(idx, baseRuIdx);
      
      if (!replacedIdx) {
        setTableCellBg(UPDATES_TABLE_NAME, idx, 1);
        statNotFoundLineIds.push(idx);
        break;
      }
  
      statFoundLinesIdx.push(replacedIdx + 1); 
    }
    
    Browser.msgBox('Результы обновления переводов', 
                   'Обновлены строки (' + statFoundLinesIdx.length + '): ' + statFoundLinesIdx.join(',') + ' | ' +
                   'Не найдены ключи: (' + statNotFoundLineIds.length + '): ' + statNotFoundLineIds.join(',')
                   , Browser.Buttons.OK);
  
    function setTableCellBg(tableName, rowIdx, colIdx, color) {
      var context = 'setTableCellBg: '; 
      if (!tableName) throw new Error(context + 'tableName is not set');
      if (!rowIdx) throw new Error(context + 'rowIdx is not set');
      if (!colIdx) throw new Error(context + 'colIdx is not set');
      if (!color) color = "red";
      
      var spreadsheet = SpreadsheetApp.getActive();
      var table = spreadsheet.getSheetByName(tableName);
      var range = table.getRange(rowIdx, colIdx, 1, 1);
      range.setBackground(color);
      Logger.log('Marked as error: row ' + rowIdx + 1 + ' column ' + colIdx);
    }
    
    function replaceRow(rowId, baseRuIdx) {
      var context = 'replaceRow: '; 
      if (!rowId) throw new Error(context + 'rowId is not set');
      
      // Get translate id 
      var translateId = findTranslateId(updatesTable, updatesDataRange, rowId);
      Logger.log('translateId ' + translateId);
      
      // var updateColName = 'de';
      // var updateColName = askData('Укажите язык', 'Укажите язык' + ' (например, "de"):');
      
      // var updatesColIdx = findColIndexByColName(updatesTable, updateColName);
      // Logger.log('RU ' + updatesRuId + ' DE ' + updatesDeId);
      
      // Found updates table row values for translateId
      // var foundUpdatesRow = getRowByTranslateId(updatesTable, updatesDataRange, translateId).data;
      // Logger.log('foundUpdatesRow ' + foundUpdatesRow);
      
      var foundUpdatesTranslations = getRowByTranslateId(updatesTable, updatesDataRange, translateId).translations;
      Logger.log('foundUpdatesTranslations ' + foundUpdatesTranslations);
      
      // var updatedValue = foundUpdatesRow[updatesColIdx];
      // Logger.log('updateColName ' + updateColName + ' updatedValue ' + updatedValue);
      
      // Found base table row with translateId
      var baseRow = getRowByTranslateId(baseTable, TABLE_RANGE, translateId);
      if (!baseRow) {
        return;
      } else {
        var baseRowIdx = baseRow.index;
        
        replaceTableRow(baseTable, baseRowIdx + 1, baseRuIdx + 1, foundUpdatesTranslations);
        return baseRowIdx;
      } 
    }
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
  
  function findTranslateId(table, dataRange, rowId) {
    var context = 'findTranslateId: '; 
    if (!table) throw new Error(context + 'table is not set');
    if (!dataRange) throw new Error(context + 'dataRange is not set');
    if (!rowId) throw new Error(context + 'rowId is not set');
    
    var row = table.getRange(dataRange);
    var rowValues = row.getValues();
    var foundId = rowValues[rowId - 1][findColIndexByColName(table, TRANSLATE_ID_NAME)];
    // var en = rowValues[rowId - 1][findColIndexByColName(table, 'en')];
    // var de = rowValues[rowId - 1][findColIndexByColName(table, 'de')];
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
    Logger.log('translateIdx ' + translateIdx);
    
    var foundRow;
    var foundIdx;
    tableValues.forEach(function (rowValues, idx) {
      if (rowValues[translateIdx].trim() === translateId.trim()) {
        foundRow = rowValues;
        foundIdx = idx;
        Logger.log('foundRow ' + foundRow);
        Logger.log('foundIdx ' + foundIdx);
      }
    });
    
    var updatesRuIdx = findColIndexByColName(table, 'ru');
    
    if (!foundRow) return;
    
    return {
      index: foundIdx,
      data: foundRow,
      translations: foundRow.slice(updatesRuIdx)
    }
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
  