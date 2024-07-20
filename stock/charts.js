/** ############################### Drive ############################### */

/** @param {DriveApp.Folder} folder */
function clearFolder(folder) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
}

/** @param {String} folder_name */
/** @param {String} [if_exists="ignore"] */
/** @return {DriveApp.Folder} */
function createFolder(folder_name, if_exists="ignore") {
  var folders = DriveApp.getFoldersByName(folder_name);
  if (folders.hasNext()) {
    var folder = folders.next();
    if (if_exists == "replace") {
      clearFolder(folder);
    }
    return folder;
  } else {
    folder = DriveApp.createFolder(folder_name);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return folder;
  }
}

function emptyTrash() {
  try {
    Drive.Files.emptyTrash();
    Logger.log('Trash emptied successfully.');
  } catch (error) {
    Logger.log('Error emptying trash: ' + error);
  }
}

/** ############################### Images ############################## */

/** @param {String} sheet_name */
/** @return {Object.<String, SpreadsheetApp.OverGridImage[]>} */
function getImages(sheet_name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  var images = {};
  sheet.getImages().forEach(image => {
    images[image.getAnchorCell().getA1Notation()] = image;
  });
  return images
}

/** @param {String} sheet_name */
function clearImages(sheet_name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  sheet.getImages().forEach(image => {
    image.remove();
  });
}

/** @param {SpreadsheetApp.Sheet} sheet */
/** @param {Integer} column */
/** @param {Integer} row */
/** @param {String} b64string */
function insertImage(sheet, column, row, b64string) {
  var decodedBytes = Utilities.base64Decode(b64string);
  var blob = Utilities.newBlob(decodedBytes, "image/jpeg", "image.jpg");
  sheet.insertImage(blob, column, row);
}

/** ############################## Groupby ############################## */

/** @param {String} sheet_name */
/** @param {String} a1_notation */
/** @param {Number} key_column */
/** @param {Array.<number>} value_columns */
/** @return {Object.<String, Array.<Array.<number>>>} */
function groupbyValues(sheet_name, a1_notation, key_column, value_columns) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  var values = sheet.getRange(a1_notation).getValues();
  var grouped = {};
  values.forEach(row => {
    if (!(row[key_column] in grouped)) {
      grouped[row[key_column]] = [];
    }
    grouped[row[key_column]].push(value_columns.map(i => { return row[i]; }));
  });
  return grouped;
}

/** @param {SpreadsheetApp.Spreadsheet} spreadsheet */
/** @param {String} sheet_name */
/** @return {SpreadsheetApp.Sheet} */
function createSheet(spreadsheet, sheet_name) {
  var sheet = spreadsheet.getSheetByName(sheet_name);
  if (!sheet) {
    return spreadsheet.insertSheet(sheet_name);
  } else {
    return sheet;
  }
}

/** ############################ Candlestick ############################ */

/** @param {Object.<String, Array.<Array.<number>>>} values */
/** @param {SpreadsheetApp.Sheet} sheet */
/** @param {Integer} [column=1] */
/** @param {Integer} [row=1] */
function setOhlcValues(values, sheet, column=1, row=1) {
  if (sheet.getLastRow() > 0) {
    sheet.getRange(column, row, sheet.getLastRow(), sheet.getLastColumn()).clear();
  }
  values = values.map(row => ["'"+row[0].toISOString().substring(0, 10), ...row.slice(1,)]);
  sheet.getRange(column, row, values.length, 5).setValues(values);
}

/** @param {Object.<String, Array.<Array.<number>>>} values */
/** @param {SpreadsheetApp.Sheet} sheet */
/** @param {Number} [width=500] */
/** @param {Number} [height=100] */
/** @param {String} [rising="green"] */
/** @param {String} [falling="red"] */
/** @return {Blob} */
function drawCandlestick(values, sheet, width=500, height=100, rising="green", falling="red") {
  // https://developers.google.com/chart/interactive/docs/gallery/candlestickchart
  minMax = {minValue: Math.min(...values.map(row => row[1])), maxValue: Math.max(...values.map(row => row[4]))};
  setOhlcValues(values, sheet);

  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.CANDLESTICK)
    .addRange(sheet.getRange(1, 1, values.length, 5))
    .setPosition(1, 1, 0, 0)
    .setOption("width", width)
    .setOption("height", height)
    .setOption("legend", {position: "none"})
    .setOption("chartArea", {width: "100%", height: "100%", left: 0, top: 0})
    .setOption("backgroundColor", {fill: "white"})
    .setOption("colors", [((values[values.length-1][4] > values[0][4]) ? rising : falling)])
    .setOption("hAxis", {textPosition: "none", gridlines: {color: "white"}})
    .setOption("vAxis", {textPosition: "none", gridlines: {color: "white"}, ...minMax})
    .setOption("candlestick", {
      risingColor: {fill: rising , stroke: "black", strokeWidth: 3},
      fallingColor: {fill: falling , stroke: "black", strokeWidth: 3}})
    .build();
  sheet.insertChart(chart);

  var chart = sheet.getCharts()[0];
  var imageBlob = chart.getAs("image/png");
  sheet.removeChart(chart);
  return imageBlob;
}

/** @param {Object.<String, Array.<Array.<number>>>} values */
/** @param {String} sheet_name */
/** @param {String} folder_name */
/** @param {String} [if_exists="replace"] */
/** @param {Integer} [start=2] */
/** @param {Integer} [end=null] */
/** @param {Boolean} [clear=true] */
/** @param {Integer} [limit=null] */
/** @param {Number} [width=500] */
/** @param {Number} [height=100] */
/** @param {String} [rising="green"] */
/** @param {String} [falling="red"] */
function updateCandlestick(values, sheet_name, folder_name, if_exists="ignore", start=2, end=null,
                          clear=true, limit=null, width=500, height=100, rising="green", falling="red") {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheet_name);
  var tempSheet = createSheet(spreadsheet, "Temp"+start);
  var folder = createFolder(folder_name, if_exists);
  if (typeof(start) != "number") { start = 2; }
  if (typeof(end) != "number") { end = sheet.getLastRow(); }
  if (clear) { sheet.getRange(start, 2, end-start+1).clear(); }

  for (let i = start ; i < (end+1) ; i++) {
    let ticker = sheet.getRange(i, 1).getValue();
    let range = sheet.getRange(i, 2);
    if ((ticker in values) & !(range.getValue())) { // SORT BY Date DESC
      imageBlob = drawCandlestick(values[ticker].slice(0,limit).reverse(), tempSheet, width, height, rising, falling);
      var fileName = ticker + "_" + Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd_HHmmss") + ".png";
      var file = folder.createFile(imageBlob.setName(fileName));
      range.setValue("https://drive.google.com/uc?id=" + file.getId());
    }
  }
  spreadsheet.deleteSheet(tempSheet);
}

/** ############################# Sparkline ############################# */

/** @param {Array.<Array.<Number>>} values */
/** @param {Integer} [dimension=0] */
/** @param {Integer} [metric=1] */
/** @return {DataTableBuilder} */
function newTimeSeries(values, dimension=0, metric=1) {
  var dataTable = Charts.newDataTable();
  dataTable.addColumn(Charts.ColumnType.DATE, "Date");
  dataTable.addColumn(Charts.ColumnType.NUMBER, "Value");
  values.forEach(row => { dataTable.addRow([new Date(row[dimension]), row[metric]]); });
  return dataTable;
}

/** @param {Array.<Array.<Number>>} values */
/** @param {Integer} [dimension=0] */
/** @param {Integer} [metric=1] */
/** @param {Integer} [width=500] */
/** @param {Integer} [height=100] */
/** @param {Integer} [line=1] */
/** @param {String} [rising="green"] */
/** @param {String} [falling="red"] */
/** @return {Blob} */
function drawSparkline(values, dimension=0, metric=1, width=500, height=100, line=1, rising="green", falling="red") {
  // https://developers.google.com/chart/interactive/docs/gallery/linechart
  metrics = values.map(row => row[metric]);
  minMax = [Math.min(...metrics), Math.max(...metrics)];

  return Charts.newLineChart()
    .setDataTable(newTimeSeries(values, dimension, metric))
    .setOption("width", width)
    .setOption("height", height)
    .setOption("legend", {position: "none"})
    .setOption("chartArea", {width: "100%", height: "100%", left: 0, top: 0, right: 0, bottom: 0})
    .setOption("backgroundColor", {fill: "transparent"})
    .setOption("colors", [((metrics[metrics.length-1] > metrics[0]) ? rising : falling)])
    .setOption("hAxis", {gridlines: {count: 0, color: "transparent"}, baselineColor: "transparent"})
    .setOption("vAxis", {gridlines: {count: 0, color: "transparent"}, baselineColor: "transparent", ticks: minMax})
    .setOption("lineWidth", line)
    .build().getAs("image/png");
}

/** @param {Object.<String, Array.<Array.<number>>>} values */
/** @param {String} sheet_name */
/** @param {String} folder_name */
/** @param {String} [if_exists="replace"] */
/** @param {Integer} [start=2] */
/** @param {Integer} [end=null] */
/** @param {Boolean} [clear=true] */
/** @param {Integer} [limit=null] */
/** @param {Integer} [width=500] */
/** @param {Integer} [height=100] */
/** @param {Integer} [line=1] */
/** @param {String} [rising="green"] */
/** @param {String} [falling="red"] */
function updateSparkline(values, sheet_name, folder_name, if_exists="ignore", start=2, end=null,
                        clear=true, limit=null, width=500, height=100, line=1, rising="green", falling="red") {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  var folder = createFolder(folder_name, if_exists);
  if (typeof(start) != "number") { start = 2; }
  if (typeof(end) != "number") { end = sheet.getLastRow(); }
  if (clear) { sheet.getRange(start, 3, end-start+1).clear(); }

  for (let i = start ; i < (end+1) ; i++) {
    let ticker = sheet.getRange(i, 1).getValue();
    let range = sheet.getRange(i, 3);
    if ((ticker in values) & !(range.getValue())) { // SORT BY Date DESC
      imageBlob = drawSparkline(values[ticker].slice(0,limit).reverse(), 0, 1, width, height, line, rising, falling);
      var fileName = ticker + "_" + Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd_HHmmss") + ".png";
      var file = folder.createFile(imageBlob.setName(fileName));
      range.setValue("https://drive.google.com/uc?id=" + file.getId());
    }
  }
}
