/** ############################### Fetch ############################### */

/** @param {String} ticker */
/** @param {Integer} start_date */
/** @param {Integer} end_date */
/** @param {String} [events="history"] */
/** @param {Array.<String>} [columns=[]] */
/** @param {Integer} [trunc=2] */
/** @return {Array.<Array.<*>>} */
function fetchYahooData(ticker, start_date, end_date, events="history", columns=[], trunc=2) {
  var url = "https://query1.finance.yahoo.com/v7/finance/download/" + formatQuery(ticker) +
            "?period1=" + Math.floor(start_date) +
            "&period2=" + Math.floor(end_date) +
            "&interval=1d&events=" + events;
  var response = UrlFetchApp.fetch(url);
  csvData = Utilities.parseCsv(response.getContentText());
  if (csvData.length > 1) {
    return parseYahooData(csvData, formatTicker(ticker), events, columns, trunc);
  } else {
    return [];
  }
}

/** @param {String} ticker */
/** @param {String} [events="history"] */
/** @param {Array.<String>} [columns=[]] */
/** @param {Integer} [trunc=2] */
/** @return {Array.<Array.<*>>} */
function fetchYahooData1y(ticker, events="history", columns=[], trunc=2) {
  var startDate = (Date.now() - 365*24*60*60*1000) / 1000;
  var endDate = Date.now() / 1000;
  return fetchYahooData(ticker, startDate, endDate, events, columns, trunc);
}

/** ############################### Parse ############################### */

/** @param {Array.<Array.<*>>} csvData */
/** @param {String} ticker */
/** @param {String} [events="history"] */
/** @param {Array.<String>} [columns=[]] */
/** @param {Integer} [trunc=2] */
/** @return {Array.<Array.<*>>} */
function parseYahooData(csvData, ticker, events="history", columns=[], trunc=2) {
  var header = ["Ticker"].concat(csvData[0]); // Ticker, Date, Open, High, Low, Close, Adj Close, Volume
  var columns = (columns.length == 0) ? header : columns.filter(e => header.includes(e));
  var data = [columns];
  for (let i = 1 ; i < csvData.length ; i++) {
    let row = [ticker].concat(csvData[i]);
    if ((events != "history") | !isNaN(row[5])) {
      data.push(columns.map(e => formatYahooData(row[header.indexOf(e)], e, trunc)));
    }
  }
  return data;
}

/** @param {Any} value */
/** @param {String} column */
/** @param {Integer} [trunc=2] */
/** @return {Any} */
function formatYahooData(value, column, trunc=2) {
  if (["Open", "High", "Low", "Close", "Dividends"].includes(column)) {
    return Number(value).toFixed(trunc);
  } else if (column == "Volume") {
    return Number(value).toFixed(0);
  } else {
    return value;
  }
}

/** @param {String} ticker */
/** @return {String} */
function formatQuery(ticker) {
  if (ticker.endsWith(".KS") | ticker.endsWith(".KQ")) {
    return ticker;
  } else {
    return ticker.replace(".","-");
  }
}

/** @param {String} ticker */
/** @return {String} */
function formatTicker(ticker) {
  if (ticker.endsWith(".KS") | ticker.endsWith(".KQ")) {
    return ticker.replace(".KS", "").replace(".KQ", "");
  } else {
    return ticker;
  }
}

/** ############################ Spreadsheet ############################ */

/** @param {String} query_sheet */
/** @param {String} query_column */
/** @param {String} return_sheet */
/** @param {String} return_range */
/** @param {String} [events="history"] */
/** @param {Array.<String>} [columns=[]] */
/** @param {Integer} [limit=null] */
/** @param {Integer} [trunc=2] */
function setData(query_sheet, query_range, return_sheet, return_range, events="history", columns=[], limit=null, trunc=2) {
  var querySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(query_sheet);
  var returnSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(return_sheet);
  var query = querySheet.getRange(query_range).getValues().flat().filter(String);
  var data = query.map(ticker => fetchYahooData1y(ticker, events, columns, trunc).slice(1,).reverse().slice(0,limit)).flat(1);
  if (data.length > 0) {
    replaceValues(data, returnSheet, returnSheet.getRange(return_range));
  }
}

/** @param {Array.<Array.<*>>} values */
/** @param {SpreadsheetApp.Sheet} sheet */
/** @param {SpreadsheetApp.Range} range */
function replaceValues(values, sheet, range) {
  var numColumns = range.getLastColumn() - range.getColumn() + 1;
  range.clearContent();
  sheet.getRange(range.getRow(), range.getColumn(), values.length, numColumns).setValues(values);
}
