/** ############################### Fetch ############################### */

/** @param {String} ticker */
/** @param {Integer} start_date */
/** @param {Integer} end_date */
/** @param {Array.<String>} [columns=[]] */
/** @param {Integer} [trunc=2] */
/** @return {Array.<Array.<*>>} */
function fetchYahooPrice(ticker, start_date, end_date, columns=[], trunc=2) {
  var url = "https://query1.finance.yahoo.com/v7/finance/download/" + formatQuery(ticker) +
            "?period1=" + Math.floor(start_date) +
            "&period2=" + Math.floor(end_date) +
            "&interval=1d&events=history&includeAdjustedClose=true";
  var response = UrlFetchApp.fetch(url);
  csvData = Utilities.parseCsv(response.getContentText());
  return parseYahooPrice(csvData, formatTicker(ticker), columns, trunc);
}

/** @param {String} ticker */
/** @param {Array.<String>} [columns=[]] */
/** @param {Integer} [trunc=2] */
/** @return {Array.<Array.<*>>} */
function fetchYahooPrice1y(ticker, columns=[], trunc=2) {
  var startDate = (Date.now() - 365*24*60*60*1000) / 1000;
  var endDate = Date.now() / 1000;
  return fetchYahooPrice(ticker, startDate, endDate, columns, trunc);
}

/** ############################### Parse ############################### */

/** @param {Array.<Array.<*>>} csvData */
/** @param {String} ticker */
/** @param {Array.<String>} [columns=[]] */
/** @param {Integer} [trunc=2] */
/** @return {Array.<Array.<*>>} */
function parseYahooPrice(csvData, ticker, columns=[], trunc=2) {
  var header = ["Ticker"].concat(csvData[0]); // Ticker, Date, Open, High, Low, Close, Adj Close, Volume
  var columns = (columns.length == 0) ? header : columns.filter(e => header.includes(e));
  var data = [columns];
  for (let i = 1 ; i < csvData.length ; i++) {
    let row = [ticker].concat(csvData[i]);
    if (!isNaN(row[5])) {
    data.push(columns.map(e => formatYahooPrice(row[header.indexOf(e)], e, trunc)));
    }
  }
  return data;
}

/** @param {Any} value */
/** @param {String} column */
/** @param {Integer} [trunc=2] */
/** @return {Any} */
function formatYahooPrice(value, column, trunc=2) {
  if (["Open", "High", "Low", "Close"].includes(column)) {
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
/** @param {Array.<String>} [columns=[]] */
/** @param {Integer} [limit=null] */
/** @param {Integer} [trunc=2] */
function setPrice(query_sheet, query_range, return_sheet, return_range, columns=[], limit=null, trunc=2) {
  var querySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(query_sheet);
  var returnSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(return_sheet);
  var query = querySheet.getRange(query_range).getValues().flat().filter(String);
  var data = query.map(ticker => fetchYahooPrice1y(ticker, columns, trunc).slice(1,).reverse().slice(0,limit)).flat(1);
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

function setPriceUs() {
  var columns = ["Ticker", "Open", "High", "Low", "Close", "Date"];
  setPrice("Chart(US)", "A2:A", "Ohlc(US)", "A2:F", columns, 200, 2);
}

function setPriceKr() {
  var columns = ["Ticker", "Open", "High", "Low", "Close", "Date"];
  setPrice("Chart(KR)", "E2:E", "Ohlc(KR)", "A2:F", columns, 240, 0);
}
