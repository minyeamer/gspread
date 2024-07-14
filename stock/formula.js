/** ############################# IMPORTJSON ############################ */

/** @param {String} url */
/** @param {String} path */
function IMPORTJSON(url, path) {
  json = fetchJson(url);
  data = hierGet(json, path.split(','));
  if (Array.isArray(data)) {
    return toTable(data);
  } else {
    return data;
  }
}

/** @param {String} url */
function fetchJson(url) {
  var response = UrlFetchApp.fetch(url);
  var content = response.getContentText();
  return JSON.parse(content);
}

/** @type {Object} data */
/** @type {Array.<String>} path  */
function hierGet(data, path) {
  for (var i = 0; i < path.length; i++) {
    if (data[path[i]]) {
      data = data[path[i]];
    } else {
      return "";
    }
  }
  return data;
}

/** @type {Map[]} data */
function toTable(data) {
  var output = [];
  var headers = Object.keys(data[0]);
  output.push(headers);

  for (var i = 0; i < data.length; i++) {
    var row = [];
    for (var j = 0; j < headers.length; j++) {
      row.push(data[i][headers[j]]);
    }
    output.push(row);
  }
  return output;
}

/** ############################## KONUMBER ############################# */

/** @param {Integer} number */
function KONUMBER(number) {
  if (/^\d+$/.test(number)) {
    const koreanUnits = ['', '만', '억', '조'];
    const unit = 10000;
    for (var i = 3; i > 0; i--) {
      if (number >= unit**i) {
        return (Math.round(number/(unit**i)*10)/10) + koreanUnits[i]
      }
    }
    return number;
  } else {
    return null;
  }
}
