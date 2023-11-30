//var toClass_ = {}.toString;
//function objIsClass_(object,class) {
//  return (toClass_.call(object).indexOf(class) !== -1);
//}

function init(tzone, locale) {
  this.tzone = tzone || this.tzone || Session.getScriptTimeZone();
  this.locale = locale || this.locale || Session.getActiveUserLocale();
  return this;
}

function _testRange() {
  var ss = SpreadsheetApp.openById('1rSkjqujnMfcuhIN2dWnKC0WH-L7lUjYoF_Mvub_zmNQ');
  sheet = ss.getSheetByName('FY-2023');
  range = sheet.getRange('A1:E15');
  var converter = init(ss.getSpreadsheetTimeZone(), ss.getSpreadsheetLocale());
  var html = converter._range2html(range);
  var wraps = range.getWraps();
  // for (i = 0; wraps.lentgh; i++) {
  //   Logger.log(wraps)
  //   for (j = 0; wraps[i].lentgh; j++) {
  //     Logger.log('word-break:XXX;'.replace('XXX', wraps[i][j] ? 'break-word' : 'normal'))
  //   }
  // }
  //  Logger.log(wraps)
  //  MailApp.sendEmail({
  //    to: "nkhan@noon.com",
  //    subject: "Logos",
  //    htmlBody: html
  //    
  //  });
}


function _range2html(range) {
  var ss = range.getSheet().getParent();
  var sheet = range.getSheet();
  startRow = range.getRow();
  startCol = range.getColumn();
  lastRow = range.getLastRow();
  lastCol = range.getLastColumn();

  var converter = this.init();

  //  var data = range.getValues();
  var data = range.getDisplayValues();

  var mergedRanges = range.getMergedRanges();
  var mappedMergedRanges = {};
  var a1Notation = '';
  for (var iter = 0; iter < mergedRanges.length; iter++) {
    a1Notation = mergedRanges[iter].getA1Notation()
    mappedMergedRanges[a1Notation] = a1Notation;
  }

  var fontColors = range.getFontColors();
  var backgrounds = range.getBackgrounds();
  var fontFamilies = range.getFontFamilies();
  var fontSizes = range.getFontSizes();
  var fontLines = range.getFontLines();
  var fontStyles = range.getFontStyles();
  var fontWeights = range.getFontWeights();
  var horizontalAlignments = range.getHorizontalAlignments();
  var verticalAlignments = range.getVerticalAlignments();
  var wraps = range.getWraps();

  var colWidths = [];
  var tableWidth = 0;
  for (var col = startCol; col <= lastCol; col++) {
    colWidths.push(120 == sheet.getColumnWidth(col) ? 100 : sheet.getColumnWidth(col));
    tableWidth += colWidths[colWidths.length - 1];
  }

  var rowHeights = [];
  for (var row = startRow; row <= lastRow; row++) {
    rowHeights.push(17 == sheet.getRowHeight(row) ? 21 : sheet.getRowHeight(row));
  }

  var numberFormats = range.getNumberFormats();

  //var wraps = range.getWraps();

  var tableFormat = 'cellspacing="0" cellpadding="0" dir="ltr" border="1" style="width:TABLEWIDTHpx;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:right;text-decoration:none;font-style:normal;"';

  var html = ['<table ' + tableFormat + '>'];

  html.push('<colgroup>');
  for (col = 0; col < colWidths.length; col++) {
    html.push('<col width=XXX>'.replace('XXX', colWidths[col]))
  }
  html.push('</colgroup>');
  html.push('<tbody>');

  var invalidRows = [];
  for (_row = 0; _row < data.length; _row++) {
    var validRow = false;
    for (_col = 0; _col < data[_row].length; _col++) {
      if (data[_row][_col]) {
        validRow = true;
        break;
      }
    }
    if (!validRow) {
      invalidRows.push(_row);
    }
  }

  for (var _row = 0; _row < data.length; _row++) {
    html.push('<tr style="height:XXXpx;vertical-align:bottom;">'.replace('XXX', rowHeights[_row]));
    var _col = 0;
    while (_col < data[_row].length) {
      var columnsToSpan = 0;
      let rowsToSpan = 0
      var cellMergedRange = range.getCell(_row + 1, _col + 1).getMergedRanges();
      if (cellMergedRange && Array.isArray(cellMergedRange) && cellMergedRange.length && cellMergedRange[0].getA1Notation() && mappedMergedRanges[cellMergedRange[0].getA1Notation()]) {
        columnsToSpan = cellMergedRange[0].getNumColumns();
        rowsToSpan = cellMergedRange[0].getNumRows()
      }
      //      var cellText = converter.convertCell(data[_row][_col],numberFormats[_row][_col],true);
      var cellText = data[_row][_col];
      var _style = 'style="'
        + 'padding:2px 3px; '
        + 'color:XXX;'.replace('XXX', fontColors[_row][_col].replace('general-', '')).replace('color:black;', '')
        + 'font-family:XXX;'.replace('XXX', fontFamilies[_row][_col]).replace('font-family:arial,sans,sans-serif;', '')
        + 'font-size:XXXpt;'.replace('XXX', fontSizes[_row][_col]).replace('font-size:10pt;', '')
        + 'font-weight:XXX;'.replace('XXX', fontWeights[_row][_col]).replace('font-weight:normal;', '')
        + 'background-color:XXX;'.replace('XXX', backgrounds[_row][_col]).replace('background-color:white;', '')
        + 'text-align:XXX;'.replace('XXX', horizontalAlignments[_row][_col].replace('general-', '').replace('general', 'center')).replace('text-align:right;', '')
        + 'vertical-align:XXX;'.replace('XXX', verticalAlignments[_row][_col]).replace('vertical-align:bottom;', '')
        + 'text-decoration:XXX;'.replace('XXX', fontLines[_row][_col]).replace('text-decoration:none;', '')
        + 'font-style:XXX;'.replace('XXX', fontStyles[_row][_col]).replace('font-style:normal;', '')
        //                + 'word-break:XXX;'.replace('XXX',wraps[_row][_col] ? 'break-word' : 'normal')
        + 'border:1px solid black;'
        + 'overflow:hidden;'
        + '"';
      if (columnsToSpan) {
        html.push('<td colspan="' + columnsToSpan + '" XXX>'.replace('XXX', _style) + String(cellText) + '</td>');
        _col += columnsToSpan - 1;
      } else {
        if (invalidRows.indexOf(row) === -1) {
          html.push('<td XXX>'.replace('XXX', _style) + String(cellText) + '</td>');

        } else {
          _style = 'style="'
            + 'border: 1px solid #FFFFFF;'
            + 'border-bottom: 1px solid #000000;'
            + '"';
          html.push('<td XXX>'.replace('XXX', _style) + '</td>');

        }
      }
      _col++;
    }
    html.push('</tr>');
  }
  html.push('</tbody>');
  html.push('</table>');

  return html.join('').replace('TABLEWIDTH', tableWidth);
}

function convertCell(cellText, format, htmlReady) {
  if (arguments.length < 2 || !objIsClass_(format, "String"))
    throw new Error('Invalid parameter(s)');

  htmlReady = htmlReady || false;
  this.init();

  if (cellText === null) return '';

  if (objIsClass_(cellText, "Date")) {
    return (convertDateTime_(cellText, format));
  }

  if (objIsClass_(cellText, "Number")) {
    if (format === "0.###############" || format === '') {
      if (Math.abs(cellText) >= 1000000000000010)
        return convertExponential_(cellText, 5);
      else
        return String(cellText);
    }

    if (format === '@') format = '0.###############';

    var re = /^([#0,]+)([\.]?)([#0,]*)$/
    var paddedDecimal = re.test(format);

    if (paddedDecimal) {
      var thous = format.match(/,/) ? ',' : '';
      format = format.replace(/,/g, '');
      var parts = format.match(re);
      var whole = parts[1];
      var wholeMin = whole.replace(/[^0]/g, '').length;
      var wholeMax = whole.length;
      var fract = parts[3];
      var fractMin = fract.replace(/[^0]/g, '').length;
      var fractMax = fract.length;
      return convertPadded_(cellText, fractMax, fractMin, wholeMin, thous);
    }

    if (format.indexOf('$') !== -1) {
      var options = { htmlReady: htmlReady };

      if (format.slice(-1) === "]") options.symLoc = "after";

      var matches = format.match(/\[\$(.*?)\]/);
      if (matches) options.symbol = matches[1];
      var thous = format.match(/,/) ? ',' : '';
      format = format.replace(/,/g, '');

      matches = format.match(/\.(0*?)($|[^0])/);
      var fract = matches ? matches[1].length : 0;

      matches = format.match(/\(.*\)/);
      if (matches) options.negBrackets = true;

      matches = format.match(/;\[(.*?)\]/);
      if (matches) options.negColor = matches[1];

      return convertCurrency_(cellText, fract, thous, options);
    }

    if (format.indexOf('%') !== -1) {
      var matches = format.match(/\.(0*?)%/);
      var fract = matches ? matches[1].length : 0;
      return convertPercent_(cellText, fract);
    }

    var expon = format.match(/\.(0*?)E\+/);
    if (expon) {
      var fract = expon[1].length;
      return convertExponential_(cellText, fract);
    }

    if (format.indexOf('?\/?') !== -1) {
      matches = format.match(/(\?*?)\//);
      var precision = matches ? matches[1].length : 1;
      return convertFraction_(cellText, precision);
    }

    if (this[format]) {
      return converter_[format](cellText);
    }
    else {
      Logger.log("Unsupported format '" + format + "', cell='" + cellText + "'");
      return cellText;
    }
  }

  var result = String(cellText);
  if (htmlReady) result = result.replace(/ /g, "&nbsp;").replace(/</g, "&lt;").replace(/\n/g, "<br>");
  return result;
}

function convertDateTime_(date, format) {
  if ('' == format) format = 'M/d/yyyy';

  if (format.indexOf(/am\/pm|AM\/PM/) === -1) {
    format = format.replace(/h/g, 'H');
  }

  if (format.indexOf("[") !== -1)
    format = updFormatElapsedTime_(date, format);

  var jsFormat = format
    .replace(/am\/pm|AM\/PM/, 'a')
    .replace('dddd', 'EEEE')
    .replace('ddd', 'EEE')
    .replace(/S/g, 's')
    .replace(/D/g, 'd')
    .replace(/M/g, 'm')
    .replace(/([hH]+)"*(.)"*(m+)/g, tempMinute_)
    .replace(/(m+)"*(.)"*(s+)/g, tempMinute_)
    .replace('mmmmm', '"@"MMM"@"')
    .replace(/m/g, "M")
    .replace(/b/g, 'm')
    .replace(/0+/, 'S')
    .replace(/"/g, "'")
  var result = Utilities.formatDate(
    date,
    this.tzone,
    jsFormat)
    .replace(/@.*@/g, firstChOfMonth_);
  return result;
}

function tempMinute_(match) {
  return match.replace(/m/g, 'b');
}

function firstChOfMonth_(match) {
  return match.charAt(1)
}

function updFormatElapsedTime_(date, format) {
  var elapsedMs = getMsSinceMidnight_(date);
  var matches = format.match(/\[([sS]+)\]/);
  var pad = matches ? matches[1].length : 1;
  var elapsedSec = convertPadded_(Math.floor(elapsedMs / 1000), 0, 0, pad);

  matches = format.match(/\[([mM]+)\]/);
  pad = matches ? matches[1].length : 1;
  var elapsedMin = convertPadded_(Math.floor(elapsedMs / 60000), 0, 0, pad);

  var format = format.replace(/\[([hH]+)\]/, elapsedHours_)
    .replace(/\[([mM]+)\]/, elapsedMin)
    .replace(/\[([sS]+)\]/, elapsedSec)
  return format;
}

function elapsedHours_(match) {
  return match.replace(/[hH]/g, 'H').replace(/[\[\]]/g, '');
}

function getMsSinceMidnight_(d) {
  var e = new Date(d);
  return d - e.setHours(0, 0, 0, 0);
}

function convertPadded_(num, fractMax, fractMin, wholeMin, thous) {
  fractMin = fractMin || 0;
  wholeMin = wholeMin || 1;
  thous = thous || '';
  var numStr = String(1 * Utilities.formatString("%.Xf".replace('X', String(fractMax)), num));
  var parts = numStr.split('.');
  var whole = pad0_(parts[0], wholeMin, true);
  var frac = pad0_((parts.length > 1) ? parts[1] : '', fractMin);
  var thouGroups = /(\d+)(\d{3})/;
  while (thous && thouGroups.test(whole)) {
    whole = whole.replace(thouGroups, '$1' + thous + '$2');
  }
  var result = whole + (frac ? ('.' + frac) : '');
  return result;
}

function pad0_(num, width, left) {
  var num = String(num);

  if (num.length >= width) return num;
  var bunchazeros = '0000000000000000000000000000000000000';
  if (left) {
    var result = (bunchazeros + num).substr(-width);
  } else {
    result = (num + bunchazeros).substr(0, width);
  }
  return result;
}

function convertExponential_(num, fract) { return num.toExponential(fract).replace('e', 'E'); }

function convertPercent_(num, fract) { return Utilities.formatString("%.Xf%".replace('X', String(fract)), 100 * num); }

function convertCurrency_(num, fract, thous, options) {
  options = options || {};
  thous = thous || '';
  var result = "#RESULT#";
  var symbol = options.symbol ? options.symbol : '$';
  if (!options.symLoc || options.symLoc === 'before') {
    result = symbol + "#RESULT#";
  } else if (options.symLoc === 'after') {
    result = "#RESULT#" + symbol;
  } else { }
  if (num < 0) {
    num = -num;
    if (options.negBrackets) {
      result = "(" + result + ")";
    } else {
      result = '-' + result;
    }
  }
  if (options.negColor && options.htmlReady) {
    result = ("<span style=\"color:XXX;\">" + result + "</span>").replace("XXX", options.negColor.toLowerCase());
  }
  num = convertPadded_(num, fract, fract, 1, thous);
  return result.replace("#RESULT#", num);
}

function convertFraction_(num, precision) {
  if (!this.fracEst) this.fracEst = new FractionEstimator_();
  var sign = (num < 0) ? -1 : 1;
  num = sign * num;
  var whole = Math.floor(num);
  var frac = num % 1;
  var result = ((whole === 0) ? '' : String(sign * whole) + ' ') + this.fracEst.estimate(frac, precision);
  return result
}

var converter_ = {};

converter_["#,##0.00;(#,##0.00)"] = function (num) {
  if (num > 0) {
    var result = 'XXX';
  }
  else {
    num = -num;
    result = '(XXX)';
  }
  return result.replace('XXX', Utilities.formatString("%.2f", num));
}

function FractionEstimator_() {
  this.fracList = {};
}

FractionEstimator_.prototype.estimate = function (value, precision) {
  if (1 <= value || 0 > value) throw new Error('invalid fraction, 0 < fraction < 1');
  precision = precision || 1;
  if (precision > 2) throw new Error('beyond max precision');

  var list = this.fracList_(precision);

  var lo = 0, hi = list.length - 1;
  while (lo < hi) {
    var mid = (lo + hi) >> 1;
    if (value < list[mid].val) hi = mid;
    else lo = mid + 1;
  }

  if (Math.abs(list[lo - 1].val - value) < Math.abs(list[lo].val - value))
    var frac = list[lo - 1].frac;
  else
    frac = list[lo].frac;
  return frac;
}

FractionEstimator_.prototype.fracList_ = function (precision) {
  if (!this.fracList[precision]) {
    var max = Math.pow(10, precision);
    var list = [];
    for (var denom = 2; denom < max; denom++) {
      for (var nom = 1; nom < denom; nom++) {
        var dec = nom / denom;
        if (!list[dec]) list[dec] = Utilities.formatString("%u/%u", nom, denom);
        if (!(denom % 2)) nom++;
      }
    }
    var a = Object.keys(list).sort();
    this.fracList[precision] = [];
    for (var i = 0; i < a.length; i++) {
      this.fracList[precision].push({ "val": parseFloat(a[i]), "frac": list[a[i]] });
    }
  }
  return this.fracList[precision];
}















