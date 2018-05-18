var ByteArrayOutputStream = Java.type('java.io.ByteArrayOutputStream');
var CellType = Java.type('org.apache.poi.ss.usermodel.CellType');
var HorizontalAlignment = Java.type('org.apache.poi.ss.usermodel.HorizontalAlignment');
var XSSFFont = Java.type('org.apache.poi.xssf.usermodel.XSSFFont');
var XSSFWorkbook = Java.type('org.apache.poi.xssf.usermodel.XSSFWorkbook');

var DEFAULT_DATE_FORMAT = 'DD/MM/yyyy';
var DEFAULT_TIME_FORMAT = 'HH:MM';
var DEFAULT_DATETIME_FORMAT = DEFAULT_DATE_FORMAT + ' ' + DEFAULT_TIME_FORMAT;

var DEFAULT_CURRENCY_FORMAT = 'R$ #,##0.00';

var FONT_OPTION_NAMES = ['fontName', 'fontSize', 'bold', 'italic', 'striked', 'underline', 'doubleUnderline'];
var FONTS_CACHE = {};

function create(rows, metadata) {
  metadata = Object.assign({
    asByteArray: true,
    autoSize: true,
    hasHeader: true
  }, metadata);

  var wb;

  try {
    wb = new XSSFWorkbook();

    var createHelper = wb.getCreationHelper();
    var sheet = wb.createSheet();

    var columnsMD = metadata.columns;
    var generalStyle = metadata.style;
    var headerStyleMD = metadata.headerStyle;

    rows.forEach(function (row, index) {
      var sheetHeader = metadata.hasHeader && index == 0 ? sheet.createRow(0) : null;
      var sheetRow = sheet.createRow(metadata.hasHeader ? index + 1 : index);

      Object.keys(row).forEach(function (key, cellIndex) {
        var value = row[key];
        var valueClass = value && value.constructor.name;

        var columnMD = columnsMD && columnsMD[key];

        if (sheetHeader != null) {
          var style = Object.assign({}, generalStyle, columnMD, headerStyleMD);

          var headerCell = createCell(wb, sheetHeader, cellIndex, style, createHelper, valueClass);

          if (columnMD && columnMD.description) {
            headerCell.setCellValue(columnMD.description);
          } else {
            headerCell.setCellValue(key);
          }
        }

        var style = Object.assign({}, generalStyle, columnMD);
        var cell = createCell(wb, sheetRow, cellIndex, style, createHelper, valueClass);
        setTypedValue(cell, value, valueClass);
      });
    });

    if (metadata.autoSize) {
      var headerRow = sheet.getRow(0);
      headerRow.cellIterator().forEachRemaining(function (cell) {
        sheet.autoSizeColumn(cell.getColumnIndex());
      });
    }

    if (metadata.asByteArray) {
      var baos = new ByteArrayOutputStream();
      wb.write(baos);
      return baos.toByteArray();
    } else {
      return wb;
    }
  } finally {
    if (wb) {
      wb.close();
    }
  }
}

function createCell(wb, sheetRow, cellIndex, style, createHelper, valueClass) {
  var cell = sheetRow.createCell(cellIndex);

  var cellStyle = wb.createCellStyle();
  updateCellStyle(wb, cellStyle, style, createHelper, valueClass);
  cell.setCellStyle(cellStyle);

  return cell;
}

function updateCellStyle(wb, cellStyle, options, createHelper, valueClass) {
  if (!options) {
    return;
  }

  var format = resolveDataFormat(options, valueClass);

  if (format) {
    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(format));
  }

  var font = resolveFont(wb, options);

  if (font) {
    cellStyle.setFont(font);
  }

  if (options.horizontalAlignment) {
    cellStyle.setAlignment(HorizontalAlignment[options.horizontalAlignment.toUpperCase()]);
  }
}

function resolveFont(wb, options) {
  if (!options) {
    return undefined;
  }

  var fontOpts = FONT_OPTION_NAMES.reduce(function (map, optName) {
    if (options[optName]) {
      map[optName] = options[optName];
    }

    return map;
  }, {});

  var font;

  if (Object.keys(fontOpts).length) {
    var key = JSON.stringify(fontOpts);

    font = FONTS_CACHE[key];

    if (!font) {
      font = wb.createFont();

      font.setFontName(fontOpts.fontName || XSSFFont.DEFAULT_FONT_NAME);
      font.setFontHeightInPoints(fontOpts.fontSize || XSSFFont.DEFAULT_FONT_SIZE);

      if (fontOpts.bold) {
        font.setBold(fontOpts.bold);
      }

      if (fontOpts.italic) {
        font.setItalic(fontOpts.italic);
      }

      if (fontOpts.striked) {
        font.setStrikeout(fontOpts.striked);
      }

      if (fontOpts.underline) {
        font.setUnderline(0x21);
      }

      if (fontOpts.doubleUnderline) {
        font.setUnderline(0x22);
      }

      FONTS_CACHE[key] = font;
    }
  }

  return font;
}

function resolveDataFormat(options, valueClass) {
  if (options.format) {
    return options.format;
  }

  switch (options.type) {
    case 'currency':
      return DEFAULT_CURRENCY_FORMAT;
    case 'time':
      return DEFAULT_TIME_FORMAT;
    case 'date':
      return DEFAULT_DATE_FORMAT;
    case 'datetime':
      return DEFAULT_DATETIME_FORMAT;
    default:
      var isDate = valueClass === 'Date';
      return isDate ? DEFAULT_DATETIME_FORMAT : null;
  }
}

function setTypedValue(cell, value, valueClass) {
  if (value === null || value === undefined) {
    cell.setCellType(CellType.BLANK);
    return;
  }

  if (valueClass == 'Date') {
    value = new java.util.Date(value.getTime());
  }

  cell.setCellValue(value);
}

exports = create;