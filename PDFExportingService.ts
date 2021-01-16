function exportInvoiceAsPDF() {

  formatTheActiveInvoice();

  var spreadsheet = SpreadsheetApp.getActive();
  var invoiceSheet = spreadsheet.getActiveSheet();
  var invoiceSheetName = invoiceSheet.getName();

  if (isInvoiceSheetName(invoiceSheetName) && invoiceSheetName != 'Invoice') {
    var pdfFile = exportSheetToPDF(spreadsheet, invoiceSheet, invoiceSheetName);

    // Display a modal dialog box with custom HtmlService content.
    const htmlOutput = HtmlService
      .createHtmlOutput('<p>Click to open <a href="' + pdfFile.getUrl() + '" target="_blank">' + invoiceSheetName + '</a></p>')
      .setWidth(800)
      .setHeight(80)
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful');
  } else {
    console.info('Sheet [%s] is not an invoice', invoiceSheetName);
    Browser.msgBox('Incorrect Invoice', 'Sheet [' + invoiceSheetName + '] is not an invoice. Please choose another sheet! (Không phải invoice)', Browser.Buttons.OK);
  }
}

function exportSheetToPDF(spreadsheet, invoiceSheet: GoogleAppsScript.Spreadsheet.Sheet, invoiceSheetName: string): GoogleAppsScript.Drive.File {

  console.info('Begin export sheet [%s]...', invoiceSheetName);

  var checkOutDate = invoiceSheet.getRange(Invoice.checkOutRangeName).getDisplayValue();
  var checkOutMonth = checkOutDate.substring(3, 5);
  invoiceSheet.expandAllRowGroups();

  var invoiceRange = getInvoiceRange(invoiceSheet);
  var blob = _getAsBlob(spreadsheet.getUrl(), invoiceSheet, invoiceRange);

  var monthFolder = detectMonthlyFolder(spreadsheet, invoiceSheetName);
  var pdfFile = _exportBlob(blob, monthFolder, invoiceSheetName);

  invoiceSheet.getRange(Invoice.linkRangeName).setValue(pdfFile.getUrl());
  invoiceSheet.collapseAllRowGroups();

  return pdfFile;
}


function detectMonthlyFolder(spreadsheet, invoiceName: string) {
  // Invoice Name sample is "15.02.20_R1_Tanja Vivien Gangl"
  var file = DriveApp.getFileById(spreadsheet.getId());
  var invoiceFolder = file.getParents().next();
  var folders = invoiceFolder.getFolders();
  var monthFolder: GoogleAppsScript.Drive.Folder;
  var monthlyInvoiceFolder = DateUtility.invoiceFolderFromInvoiceName(invoiceName);
  while (folders.hasNext()) {
    var tmpFolder = folders.next();
    if (tmpFolder.getName() == monthlyInvoiceFolder) {
      monthFolder = tmpFolder;
      break;
    }
  }

  if (!monthFolder) {
    monthFolder = invoiceFolder.createFolder(monthlyInvoiceFolder);
  }

  var invoiceFiles = monthFolder.getFilesByName(invoiceName);
  if (invoiceFiles.hasNext()) {
    var invoiceFile = invoiceFiles.next();
  }
  if (invoiceFile) {
    //    invoiceFile = folder.createFile(invoiceName);
    monthFolder.removeFile(invoiceFile);
  }
  return monthFolder;
}



function _exportBlob(blob, folder: GoogleAppsScript.Drive.Folder, fileName: string): GoogleAppsScript.Drive.File {
  blob = blob.setName(fileName)
  var pdfFile = folder.createFile(blob);
  return pdfFile;
}



function _getAsBlob(url, sheet, range) {
  var rangeParam = ''
  if (range) {
    rangeParam =
      '&r1=' + (range.getRow() - 1)
      + '&r2=' + range.getLastRow()
      + '&c1=' + (range.getColumn() - 1)
      + '&c2=' + range.getLastColumn()
  }
  var exportUrl = url.replace(/\/edit.*$/, '')
    + '/export?exportFormat=pdf&format=pdf'
    + '&size=STATEMENT'
    + '&portrait=true'
    + '&fitw=true'
    + '&top_margin=0.75'
    + '&bottom_margin=0.75'
    + '&left_margin=0.3'
    + '&right_margin=0.3'
    + '&sheetnames=false&printtitle=false'
    + '&pagenum=false'
    + '&gridlines=true'
    + '&fzr=FALSE'
    + '&gid=' + sheet.getSheetId()
    + rangeParam

  Logger.log('exportUrl=' + exportUrl)
  var response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
  })

  return response.getBlob()
}


function exportAllSheetsAsSeparatePDFs() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var files = []
  spreadsheet.getSheets().forEach(function (sheet) {
    spreadsheet.setActiveSheet(sheet)

    var blob = _getAsBlob(spreadsheet.getUrl(), sheet)
    var fileName = sheet.getName()
    blob = blob.setName(fileName)
    var pdfFile = DriveApp.createFile(blob)

    files.push({
      url: pdfFile.getUrl(),
      name: fileName,
    })
  })

  const htmlOutput = HtmlService
    .createHtmlOutput('<p>Click to open PDF files</p>'
      + '<ul>'
      + files.reduce(function (prev, file) {
        prev += '<li><a href="' + file.url + '" target="_blank">' + file.name + '</a></li>'
        return prev
      }, '')
      + '</ul>')
    .setWidth(300)
    .setHeight(150)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful')
}
function exportCurrentSheetAsPDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var currentSheet = SpreadsheetApp.getActiveSheet()

  var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet)
  _exportBlob(blob, currentSheet.getName())
}



function str() {
  Logger.log('31.12.19'.substring(3, 5));
}

function getInvoiceRange(sheet) {

  var totalRow = SheetUtility.getTotalAmountRow(sheet);
  totalRow = totalRow + 25;

  return sheet.getRange('B2:G' + totalRow);
}

function exportPartAsPDF(predefinedRanges) {
  var ui = SpreadsheetApp.getUi()
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()

  var selectedRanges
  var fileSuffix
  if (predefinedRanges) {
    selectedRanges = predefinedRanges
    fileSuffix = '-predefined'
  } else {
    var activeRangeList = spreadsheet.getActiveRangeList()
    if (!activeRangeList) {
      ui.alert('Please select at least one range to export')
      return
    }
    selectedRanges = activeRangeList.getRanges()
    fileSuffix = '-selected'
  }

  if (selectedRanges.length === 1) {
    // special export with formatting
    var currentSheet = selectedRanges[0].getSheet()
    var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet, selectedRanges[0])

    var fileName = spreadsheet.getName() + fileSuffix
    _exportBlob(blob, fileName)
    return
  }

  var tempSpreadsheet = SpreadsheetApp.create(spreadsheet.getName() + fileSuffix)
  var tempSheets = tempSpreadsheet.getSheets()
  var sheet1 = tempSheets.length > 0 ? tempSheets[0] : undefined
  SpreadsheetApp.setActiveSpreadsheet(tempSpreadsheet)

  for (var i = 0; i < selectedRanges.length; i++) {
    var selectedRange = selectedRanges[i]
    var originalSheet = selectedRange.getSheet()
    var originalSheetName = originalSheet.getName()

    var destSheet = tempSpreadsheet.getSheetByName(originalSheetName)
    if (!destSheet) {
      destSheet = tempSpreadsheet.insertSheet(originalSheetName)
    }

    Logger.log('a1notation=' + selectedRange.getA1Notation())
    var destRange = destSheet.getRange(selectedRange.getA1Notation())
    destRange.setValues(selectedRange.getValues())
    destRange.setTextStyles(selectedRange.getTextStyles())
    destRange.setBackgrounds(selectedRange.getBackgrounds())
    destRange.setFontColors(selectedRange.getFontColors())
    destRange.setFontFamilies(selectedRange.getFontFamilies())
    destRange.setFontLines(selectedRange.getFontLines())
    destRange.setFontStyles(selectedRange.getFontStyles())
    destRange.setFontWeights(selectedRange.getFontWeights())
    destRange.setHorizontalAlignments(selectedRange.getHorizontalAlignments())
    destRange.setNumberFormats(selectedRange.getNumberFormats())
    destRange.setTextDirections(selectedRange.getTextDirections())
    destRange.setTextRotations(selectedRange.getTextRotations())
    destRange.setVerticalAlignments(selectedRange.getVerticalAlignments())
    destRange.setWrapStrategies(selectedRange.getWrapStrategies())
  }

  // remove empty Sheet1
  if (sheet1) {
    Logger.log('lastcol = ' + sheet1.getLastColumn() + ',lastrow=' + sheet1.getLastRow())
    if (sheet1 && sheet1.getLastColumn() === 0 && sheet1.getLastRow() === 0) {
      tempSpreadsheet.deleteSheet(sheet1)
    }
  }

  exportAsPDF()
  SpreadsheetApp.setActiveSpreadsheet(spreadsheet)
  DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true)
}


function exportNamedRangesAsPDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var allNamedRanges = spreadsheet.getNamedRanges()
  var toPrintNamedRanges = []
  for (var i = 0; i < allNamedRanges.length; i++) {
    var namedRange = allNamedRanges[i]
    if (/^print_area_.*$/.test(namedRange.getName())) {
      Logger.log('found named range ' + namedRange.getName())
      toPrintNamedRanges.push(namedRange.getRange())
    }
  }
  if (toPrintNamedRanges.length === 0) {
    SpreadsheetApp.getUi().alert('No print areas found. Please add at least one \'print_area_1\' named range in the menu Data > Named ranges.')
    return
  } else {
    toPrintNamedRanges.sort(function (a, b) {
      return a.getSheet().getIndex() - b.getSheet().getIndex()
    })
    exportPartAsPDF(toPrintNamedRanges)
  }
}
