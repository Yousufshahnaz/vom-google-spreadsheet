function UntitledMacro() {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange('J16').activate();

    //Logger.log(spreadsheet.getRange('J16').getLastRow());
};

function onEdit(e) {
    var range = e.range;
    var spreadSheet = SpreadsheetApp.getActive();

    var invoiceSheet = spreadSheet.getActiveSheet();
    var invoiceSheetName = invoiceSheet.getName();

    if (isInvoiceSheetName(invoiceSheetName)) {
        var currentRow = range.getLastRow();

        if (nextIsTotal(invoiceSheet, currentRow)) {
            insertItemAfter(spreadSheet, currentRow);
            var totalRow = currentRow + 2;
            invoiceSheet.getRange('G' + totalRow).setValue('=SUM(G13:G' + currentRow + ')');
        }
    }
}

function formatTheActiveInvoice() {

    var spreadSheet = SpreadsheetApp.getActive();
    var spreadSheetName = spreadSheet.getName();

    var invoiceSheet = spreadSheet.getActiveSheet();
    var invoiceSheetName = invoiceSheet.getName();


    if (isInvoiceSheetName(invoiceSheetName)) {

        var range = spreadSheet.getActiveCell();
        var currentRow = range.getLastRow();

        if (nextIsTotal(invoiceSheet, currentRow)) {
            insertItemAfter(spreadSheet, currentRow);
        }
        // Remove fragment rows, which has NO item name
        var r = Invoice.itemsBeginRow;
        while (!nextIsTotal(invoiceSheet, r)) {
            var itemName = spreadSheet.getRange(Invoice.itemNameColumnName + r).getValue();
            if (!itemName) {
                removeRows(spreadSheet, r, 1);
            } else {
                r = r + 1;
            }
        }

        // Correct item index 1,2,3,..
        r = Invoice.itemsBeginRow;
        var i = 1;
        while (!nextIsTotal(invoiceSheet, r)) {
            spreadSheet.getRange(Invoice.itemOrderColumnName + r).setValue(i);
            spreadSheet.getRange(Invoice.itemOrderColumnName + r).setHorizontalAlignment("Center");
            r = r + 1;
            i = i + 1;
        }
        autoResizeRows(spreadSheet, Invoice.itemsBeginRow, r);
        addjustColumnWidth(spreadSheet);
        addjustPayments(spreadSheet, r + 6);
    } else {
        Browser.msgBox('Invalid Context', 'It seems you tried to format a sheet [' + invoiceSheetName + '] NOT an invoice. Please choose another one', Browser.Buttons.OK);
    }
}

function addjustColumnWidth(invoiceSheet) {

    // Set size of columns
    invoiceSheet.setColumnWidth(1, 20);
    invoiceSheet.setColumnWidth(2, 90);
    invoiceSheet.setColumnWidth(3, 270);
    invoiceSheet.setColumnWidth(4, 100);
    invoiceSheet.setColumnWidth(5, 130);
    invoiceSheet.setColumnWidth(6, 40);
    invoiceSheet.setColumnWidth(7, 130);
}
function addjustPayments(invoiceSheet, paymentRow) {
    invoiceSheet.setRowHeight(paymentRow, 100);
    invoiceSheet.setRowHeight(paymentRow + 1, 5);
}

function isInvoiceSheetName(invoiceSheetName: string) {
    return invoiceSheetName != 'Invoice Manager'
        && invoiceSheetName != 'Dinner'
        && invoiceSheetName != 'Items'
        && invoiceSheetName != 'Countries'
        && invoiceSheetName != 'Config'
        && invoiceSheetName != 'Bank Account';
}

function isRealInvoiceSheetName(invoiceSheetName: string) {
    const regex = '^[0-9]{2}.[0-9]{2}.[0-9]{2}_R';
    var found = invoiceSheetName.match(regex);
    return found;
}

function isManagementSheetName(invoiceSheetName: string) {
    return invoiceSheetName == 'Invoice Manager'
        || invoiceSheetName == 'Dinner'
        || invoiceSheetName == 'Items'
        || invoiceSheetName == 'Countries'
        || invoiceSheetName == 'Config'
        || invoiceSheetName == 'Bank Account'
        || invoiceSheetName == 'Invoice';
}

function nextIsTotal(spreadSheet: GoogleAppsScript.Spreadsheet.Sheet, currentRow: number) {
    var nextRow = currentRow + 1;
    var nValue = spreadSheet.getRange(Invoice.itemOrderColumnName + nextRow).getValue();
    return nValue == Invoice.itemsEndRowIndicator;
}

function insertItemAfter(spreadsheet, currentRow) {
    spreadsheet.insertRowAfter(currentRow);
    equipDataFilter(spreadsheet, currentRow + 1);
}

function removeRows(spreadsheet, fromRow, numOfDeletedRow) {
    spreadsheet.deleteRows(fromRow, numOfDeletedRow);
}

function DataFormat() {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange('E13').activate();
    spreadsheet.getActiveRangeList().setNumberFormat('#,##0.00[$ ₫]');
};

function equipDataFilter(spreadsheet, currentRow) {
    var cell = spreadsheet.getRange('C' + currentRow);

    cell.setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(spreadsheet.getRange('Items!$B$2:$B$700'), true)
        .build());

    cell = spreadsheet.getRange('D' + currentRow);
    cell.setFormula('=iferror(VLOOKUP(C' + currentRow + ',ListItems, 3, FALSE))');

    cell = spreadsheet.getRange('E' + currentRow);
    cell.setFormula('=iferror(VLOOKUP(C' + currentRow + ',ListItems, 2, FALSE))');


    cell = spreadsheet.getRange('G' + currentRow);
    cell.setFormula('=D' + currentRow + '*E' + currentRow);
};

//=================================================================
function onOpen(e) {
    config();
    //Create custom menu to export range to Slides.
    SpreadsheetApp.getUi()
        .createMenu('Billing')
        .addItem('Auto Format (Cân chỉnh)', 'formatTheActiveInvoice')
        .addItem('Export Bill (Xuất bill)', 'exportInvoiceAsPDF')
        .addItem('Archive (Dọn dẹp và lưu trữ)', 'cleanAndArchive')
        .addItem('Load Configuration (Load Config)', 'config')
        .addItem('Populate Invoice', 'populateInvoices')
        .addItem('Sign !!', 'signTheBill')
        .addItem('Invoice JSON', 'showInvoiceJSONMessage')
        .addToUi();
    // SpreadsheetApp.getUi()
    //     .createAddonMenu()
    //     .addItem('Export all sheets', 'exportAsPDF')
    //     .addItem('Export all sheets as separate files', 'exportAllSheetsAsSeparatePDFs')
    //     .addItem('Export current sheet', 'exportCurrentSheetAsPDF')
    //     .addItem('Export selected area', 'exportPartAsPDF')
    //     .addItem('Export predefined area', 'exportNamedRangesAsPDF')
    //     .addToUi()



    loadRecentBills();
    listTodayBills();

    formatRecentBillList();
    formatTodayBillList();
    focusManager();
    formatInvoiceManager();
}

function showInvoiceJSONMessage() {
    var spreadSheet = SpreadsheetApp.getActive();
    var sheet = spreadSheet.getActiveSheet();

    // var sheet = sheets[i];
    console.info('Checking sheet [%s]...', sheet.getName());
    var sheetName = sheet.getName();
    if (isRealInvoiceSheetName(sheetName)) {
        // Check if the invoice is older than yesterday then export and remove that invoice
        try {
            console.info('Generate invoice [%s] JSON message...', sheetName);
            var data: object = buildInvoiceJSONMessage(sheet);
            var ui = SpreadsheetApp.getUi(); // Same variations.
            var result = ui.prompt(
                'Invoice',
                JSON.stringify(data),
                ui.ButtonSet.OK);

        } catch (error) {
            console.error(error);
        }
    } else if (!isManagementSheetName(sheetName)) {
        try {
            console.info('Sheet [%s] is not invoice or management, ignore it...', sheetName);
        } catch (error) {
            console.error(error);
        }
    }
}

function getManagementSheet() {
    return SpreadsheetApp.getActive().getSheetByName(InvoiceManager.sheetName);
}

function focusManager() {
    var spreadSheet = SpreadsheetApp.getActive();
    spreadSheet.getSheetByName('Invoice Manager').activate();
}

function loadRecentBills() {
    var spreadSheet = SpreadsheetApp.getActive();

    var counter = 0;
    var i = 0;
    var sheets = spreadSheet.getSheets();

    var recentSheets = [];

    while (counter < 5 && i < sheets.length) {
        var currentSheet = sheets[i];
        if (isRealInvoiceSheetName(currentSheet.getName())) {
            recentSheets.push(currentSheet);
            counter++;
        }
        i++;
    }


    cleanRecentBillList();

    var managementSheet: GoogleAppsScript.Spreadsheet.Sheet = getManagementSheet();
    var range = managementSheet.getRange(InvoiceManager.recentBillRangeName);
    var col = range.getColumn();
    var row = range.getRow();
    var i = 0;
    while (i < recentSheets.length) {
        var value = SheetUtility.buildSheetHyperLink(recentSheets[i])
        managementSheet.getRange(row + i, col).setValue(value);
        managementSheet.getRange(row + i, col - 1).insertCheckboxes();
        i++;
    }
}

function SelectedRangeToImage() {
    var slide = RangeToSlides();
    var slideId = slide.getId();
    var images = [];
    for (var x = 0; x < slide.getSlides().length; x++) {
        var image = SlidesToImage(slide.getName() + x, slideId, slide.getSlides()[x].getObjectId());
        images.push(image);
    }

    //Show interface with links to all images
    var ui = SpreadsheetApp.getUi();
    var html = HtmlService.createHtmlOutput();
    html.append("<p>Your images:</p>");
    html.append("<ul>");
    for (var i = 0; i < images.length; i++) {
        html.append("<li><a href='" + images[i].getUrl() + "'>" + images[i].getName() + "</a></li>");
    }
    html.append("</ul>");
    html.append("<input type='button' value='Close' onclick='google.script.host.close()' />");
    ui.showModalDialog(html, "Exporting results:");
}

function RangeToSlides() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var range = ss.getActiveRange();

    var rangeValues = range.getDisplayValues();
    var rangeHorizontalAlignments = range.getHorizontalAlignments()
    var rangeBackgrounds = range.getBackgrounds();
    var rangeFontWeights = range.getFontWeights();


    var sl = SlidesApp.create("ExportToImage" + new Date());
    var slide = sl.getSlides()[0];

    //Create table with size of the range
    var table = slide.insertTable(rangeValues.length, rangeValues[0].length);
    for (var x = 0; x < rangeValues.length; x++) {
        for (var y = 0; y < rangeValues[x].length; y++) {
            var cell = table.getCell(x, y);
            cell.getText().setText(rangeValues[x][y]); //Set text
            cell.getFill().setSolidFill(rangeBackgrounds[x][y]); //Set background
            cell.getText().getTextStyle().setBold(rangeFontWeights[x][y] == "bold" ? true : false); //Set text formatting
            var alignment;
            switch (rangeHorizontalAlignments[x][y]) {
                case "general-left":
                    alignment = SlidesApp.ParagraphAlignment.START;
                    break;
                case "general-right":
                    alignment = SlidesApp.ParagraphAlignment.END;
                    break;
                case "center":
                    alignment = SlidesApp.ParagraphAlignment.CENTER;
                    break;
            }
            cell.getText().getParagraphStyle().setParagraphAlignment(alignment); //Set text alignment
        }
    }
    sl.saveAndClose();
    return sl;
}

function SlidesToImage(name, presentationId, slideId) {
    var url = "https://slides.googleapis.com/v1/presentations/" + presentationId + "/pages/" + slideId + "/thumbnail";
    var options = {
        headers: {
            Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        }
    };
    var response = UrlFetchApp.fetch(url, options);
    var responseJson = JSON.parse(response.getContentText());
    var imageurl = responseJson.contentUrl;
    var imageResponse = UrlFetchApp.fetch(imageurl, options);
    var blob = imageResponse.getBlob();
    blob.setName(name);
    var resultingFile = DriveApp.createFile(blob);
    return resultingFile;
}



function exportPdf() {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange('I18').activate();
    spreadsheet.getActiveSheet().getRowGroup(1, 1).expand();
};

//===================================  EXPORT PDF =======================================================//
function generatePdf() {
    // Get active spreadsheet.
    var sourceSpreadsheet = SpreadsheetApp.getActive();

    // Get active sheet.
    var sheets = sourceSpreadsheet.getSheets();
    var sheetName = sourceSpreadsheet.getActiveSheet().getName();
    var sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);

    // Set the output filename as SheetName.
    var pdfName = sheetName;

    // Get folder containing spreadsheet to save pdf in.
    var parents = DriveApp.getFileById(sourceSpreadsheet.getId()).getParents();
    if (parents.hasNext()) {
        var folder = parents.next();
    }
    else {
        folder = DriveApp.getRootFolder();
    }

    // Copy whole spreadsheet.
    var destSpreadsheet = SpreadsheetApp.open(DriveApp.getFileById(sourceSpreadsheet.getId()).makeCopy("tmp_convert_to_pdf", folder))

    // Delete redundant sheets.
    var sheets = destSpreadsheet.getSheets();
    for (i = 0; i < sheets.length; i++) {
        if (sheets[i].getSheetName() != sheetName) {
            destSpreadsheet.deleteSheet(sheets[i]);
        }
    }

    var destSheet = destSpreadsheet.getSheets()[0];

    // Repace cell values with text (to avoid broken references).
    var sourceRange = sourceSheet.getRange(1, 1, sourceSheet.getMaxRows(), sourceSheet.getMaxColumns());
    var sourcevalues = sourceRange.getValues();
    var destRange = destSheet.getRange(1, 1, destSheet.getMaxRows(), destSheet.getMaxColumns());
    destRange.setValues(sourcevalues);

    // Save to pdf.
    var theBlob = destSpreadsheet.getBlob().getAs('application/pdf').setName(pdfName);
    var newFile = folder.createFile(theBlob);

    // Delete the temporary sheet.
    DriveApp.getFileById(destSpreadsheet.getId()).setTrashed(true);
}

function autoResizeRows(spreadSheet, fromRow, toRow) {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange(fromRow + ':' + toRow).activate();
    var numRows = toRow - fromRow + 1;
    spreadsheet.getActiveSheet().autoResizeRows(fromRow, numRows);
};

function createInvoiceFromTemplate() {
    var spreadsheet = SpreadsheetApp.getActive();
    var templateSheet = spreadsheet.getSheetByName('Invoice');
    var managementSheet = getManagementSheet();

    var invoiceSheetName = managementSheet.getRange(InvoiceManager.currentInvoiceNameRange).getValue();
    // TODO -- Need to check if the sheet exist and open that one instead

    var invoiceSheet = spreadsheet.getSheetByName(invoiceSheetName);
    if (!invoiceSheet) {
        var invoiceSheetIndex = managementSheet.getIndex() + 1;
        invoiceSheet = spreadsheet.insertSheet(invoiceSheetIndex);
        invoiceSheet.setName(invoiceSheetName);
        invoiceSheet.setHiddenGridlines(true);
        templateSheet.getRange("B:G").copyTo(invoiceSheet.getRange('B:B'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

        invoiceSheet.getRange("2:11").shiftColumnGroupDepth(1);
        invoiceSheet.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
        invoiceSheet.collapseAllRowGroups();

        invoiceSheet.getRange('D7').copyTo(invoiceSheet.getRange('D7'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
        invoiceSheet.getRange('G7').copyTo(invoiceSheet.getRange('G7'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
        invoiceSheet.getRange('D9').copyTo(invoiceSheet.getRange('D9'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

        // Set size of columns
        addjustColumnWidth(invoiceSheet);
        addjustPayments(invoiceSheet, 22);

        // Add invoice to recent list
        loadRecentBills();
        listTodayBills();
        formatRecentBillList();
        formatTodayBillList();
    }
    invoiceSheet.activate();
}

function copyToNewSheet() {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange('B13:G19').activate();
    var newInvoice = SpreadsheetApp.create("TEST");
    //  var file = DriveApp.getFilesByName("TEST");
    var tempSheet = SpreadsheetApp.openById(newInvoice.getId()).getSheets()[0];
    tempSheet.setName('ABC');
    tempSheet.setHiddenGridlines(true);
    spreadsheet.getRange('Invoice!B13:G19').copyTo(tempSheet.getRange('B2'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    tempSheet.getRange('C2:C5').clearDataValidations();
    tempSheet.getRange('C').activate();
    tempSheet.autoResizeColumns(3, 1);

};

// function generateInvoice() {
//     var spreadsheet = SpreadsheetApp.getActive();
//     var folder = getFolderOfFile(spreadsheet);
//     var destSpreadsheet = SpreadsheetApp.open(DriveApp.getFileById(spreadsheet.getId()).makeCopy("INVOICE", folder));

//     // Delete redundant sheets.
//     var sheets = destSpreadsheet.getSheets();
//     for (var i = 0; i < sheets.length; i++) {
//         if (sheets[i].getSheetName() != 'Invoice' && sheets[i].getSheetName() != 'Items') {
//             destSpreadsheet.deleteSheet(sheets[i]);
//         }
//     }

//     var invoiceSheet = destSpreadsheet.getSheets()[0];
//     // Clear inneccessary data
//     invoiceSheet.getRange('C13:C15').clearDataValidations();
//     invoiceSheet.setFrozenRows(0);
//     invoiceSheet.getRowGroup(1, 1).remove();
//     invoiceSheet.getRowGroup(4, 1).remove();
// }

function getFolderOfFile(f) {
    var parents = DriveApp.getFileById(f.getId()).getParents();
    if (parents.hasNext()) {
        var folder = parents.next();
    }
    else {
        folder = DriveApp.getRootFolder();
    }

    return folder;
}

function hideGridView() {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange('H6').activate();
    spreadsheet.getActiveSheet().setHiddenGridlines(true);
};