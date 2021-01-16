class InvoiceManager {
    static sheetName: string = 'Invoice Manager';

    static listRecentBillColumn: string = 'F';
    static listRecentBillColumnWidth: number = 300;

    static listTodayBillColumn: string = 'F';
    static listTodayBillColumnWidth: number = 400;

    static listBillFromRow: number = 5;
    static listBillToRow: number = 10;

    static todayBillRangeName: string = 'F5:F10';
    static recentBillRangeName: string = 'F13:F18';

    static checkInDateRangeName: string = 'C8';
    static checkOutDateRangeName: string = 'C9';
    static guestNameRangeName: string = 'C6';
    static countryRangeName: string = 'C7';
    static roomsRangeRow: number = 10;
    static roomsRangeColumnName: string = 'C';

    static editButtonInvoiceColumn: number = 5;
    static currentInvoiceIdRange = 'C17';
    static currentInvoiceNameRange = 'C16';

    static invoiceNameFomular: string = '=IFNA(CONCATENATE(TEXT(C9, "dd.mm.yy"),"_",SUBSTITUTE(CONCATENATE(FILTER(B10:C15,C10:C15 = true)),"TRUE","_"),C6),"")';


    constructor(parameters) {

    }
}

// function addRecentInvoices(sheet: GoogleAppsScript.Spreadsheet.Sheet) {

//     var managementSheet = getManagementSheet();

//     var recentBillRange: GoogleAppsScript.Spreadsheet.Range = managementSheet.getRange(InvoiceManager.recentBillRangeName);
//     var column: number = recentBillRange.getColumn();

//     var i = InvoiceManager.listBillToRow;
//     while (i > InvoiceManager.listBillFromRow) {
//         var aboveRow = i - 1;
//         managementSheet.getRange(column, aboveRow)
//             .copyTo(managementSheet.getRange(column, i), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
//         managementSheet.getRange(column - 1, aboveRow)
//             .copyTo(managementSheet.getRange(column - 1, i), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
//         i--;
//     }
//     var linkValue = buildSheetHyperLink(sheet);
//     managementSheet.getRange(column, i).setValue(linkValue);
//     managementSheet.getRange(column - 1, i).insertCheckboxes();
// }

function cleanRecentBillList() {
    var managementSheet: GoogleAppsScript.Spreadsheet.Sheet = getManagementSheet();
    var range: GoogleAppsScript.Spreadsheet.Range = managementSheet.getRange(InvoiceManager.recentBillRangeName);
    managementSheet.getRange(range.getRow()
        , range.getColumn() - 1
        , range.getNumRows()
        , range.getNumColumns() + 1).clear();
}

function cleanTodayBillList() {
    var managementSheet: GoogleAppsScript.Spreadsheet.Sheet = getManagementSheet();
    var range: GoogleAppsScript.Spreadsheet.Range = managementSheet.getRange(InvoiceManager.todayBillRangeName);
    managementSheet.getRange(range.getRow()
        , range.getColumn() - 1
        , range.getNumRows()
        , range.getNumColumns() + 1).clear();
}

function renderSelectedInvoice() {
    var managementSheet = getManagementSheet();
    var range = managementSheet.getRange(InvoiceManager.todayBillRangeName);
    var col = range.getColumn();
    var row = range.getRow();
    var numRows = range.getNumRows();
    var i: number = row;
    while (i <= i + numRows && !managementSheet.getRange(i, col - 1).isChecked()) {
        i++;
    }
    console.info('Last checked row [%d] and its check status is [%s]', i, managementSheet.getRange(i, col - 1).isChecked());
    var foundSheetName = managementSheet.getRange(i, col).getValue();
    var foundSheet = SpreadsheetApp.getActive().getSheetByName(foundSheetName);
    if (foundSheet) {
        console.info('Rendering invoice [%s]...', foundSheet.getName());
        renderInvoiceBySheet(foundSheet);
        managementSheet.getRange(InvoiceManager.currentInvoiceIdRange).setValue(managementSheet.getRange(i, col + 1).getValue());
    } else {
        console.info('Could not find any sheet with name [%s]', managementSheet.getRange(i, col).getValue());
    }
}

function renderInvoiceBySheet(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const sheetName = sheet.getName();
    const checkOutDate = SheetUtility.getCheckOutdate(sheetName);

    var checkInDate = SheetUtility.getCheckOutdate(sheet.getRange(InvoiceTemplate.checkInRangeName).getValue());

    var guestName: string = sheet.getRange(InvoiceTemplate.guestNameRangeName).getValue();
    var country: string = sheet.getRange(InvoiceTemplate.countryRangeName).getValue();

    var nameParts: string[] = SheetUtility.extractSheetName(sheetName);
    console.info(nameParts.toString());
    var rooms: boolean[] = [false, false, false, false, false, false];

    var invoice: Invoice = new Invoice(checkInDate, checkOutDate, country, rooms);
    invoice.guestName = nameParts[nameParts.length - 1];
    var i = 1;
    while (i < nameParts.length - 1) {
        var roomIndex = +nameParts[i].substr(1);
        invoice.rooms[roomIndex - 1] = true;
        i++;
    }
    console.info('Guest [%s] from [%s] check in on [%s] check out on [%s] rooms [%s]', guestName, country, checkInDate.toString(), checkOutDate.toString()
        , invoice.rooms.toString());

    renderInvoiceManager(invoice);
}

function renderInvoiceManager(invoice: Invoice) {
    var managerSheet: GoogleAppsScript.Spreadsheet.Sheet = getManagementSheet()

    managerSheet.getRange(InvoiceManager.checkInDateRangeName).setValue(DateUtility.formatDate(invoice.checkInDate));
    managerSheet.getRange(InvoiceManager.checkOutDateRangeName).setValue(DateUtility.formatDate(invoice.checkOutDate));
    managerSheet.getRange(InvoiceManager.guestNameRangeName).setValue(invoice.guestName);
    managerSheet.getRange(InvoiceManager.countryRangeName).setValue(invoice.country);
    invoice.rooms.forEach((r, i) => {
        managerSheet.getRange(InvoiceManager.roomsRangeColumnName + (InvoiceManager.roomsRangeRow + i)).setValue(r);
    });
}

function editExistingInvoice() {
    var managementSheet = getManagementSheet();
    var currentInvoiceId: number = managementSheet.getRange(InvoiceManager.currentInvoiceIdRange).getValue();
    var sheets: GoogleAppsScript.Spreadsheet.Sheet[] = SpreadsheetApp.getActive().getSheets();
    var foundSheet: GoogleAppsScript.Spreadsheet.Sheet;

    for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getSheetId() == currentInvoiceId) {
            foundSheet = sheets[i];
            break;
        }
    }

    if (foundSheet) {
        foundSheet.setName(managementSheet.getRange(InvoiceManager.currentInvoiceNameRange).getValue());
        foundSheet.getRange(Invoice.checkOutRangeName).setValue(managementSheet.getRange(InvoiceManager.checkOutDateRangeName).getValue());
        foundSheet.activate();
    } else {
        if ('YES' == Browser.msgBox('Info', 'Can not find any sheet with id [' + currentInvoiceId + ']. Do you want to create a new one?', Browser.Buttons.YES_NO)) {
            createInvoiceFromTemplate();
        }
    }
}

function insertImageIntoCell() {

}

function formatInvoiceManager() {
    var managementSheet = getManagementSheet();
    managementSheet.setColumnWidth(InvoiceManager.editButtonInvoiceColumn, 200);
}