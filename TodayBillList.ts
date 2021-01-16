function listTodayBills() {
    var date = new Date();
    date.setHours(0, 0, 0, 0);
    listBillOfDate(date);
}

// function listBillOfDate(date: GoogleAppsScript.Base.Date) {
//     var dateString = Utilities.formatDate(date, "GMT+7", "dd.MM.yy");
//     console.info('List the created bills of since [%s]', date.toString());
//     var spreadsheet = SpreadsheetApp.getActive();

//     var sheets = spreadsheet.getSheets();
//     let todayInvoiceSheets: string[] = [];
//     for (var i = 0; i < sheets.length; i++) {
//         var sheet = sheets[i];
//         if (isRealInvoiceSheetName(sheet.getName())) {
//             if (isTodayInvoice(sheet)) {
//                 todayInvoiceSheets.push(sheet.getName());
//             }
//         }
//     }

//     cleanTodayBillList();
//     todayInvoiceSheets.sort((a, b) => a.localeCompare(b));
//     var managementSheet: GoogleAppsScript.Spreadsheet.Sheet = getManagementSheet();
//     var range = managementSheet.getRange(InvoiceManager.todayBillRangeName);
//     var col = range.getColumn();
//     var row = range.getRow();
//     var i = 0;
//     while (i < todayInvoiceSheets.length) {
//         managementSheet.getRange(row + i, col).setValue(todayInvoiceSheets[i]);
//         managementSheet.getRange(row + i, col - 1).insertCheckboxes();
//         i++;
//     }
// }

function listBillOfDate(date: GoogleAppsScript.Base.Date) {
    var dateString = Utilities.formatDate(date, "GMT+7", "dd.MM.yy");
    console.info('List the created bills of since [%s]', date.toString());
    var spreadsheet = SpreadsheetApp.getActive();

    var sheets = spreadsheet.getSheets();
    let todayInvoiceSheets: GoogleAppsScript.Spreadsheet.Sheet[] = [];
    for (var i = 0; i < sheets.length; i++) {
        var sheet = sheets[i];
        if (isRealInvoiceSheetName(sheet.getName())) {
            if (isTodayInvoice(sheet)) {
                todayInvoiceSheets.push(sheet);
            }
        }
    }

    cleanTodayBillList();
    todayInvoiceSheets.sort((a, b) => a.getName().localeCompare(b.getName()));
    var managementSheet: GoogleAppsScript.Spreadsheet.Sheet = getManagementSheet();
    var range = managementSheet.getRange(InvoiceManager.todayBillRangeName);
    var col = range.getColumn();
    var row = range.getRow();
    var i = 0;
    while (i < todayInvoiceSheets.length) {
        managementSheet.getRange(row + i, col).setValue(todayInvoiceSheets[i].getName());
        managementSheet.getRange(row + i, col + 1).setValue(todayInvoiceSheets[i].getSheetId());
        managementSheet.getRange(row + i, col - 1).insertCheckboxes();
        i++;
    }
}

function isTodayInvoice(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    var sheetName = sheet.getName();

    var parts: string[] = sheetName.substring(0, 8).split('.');
    var invoiceDate = new Date(+(parts[2] + '20'), +parts[1] - 1, +parts[0]);

    //  var parts ='2014-04-03'.split('-');
    //// Please pay attention to the month (parts[1]); JavaScript counts months from 0:
    //// January - 0, February - 1, etc.
    //var mydate = new Date(parts[0], parts[1] - 1, parts[2]); 


    var today = new Date();
    today.setHours(0); today.setMinutes(0); today.setSeconds(0); today.setMilliseconds(0);

    var delta: number = invoiceDate.getTime() - today.getTime();
    console.log('The invoice date is [%s]. Delta is [%d] milliseconds', invoiceDate, delta);
    var result = (delta < 86400000) && (delta >= -86400000);
    if (!result) {
        console.info('Sheet [%s] is not an invoice of today', sheetName);
    }
    return result;
}
