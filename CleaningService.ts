function cleanAndArchive() {

    var spreadSheet = SpreadsheetApp.getActive();

    var sheets = spreadSheet.getSheets();
    var i: number = 0;

    sheets.forEach(function (sheet) {
        // var sheet = sheets[i];
        console.info('Check sheet [%s]', sheet.getName());
        var sheetName = sheet.getName();
        if (isRealInvoiceSheetName(sheetName)) {
            // Check if the invoice is older than yesterday then export and remove that invoice
            if (needTobeArchived(sheetName)) {
                try {
                    console.info('Archiving invoice [%s]...', sheetName);
                    exportSheetToPDF(spreadSheet, sheet, sheetName);
                    spreadSheet.deleteSheet(sheet);
                } catch (error) {
                    console.error(error);
                }
            }
        } else if (!isManagementSheetName(sheetName)) {
            try {
                console.info('Sheet [%s] is not invoice or management, it will be removed...', sheetName);
                spreadSheet.deleteSheet(sheet);
            } catch (error) {
                console.error(error);
            }
        }
    });
}

function needTobeArchived(invoiceName: string): boolean {
    var invoiceDate = DateUtility.fromInvoiceName(invoiceName);
    var today = new Date();
    var delta = today.getTime() - invoiceDate.getTime();
    return delta > 86400000;
}
