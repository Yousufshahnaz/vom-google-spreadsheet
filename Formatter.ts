function formatRecentBillList() {
    var managementSheet = getManagementSheet();

    var range = managementSheet.getRange(InvoiceManager.recentBillRangeName);

    managementSheet.setColumnWidth(range.getColumn(), InvoiceManager.listRecentBillColumnWidth);
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
    range.setFontFamily('Verdana');

    managementSheet.getRange(range.getRow(), range.getColumn() - 1, range.getNumRows(), 1)
        .setHorizontalAlignment("right");
}


function formatTodayBillList() {
    var managementSheet = getManagementSheet();

    var range = managementSheet.getRange(InvoiceManager.todayBillRangeName);

    managementSheet.setColumnWidth(range.getColumn(), InvoiceManager.listTodayBillColumnWidth);
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
    range.setFontFamily('Verdana');

    managementSheet.getRange(range.getRow(), range.getColumn() - 1, range.getNumRows(), 1)
        .setHorizontalAlignment("right");
}