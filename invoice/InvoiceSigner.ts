function signTheBill() {
    var spreadSheet = SpreadsheetApp.getActive();
    var spreadSheetName = spreadSheet.getName();

    var sheet = spreadSheet.getActiveSheet();
    var invoiceSheetName = sheet.getName();

    var signatureExistingIndicator = 'signed';

    if (isInvoiceSheetName(invoiceSheetName)) {

        if (sheet.getRange(Invoice.signatureExistingIndicatorRangeName).getValue() != signatureExistingIndicator) {
            var signatureRow = SheetUtility.getTotalAmountRow(sheet) + 8;

            // var response = UrlFetchApp.fetch(
            //     'https://upload.wikimedia.org/wikipedia/commons/thumb/7/77/Trevor_Noah_signature.svg/1200px-Trevor_Noah_signature.svg.png');
            // var binaryData = response.getContent();
            // var blob = Utilities.newBlob(binaryData, 'image/png', 'MyImageName');

            var blob = DriveApp.getFileById('1xCe3h_U-RJybhMIesyJkLkm6aGzo3qwJ').getBlob();

            var image:GoogleAppsScript.Spreadsheet.OverGridImage = sheet.insertImage(blob, Invoice.signatureColumnNum, signatureRow);
            // console.log(image.getWidth());
            // image.setWidth(250);
            console.log(image.getWidth());
            console.log(image.getHeight());
            sheet.getRange(Invoice.signatureExistingIndicatorRangeName).setValue(signatureExistingIndicator);
        }
    }
}