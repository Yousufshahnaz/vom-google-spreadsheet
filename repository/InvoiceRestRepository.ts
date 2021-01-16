function populateInvoices() {
    var spreadSheet = SpreadsheetApp.getActive();

    var sheets = spreadSheet.getSheets();
    var i: number = 0;

    sheets.forEach(function (sheet) {
        // var sheet = sheets[i];
        console.info('Check sheet [%s]', sheet.getName());
        var sheetName = sheet.getName();
        if (isRealInvoiceSheetName(sheetName)) {
            // Check if the invoice is older than yesterday then export and remove that invoice
            try {
                console.info('Populate invoice [%s]...', sheetName);
                writeSheet(sheet);
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
    });
}

function writeSheet(invoiceSheet: GoogleAppsScript.Spreadsheet.Sheet) {

    var data: object = buildInvoiceJSONMessage(invoiceSheet);
    var url = 'https://vom-homestay.herokuapp.com/vom/invoice';

    var options: URLFetchRequestOptions = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify(data)
    };

    console.log('Write invoice [%s] to endpoint [%s]...', invoiceSheet.getName(), url);
    var response = UrlFetchApp.fetch(url, options);
}


// // Make a POST request with form data.
// var resumeBlob = Utilities.newBlob('Hire me!', 'text/plain', 'resume.txt');
// var formData = {
//   'name': 'Bob Smith',
//   'email': 'bob@example.com',
//   'resume': resumeBlob
// };
// // Because payload is a JavaScript object, it is interpreted as
// // as form data. (No need to specify contentType; it automatically
// // defaults to either 'application/x-www-form-urlencoded'
// // or 'multipart/form-data')
// var options = {
//   'method' : 'post',
//   'payload' : formData
// };
// UrlFetchApp.fetch('https://httpbin.org/post', options);
// // call the API
// var response = UrlFetchApp.fetch(url, params);
// var data = response.getContentText();
// var json = JSON.parse(data);


function buildInvoiceJSONMessage(invoiceSheet: GoogleAppsScript.Spreadsheet.Sheet): object {
    var invoiceName = invoiceSheet.getName();
    var invoiceId: string = invoiceName;

    var guestName: string = invoiceSheet.getRange(Invoice.guestNameRangeName).getValue();
    var country: string = invoiceSheet.getRange(Invoice.countryRangeName).getValue();
    var checkInDate: string = invoiceSheet.getRange(Invoice.checkInRangeName).getValue();
    var checkOutDate: string = invoiceSheet.getRange(Invoice.checkOutRangeName).getValue();
    var link: string = invoiceSheet.getRange(Invoice.linkRangeName).getValue();

    var items: Item[] = [];
    var r = Invoice.itemsBeginRow;
    while (!nextIsTotal(invoiceSheet, r)) {
        var itemName: string = invoiceSheet.getRange(Invoice.itemNameColumnName + r).getValue();
        if (itemName != null && itemName != '') {
            console.log('Collecting item [%s]...', itemName);
            var itemOrder: number = invoiceSheet.getRange(Invoice.itemOrderColumnName + r).getValue();
            var itemQuantity: number = invoiceSheet.getRange(Invoice.itemQuantityColumnName + r).getValue();
            var itemUnitPrice: number = invoiceSheet.getRange(Invoice.itemUnitPriceColumnName + r).getValue();
            var itemAmount: number = invoiceSheet.getRange(Invoice.itemAmountColumnName + r).getValue();
            items.push(new Item(itemOrder, itemName, itemUnitPrice, itemQuantity, itemAmount));
        }
        r = r + 1;
    }
    r = r + 1;
    var totalAmount: number = invoiceSheet.getRange(Invoice.itemAmountColumnName + r).getValue();
    r = r + 1;
    var paidAmount: number = invoiceSheet.getRange(Invoice.itemAmountColumnName + r).getValue();
    r = r + 1;
    var remainAmount: number = invoiceSheet.getRange(Invoice.itemAmountColumnName + r).getValue();


    var data = {
        'invoiceId': invoiceId,
        'sheetName': invoiceName,
        'name': guestName,
        'country': country,
        'checkInDate': checkInDate,
        'checkOutDate': checkOutDate,
        'items': items,
        'totalAmount': totalAmount,
        'paidAmount': paidAmount,
        'remainAmount': remainAmount,
        'link': link
    };

    return data;
}