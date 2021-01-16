class SheetUtility {

    static getTotalAmountRow(sheet: GoogleAppsScript.Spreadsheet.Sheet): number {
        var totalAmountRow: number = Invoice.itemsBeginRow;
        while (sheet.getRange(Invoice.itemOrderColumnName + totalAmountRow).getValue() != Invoice.itemsEndRowIndicator && totalAmountRow < 200) {
            totalAmountRow++;
        }

        return totalAmountRow;
    }

    static getCheckOutdate(sheetName: string): Date {
        if (String) {
            var dateString: String = sheetName.substr(0, '13.01.20'.length);
            return DateUtility.fromString(sheetName);
        }
        return;
    }

    static extractSheetName(sheetName: string): string[] {
        let sheetNameRegex = /([0-9]{2}.[0-9]{2}.[0-9]{2})_([R1-6_]+)_([A-Za-z ]+)/;

        var tempGroups: string[] = sheetName.match(sheetNameRegex);
        var groups: string[] = [];
        groups.push(tempGroups[1]);
        tempGroups[2].split('_').forEach(i => groups.push(i));
        groups.push(tempGroups[3]);

        return groups;
    }

    static buildSheetHyperLink(sheet: GoogleAppsScript.Spreadsheet.Sheet): string {
        return '=HYPERLINK("#gid=' + sheet.getSheetId() + '","' + sheet.getName() + '")';
    }
}