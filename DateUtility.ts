class DateUtility {

    // Input sample: "15.02.20_R1_Tanja Vivien Gangl"
    static fromInvoiceName(dateString: string): Date {
        var parts: string[] = dateString.substring(0, 8).split('.');
        return new Date(+'20'.concat(parts[2]), +parts[1] - 1, +parts[0]);
    }

    static formatDate(date: Date): string {
        return Utilities.formatDate(date, "GMT+7", "dd.MM.yy");
    }

    static invoiceFolderFromInvoiceName(invoiceName: string): string {
        var parts: string[] = invoiceName.substring(0, 8).split('.');
        return "Bill " + parts[1] + ".20" + parts[2];
    }
}