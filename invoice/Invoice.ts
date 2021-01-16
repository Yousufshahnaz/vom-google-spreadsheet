class Invoice {

    static guestNameRangeName = 'D9';
    static countryRangeName = 'D10';
    static checkInRangeName = 'D7';
    static checkOutRangeName = 'G7';
    static linkRangeName = 'E10';

    static itemOrderColumnName = 'B';
    static itemNameColumnName = 'C';
    static itemQuantityColumnName = 'D';
    static itemUnitPriceColumnName = 'E';
    static itemAmountColumnName = 'G';

    static itemsBeginRow: number = 13;
    static itemsEndRowIndicator = 'Total';

    static signatureColumnNum = 2;
    static signatureExistingIndicatorRangeName = 'E11';

    private _checkInDate: Date;
    public get checkInDate(): Date {
        return this._checkInDate;
    }
    public set checkInDate(value: Date) {
        this._checkInDate = value;
    }
    private _checkOutDate: Date;
    public get checkOutDate(): Date {
        return this._checkOutDate;
    }
    public set checkOutDate(value: Date) {
        this._checkOutDate = value;
    }
    private _country: string;
    public get country(): string {
        return this._country;
    }
    public set country(value: string) {
        this._country = value;
    }
    private _guestName: string;
    public get guestName(): string {
        return this._guestName;
    }
    public set guestName(value: string) {
        this._guestName = value;
    }
    private _rooms: boolean[];
    public get rooms(): boolean[] {
        return this._rooms;
    }
    public set rooms(value: boolean[]) {
        this._rooms = value;
    }
    constructor(checkInDate: Date, checkOutDate: Date, country: string, rooms: boolean[]) {
        this._checkInDate = checkInDate;
        this._checkOutDate = checkOutDate;
        this._country = country;
        this._rooms = rooms;
    }


}