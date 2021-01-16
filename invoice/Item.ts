class Item {
    private itemOrder: number;
    public get getItemOrder(): number {
        return this.itemOrder;
    }
    public set setItemOrder(value: number) {
        this.itemOrder = value;
    }

    private itemUnitPrice: number;
    public get getItemUnitPrice(): number {
        return this.itemUnitPrice;
    }
    public set setItemUnitPrice(value: number) {
        this.itemUnitPrice = value;
    }

    private itemQuantity: number;
    public get getItemQuantity(): number {
        return this.itemQuantity;
    }
    public set setItemQuantity(value: number) {
        this.itemQuantity = value;
    }
    private itemName: string;
    public get getItemName(): string {
        return this.itemName;
    }
    public set setItemName(value: string) {
        this.itemName = value;
    }

    private itemAmount: number;
    public get getItemAmount(): number {
        return this.itemAmount;
    }
    public set setItemAmount(value: number) {
        this.itemAmount = value;
    }

    constructor(itemOrder: number, itemName: string, itemUnitPrice: number, itemQuantity: number, itemAmount: number) {
        this.itemOrder = itemOrder;
        this.itemName = itemName;
        this.itemUnitPrice = itemUnitPrice;
        this.itemQuantity = itemQuantity;
        this.itemAmount = itemAmount;
    }
}