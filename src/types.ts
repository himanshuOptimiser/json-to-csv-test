export interface Address {
    street: string;
    city: string;
    country: string;
}

export interface Order {
    orderId: string;
    product: string;
    quantity: number;
    price: number;
}

export interface Customer {
    id: number;
    name: string;
    email: string;
    address: Address;
    orders: Order[];
}
