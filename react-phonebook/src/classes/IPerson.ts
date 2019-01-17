export interface IPerson{
    fullName: string;
    email: string;
    phone: string;
}

export class Person implements IPerson {
    public fullName: string;    
    public email: string;
    public phone: string;
}