export interface IPerson {
    fullName: string;
    email: string;
    workPhone?: string;
    mobilePhone?: string;
    jobTitle?: string;
    department?: string;
    pictureUrl?: string;
    accountUrl: string;
}

export class Person implements IPerson {
    public fullName: string;
    public email: string;
    public workPhone?: string;
    public mobilePhone?: string;
    public jobTitle?: string;
    public department: string;
    public pictureUrl?: string;
    public accountUrl: string;
}