export interface IOrgUnit {
    title: string;
    child_orgunits: IOrgUnit[];
}

export class OrgUnit implements IOrgUnit {
    public title: string;    
    public child_orgunits: IOrgUnit[];
}