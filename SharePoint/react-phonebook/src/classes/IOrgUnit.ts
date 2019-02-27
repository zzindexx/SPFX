import { Guid } from '@microsoft/sp-core-library';

export interface IOrgUnit {
    id: Guid;
    title: string;
    childOrgunits: IOrgUnit[];
    parentName: string;
}

export class OrgUnit implements IOrgUnit {
    public id: Guid;
    public title: string;
    public childOrgunits: IOrgUnit[];
    public parentName: string;
}