import { IPerson } from "../../../classes/IPerson";
import { IOrgUnit } from "../../../classes/IOrgUnit";
import { Guid } from "@microsoft/sp-core-library";

export interface IPhonebookProps {
  getPeople(query:string, orgUnitId:Guid, page:number, pageSize: number): Promise<{persons: IPerson[], numberOfResults: number}>;
  getOrgUnits: Promise<IOrgUnit[]>;
  pageSize: number;
  orgStructureTermSetId: string;
}

export interface IPhonebookState {
  orgUnitId: Guid;
}