import { IPerson } from "../../../classes/IPerson";
import { IOrgUnit } from "../../../classes/IOrgUnit";

export interface IPhonebookProps {
  getPeople: Promise<IPerson[]>;
  getOrgUnits: Promise<IOrgUnit[]>;
}
