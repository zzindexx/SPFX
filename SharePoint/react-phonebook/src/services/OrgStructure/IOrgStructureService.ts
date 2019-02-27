import { IOrgUnit } from "../../classes/IOrgUnit";

export interface IOrgStructureService {
 getFullTree(groupId: string, termSetId: string): Promise<IOrgUnit[]>;
};