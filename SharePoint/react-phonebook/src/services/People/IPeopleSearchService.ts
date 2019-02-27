import { IPerson } from "../../classes/IPerson";
import { Guid } from "@microsoft/sp-core-library";

export interface IPeopleSearchService{
    getPeople(searchQuery: string, orgUnitId: Guid, page: number, pageSize: number): Promise<{persons: IPerson[], numberOfResults: number}> 
}