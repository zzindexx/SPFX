import { Person, IPerson } from "../../../../classes/IPerson";
import { Guid } from "@microsoft/sp-core-library";

export interface IPeopleViewProps {
    getData: (query: string, orgUnitId: Guid, page:number) => Promise<{persons: IPerson[], numberOfResults: number}>;
    orgUnitId: Guid;
    pageSize: number;
}