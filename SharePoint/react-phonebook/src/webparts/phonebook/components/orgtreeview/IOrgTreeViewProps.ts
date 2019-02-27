import { IOrgUnit } from "../../../../classes/IOrgUnit";
import { Guid } from "@microsoft/sp-core-library";

export interface IOrgTreeViewProps {
    getData: () => Promise<IOrgUnit[]>;
    orgUnitSelected: (orgUnitId: Guid) => void;
    orgStructureTermSet: string;
}