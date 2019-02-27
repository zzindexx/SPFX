import { IOrgUnit } from "../../../../classes/IOrgUnit";
import { Guid } from "@microsoft/sp-core-library";

export interface IOrgTreeViewState {
    isLoading: boolean;
    units: IOrgUnit[];
    selectedUnitId: Guid;
}