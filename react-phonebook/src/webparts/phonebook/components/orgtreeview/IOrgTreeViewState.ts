import { IOrgUnit } from "../../../../classes/IOrgUnit";

export interface IOrgTreeViewState {
    isLoading: boolean;
    units: IOrgUnit[];
}