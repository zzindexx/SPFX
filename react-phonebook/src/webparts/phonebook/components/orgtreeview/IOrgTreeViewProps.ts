import { IOrgUnit } from "../../../../classes/IOrgUnit";

export interface IOrgTreeViewProps {
    getData: () => Promise<IOrgUnit[]>;
}