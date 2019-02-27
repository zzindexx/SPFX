import { Guid } from "@microsoft/sp-core-library";

export interface ITermsetSelectorProps {
    label: string;
    onChanged: (termSetId: Guid, groupId: Guid) => void;
    currentTermSetId: string;
}

export interface ITermsetSelectorState {
    isLoading: boolean;
    stores: ITermStore[];
    selectedId: Guid;
}

export interface ITermSet {
    id: Guid;
    title: string;
}

export interface ITermGroup {
    id: Guid;
    title: string;
    termSets: ITermSet[];
}

export interface ITermStore {
    id: Guid;
    title: string;
    termGroups: ITermGroup[];
}