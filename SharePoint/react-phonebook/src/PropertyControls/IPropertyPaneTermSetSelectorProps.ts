import { Guid } from "@microsoft/sp-core-library";

export interface IPropertyPaneTermSetSelectorProps {
    label: string;
    onPropertyChange: (propertyPath: string, termSetId: Guid, groupId: Guid) => void;
    currentTermSetId: string;
}