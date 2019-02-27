import * as React from 'react';
import * as ReactDom from 'react-dom';

import { IPropertyPaneField, PropertyPaneFieldType } from "@microsoft/sp-webpart-base";
import { IPropertyPaneTermSetSelectorProps } from "./IPropertyPaneTermSetSelectorProps";
import { IPropertyPaneTermSetSelectorInternalProps } from "./IPropertyPaneTermSetSelectorInternalProps";
import TermsetSelector from './TermsetSelector/TermsetSelector';
import { ITermsetSelectorProps } from './TermsetSelector/ITermsetSelectorProps';
import { Guid } from '@microsoft/sp-core-library';

export class PropertyPaneTermSetSelector implements IPropertyPaneField<IPropertyPaneTermSetSelectorProps> {
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public properties: IPropertyPaneTermSetSelectorInternalProps;
    private elem: HTMLElement;

    constructor(targetProperty: string, properties: IPropertyPaneTermSetSelectorProps) {
        this.targetProperty = targetProperty;
        this.properties = {
            key: properties.label,
            label: properties.label,
            onRender: this.onRender.bind(this),
            onPropertyChange: properties.onPropertyChange,
            currentTermSetId: properties.currentTermSetId
        };
    }

    private onRender(elem: HTMLElement): void {
        if (!this.elem) {
          this.elem = elem;
        }
     
        const element: React.ReactElement<ITermsetSelectorProps> = React.createElement(TermsetSelector, {
            label: this.properties.label,
            onChanged: this.onPropertyChanged.bind(this),
            currentTermSetId: this.properties.currentTermSetId
        });
        ReactDom.render(element, elem);
    }

    private onPropertyChanged(termSetId: Guid, grouptId: Guid) {
        this.properties.onPropertyChange(this.targetProperty, termSetId, grouptId);
    }
}