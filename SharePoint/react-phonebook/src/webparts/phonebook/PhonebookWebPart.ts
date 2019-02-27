import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Guid, ServiceScope } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';

import * as strings from 'PhonebookWebPartStrings';
import Phonebook from './components/Phonebook';
import { IPhonebookProps } from './components/IPhonebookProps';
import { OrgStructureService } from '../../services/OrgStructure/OrgStructureService';
import { IPerson } from '../../classes/IPerson';
import { IOrgUnit } from '../../classes/IOrgUnit';
import { PropertyPaneTermSetSelector } from '../../PropertyControls/PropertyPaneTermSetSelector';
import { update } from 'lodash';
import { IOrgStructureService } from '../../services/OrgStructure/IOrgStructureService';
import { IPeopleSearchService } from '../../services/People/IPeopleSearchService';
import { PeopleSearchService } from '../../services/People/PeopleSearchService';

export interface IPhonebookWebPartProps {
  termSetId: string;
  groupId: string;
  pageSize: number;
}



export default class PhonebookWebPart extends BaseClientSideWebPart<IPhonebookWebPartProps> {
  private _orgStructureService: IOrgStructureService;
  private _peopleSeacrhService: IPeopleSearchService;

  public onInit(): Promise<void> {
    this._orgStructureService = this.context.serviceScope.consume(OrgStructureService.serviceKey);
    this._peopleSeacrhService = this.context.serviceScope.consume(PeopleSearchService.serviceKey);
    return super.onInit();
  }

  private getPeople(query: string, orgUnitId: Guid, page: number, pageSize: number):Promise<{persons: IPerson[], numberOfResults: number}>{
    return this._peopleSeacrhService.getPeople(query, orgUnitId, page, pageSize);
  }
  private getOrgUnits():Promise<IOrgUnit[]>{
    return this._orgStructureService.getFullTree(this.properties.groupId, this.properties.termSetId);
  }
  
  public render(): void {
    const element: React.ReactElement<IPhonebookProps > = React.createElement(
      Phonebook,
      {
        getPeople: (query: string, orgUnitId: Guid, page: number, pageSize: number) => this.getPeople(query,orgUnitId, page, pageSize),
        getOrgUnits: this.getOrgUnits(),
        pageSize: this.properties.pageSize,
        orgStructureTermSetId: this.properties.termSetId
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private onTermSetChange(propertyPath: string, termSetId: Guid, groupId: Guid) {
    update(this.properties, propertyPath, (): any => { return termSetId.toString(); });
    update(this.properties, "groupId", (): any => { return groupId.toString(); });
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneSlider('pageSize', {
                  min: 5,
                  max:100,
                  label: strings.PageSizeFieldLabel
                }),
                new PropertyPaneTermSetSelector('termSetId', {
                  label: strings.OrgStructureTermSetFieldLabel,
                  onPropertyChange: this.onTermSetChange.bind(this),
                  currentTermSetId: this.properties.termSetId
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
