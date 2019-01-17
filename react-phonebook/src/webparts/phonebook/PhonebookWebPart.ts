import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'PhonebookWebPartStrings';
import Phonebook from './components/Phonebook';
import { IPhonebookProps } from './components/IPhonebookProps';
import { PeopleDataProvider } from '../../services/PeopleDataProvider';
import { OrgStructureProvider } from '../../services/OrgStructureProvider';
import { IPerson } from '../../classes/IPerson';
import { IOrgUnit } from '../../classes/IOrgUnit';

export interface IPhonebookWebPartProps {
  description: string;
}



export default class PhonebookWebPart extends BaseClientSideWebPart<IPhonebookWebPartProps> {

  private getPeople():Promise<IPerson[]>{
    return PeopleDataProvider.getPeople("");
  }
  private getOrgUnits():Promise<IOrgUnit[]>{
    return OrgStructureProvider.getAllTree();
  }
  
  public render(): void {
    const element: React.ReactElement<IPhonebookProps > = React.createElement(
      Phonebook,
      {
        getPeople: this.getPeople(),
        getOrgUnits: this.getOrgUnits()
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
