import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-webpart-base';

import * as strings from 'BirthdayListWebPartStrings';
import BirthdayList from './components/BirthdayList';
import { IBirthdayListProps } from './components/IBirthdayListProps';
import { IPeopleService, PeopleService } from '../../services/PeopleService';
import { IPerson } from '../../classes/IPerson';

export interface IBirthdayListWebPartProps {
  description: string;
  displayType: number;
  birthdayManagedProperty: string;
  additionalQuery: string;
}

export default class BirthdayListWebPart extends BaseClientSideWebPart<IBirthdayListWebPartProps> {
  private _peopleService: IPeopleService;

  public onInit(): Promise<void> {
    this._peopleService = this.context.serviceScope.consume(PeopleService.serviceKey);
    return super.onInit();
  }

  private async loadBirthdayList(): Promise<IPerson[]> {
    let currentDate = new Date();
    let startDate: Date = new Date();
    let endDate: Date = new Date();
    switch (this.properties.displayType) {
      case 0:
        return this._peopleService.getBirthdayList(currentDate, this.properties.birthdayManagedProperty, this.properties.additionalQuery);
      case 1:
        const dayIndex: number = currentDate.getDay();
        startDate.setDate(currentDate.getDate() + 1 - dayIndex);
        endDate.setDate(currentDate.getDate() + 7 - dayIndex);
        return this._peopleService.getBirthdayListForRange(startDate, endDate, this.properties.birthdayManagedProperty, this.properties.additionalQuery);
      case 2:
        startDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
        endDate = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 1);
        let endDateFinal = new Date(endDate.setDate(endDate.getDate() - 1));
        return this._peopleService.getBirthdayListForRange(startDate, endDateFinal, this.properties.birthdayManagedProperty, this.properties.additionalQuery);
    }
    
  }

  public render(): void {
    const element: React.ReactElement<IBirthdayListProps > = React.createElement(
      BirthdayList,
      {
        description: this.properties.description,
        getData: this.loadBirthdayList.bind(this),
        displayType: this.properties.displayType,
        additionalQuery: this.properties.additionalQuery
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
    const day: any = require('./components/images/calendar_day.png');
    const week: any = require('./components/images/calendar_week.png');
    const month: any = require('./components/images/calendar_month.png');

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
                PropertyPaneTextField('birthdayManagedProperty', {
                  label: strings.birthdayManagedPropertyLabel
                }),
                PropertyPaneTextField('additionalQuery', {
                  multiline: true,
                  label: strings.additionalQueryLabel
                }),
                PropertyPaneChoiceGroup('displayType', {
                  label: "Select display type",
                  options: [
                    {key: 0, text: "Today", checked: this.properties.displayType == 0, imageSrc: day, selectedImageSrc: day},
                    {key: 1, text: "Week" , checked: this.properties.displayType == 1, imageSrc: week, selectedImageSrc: week},
                    {key: 2, text: "Month" , checked: this.properties.displayType == 2, imageSrc: month, selectedImageSrc: month}
                  ], 
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
