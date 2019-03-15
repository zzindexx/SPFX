import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SiteQuotaWebPartStrings';
import SiteQuota from './components/SiteQuota';
import { ISiteQuotaProps } from './components/ISiteQuotaProps';

export interface ISiteQuotaWebPartProps {
}

export default class SiteQuotaWebPart extends BaseClientSideWebPart<ISiteQuotaWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISiteQuotaProps > = React.createElement(
      SiteQuota,
      {
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
                
              ]
            }
          ]
        }
      ]
    };
  }
}
