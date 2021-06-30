import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'DashboardGridWebPartStrings';
import DashboardGrid from './components/DashboardGrid';
import { IDashboardGridProps } from './components/IDashboardGridProps';

export interface IDashboardGridWebPartProps {
  description: string;
  currentContext: WebPartContext;
}

export default class DashboardGridWebPart extends BaseClientSideWebPart <IDashboardGridWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDashboardGridProps> = React.createElement(
      DashboardGrid,
      {
        description: this.properties.description,
        currentContext: this.context
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
