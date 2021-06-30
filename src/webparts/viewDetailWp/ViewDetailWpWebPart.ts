import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart,WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'ViewDetailWpWebPartStrings';
import ViewDetailWp from './components/ViewDetailWp';
import { IViewDetailWpProps } from './components/IViewDetailWpProps';

export interface IViewDetailWpWebPartProps {
  description: string;
  currentContext: WebPartContext;
  listGUID: string;
}

export default class ViewDetailWpWebPart extends BaseClientSideWebPart <IViewDetailWpWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IViewDetailWpProps> = React.createElement(
      ViewDetailWp,
      {
        description: this.properties.description,
        currentContext: this.context,
        listGUID: this.properties.listGUID
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
                }),
                PropertyPaneTextField('listGUID', {
                  label: 'Enter the List GUID'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
