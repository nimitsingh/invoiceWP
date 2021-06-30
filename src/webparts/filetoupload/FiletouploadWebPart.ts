import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart,WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'FiletouploadWebPartStrings';
import Filetoupload from './components/Filetoupload';
import { IFiletouploadProps } from './components/IFiletouploadProps'; 
import { SPHttpClient } from '@microsoft/sp-http';

export interface IFiletouploadWebPartProps {
  description: string;
  spHttpClient: SPHttpClient;
  currentContext: WebPartContext;
  label:string;
  listGUID: string;
}

export default class FiletouploadWebPart extends BaseClientSideWebPart <IFiletouploadWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IFiletouploadProps> = React.createElement(
      Filetoupload,
      {
        description: this.properties.description,
        spHttpClient: this.context.spHttpClient,
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
