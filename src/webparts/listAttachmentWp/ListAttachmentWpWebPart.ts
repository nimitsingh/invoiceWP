import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart,WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'ListAttachmentWpWebPartStrings';
import ListAttachmentWp from './components/ListAttachmentWp';
import { IListAttachmentWpProps } from './components/IListAttachmentWpProps';

export interface IListAttachmentWpWebPartProps {
  description: string;
  currentContext: WebPartContext;
}

export default class ListAttachmentWpWebPart extends BaseClientSideWebPart <IListAttachmentWpWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IListAttachmentWpProps> = React.createElement(
      ListAttachmentWp,
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
