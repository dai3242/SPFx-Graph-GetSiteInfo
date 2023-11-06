import * as ReactDom from 'react-dom';
import * as React from "react";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'GraphApiWebPartStrings';
import { IGraphApiProps } from './components/IGraphApiProps';
import GraphApi from './components/GraphApi';

export default class GraphApiWebPart extends BaseClientSideWebPart<IGraphApiProps> {


  public render(): void {
    ReactDom.render(
      <GraphApi context={this.context} properties={this.properties} />,
      this.domElement
    );
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
