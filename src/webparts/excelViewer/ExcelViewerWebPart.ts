import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ExcelViewerWebPartStrings';
import ExcelViewer from './components/ExcelViewer';
import { IExcelViewerProps } from './components/IExcelViewerProps';

export interface IExcelViewerWebPartProps {
  description: string;
}

export default class ExcelViewerWebPart extends BaseClientSideWebPart<IExcelViewerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IExcelViewerProps > = React.createElement(
      ExcelViewer,
      {
        description: this.properties.description
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
