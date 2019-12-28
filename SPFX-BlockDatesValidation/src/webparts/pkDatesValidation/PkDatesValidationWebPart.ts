import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'PkDatesValidationWebPartStrings';
import BodyContent from './components/BodyContent/BodyContent';
import { IPkDatesValidationProps } from './components/IPkDatesValidationProps';
import { IIconProps } from 'office-ui-fabric-react';

export interface IPkDatesValidationWebPartProps {
  description: string;
}

export default class PkDatesValidationWebPart extends BaseClientSideWebPart<IPkDatesValidationWebPartProps> {

  public render(): void {
    const element: React.ReactElement = React.createElement(
      BodyContent
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
