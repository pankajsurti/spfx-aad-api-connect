import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-webpart-base';

import * as strings from 'SampleMsGraphApiWebPartStrings';
import SampleMsGraphApi from './components/SampleMsGraphApi';
import { ClientMode } from './components/ClientMode';
import { ISampleMsGraphApiProps } from './components/ISampleMsGraphApiProps';

export interface ISampleMsGraphApiWebPartProps {
  description: string;
  clientMode: ClientMode;
}

export default class SampleMsGraphApiWebPart extends BaseClientSideWebPart<ISampleMsGraphApiWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISampleMsGraphApiProps > = React.createElement(
      SampleMsGraphApi,
      {
        description: this.properties.description,
        clientMode: this.properties.clientMode,
        context: this.context,
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
                PropertyPaneChoiceGroup('clientMode', {
                  label: strings.ClientModeLabel,
                  options: [
                    { key: ClientMode.aad, text: "AadHttpClient"},
                    { key: ClientMode.graph, text: "MSGraphClient"},
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
