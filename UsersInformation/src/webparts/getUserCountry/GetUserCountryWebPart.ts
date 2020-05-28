import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'GetUserCountryWebPartStrings';
import GetUserCountry from './components/GetUserCountry';
import { IGetUserCountryProps } from './components/IGetUserCountryProps';
import { MSGraphService } from "../../Services/MSGraphService";

export interface IGetUserCountryWebPartProps {
  description: string;
}

export default class GetUserCountryWebPart extends BaseClientSideWebPart <IGetUserCountryWebPartProps> {
  private MsGraphServiceInstance:MSGraphService;
  public render(): void {
    const element: React.ReactElement<IGetUserCountryProps> = React.createElement(
      GetUserCountry,
      {
        description: this.properties.description,
        context:this.context,
        MsGraphServiceInstance: this.MsGraphServiceInstance
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(){
    await super.onInit();

    this.MsGraphServiceInstance = new MSGraphService();
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
