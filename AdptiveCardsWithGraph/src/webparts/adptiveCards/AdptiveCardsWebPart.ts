import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AdptiveCardsWebPartStrings';
import AdptiveCards from './components/AdptiveCards';
import { IAdptiveCardsProps } from './components/IAdptiveCardsProps';
import { MSGraphService } from '../../Services/MSGraphService';

export interface IAdptiveCardsWebPartProps {
  description: string;
}

export default class AdptiveCardsWebPart extends BaseClientSideWebPart <IAdptiveCardsWebPartProps> {
  private MSGraphInstance:MSGraphService
  public render(): void {
    const element: React.ReactElement<IAdptiveCardsProps> = React.createElement(
      AdptiveCards,
      {
        description: this.properties.description,
        context:this.context,
        MSGraphInstance:this.MSGraphInstance
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected async onInit(){
    await super.onInit();

    this.MSGraphInstance = new MSGraphService();

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
