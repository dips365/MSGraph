import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ReactGroupSamplesWebPartStrings';
import ReactGroupSamples from './components/ReactGroupSamples';
import { IReactGroupSamplesProps } from './components/IReactGroupSamplesProps';

export interface IReactGroupSamplesWebPartProps {
  description: string;
  title: string;
}

export default class ReactGroupSamplesWebPart extends BaseClientSideWebPart <IReactGroupSamplesWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactGroupSamplesProps> = React.createElement(
      ReactGroupSamples,
      {
        description: this.properties.description,
        context:this.context,
        title:this.properties.title,
        displayMode:this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
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
