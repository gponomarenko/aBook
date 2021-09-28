import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ABookWebPartStrings';

import "@pnp/sp/webs";
import { IItemAddResult, sp, Web } from "@pnp/sp/presets/all";
import { ABookFC } from './components/ABookFC';
import { IABookWebPartProps } from './components/IABookWebPartProps'

export default class ABookWebPart extends BaseClientSideWebPart<IABookWebPartProps> {

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {  
      sp.setup({
        spfxContext: this.context,
        sp: {
          headers: {
            Accept: "application/json; odata=nometadata"
          }
        }
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IABookWebPartProps> = React.createElement(
      ABookFC,
      {
        description: this.properties.description,
        context: this.context
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
