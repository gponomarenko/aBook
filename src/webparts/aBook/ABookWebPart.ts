import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ABookWebPartStrings';

import { IABookProps } from './components/IABookProps';
import "@pnp/sp/webs";
import { IItemAddResult, sp, Web } from "@pnp/sp/presets/all";
import ABookFC from './components/ABookFC';

export interface IABookWebPartProps {
  Title: string;
  addressEmployee?: string;
  birthdayEmployee?:  string;
  employeeCard?: {
    Department:  string;
    EMail: string;
    Id: number;
    JobTitle: string;
    MobilePhone: string;
    Office: string;
    Title: string;
    WorkPhone: string;
  };
        
  employeeCardId?: number;
  employeeCardStringId?: string;
  fullName?: string;
  jobTitle?: string;
  levelEmployee: string;
  managerCard?: {
    EMail: string;
    Id: number;
    Title: string;
  };          
  managerCardId?: number;
  managerCardStringId?: string;
  managerOfEmployee?: string;
  statusEmployee: string;
}

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
        Title: this.properties.Title,
        addressEmployee: this.properties.addressEmployee,
        birthdayEmployee: this.properties.birthdayEmployee,
        fullName: this.properties.fullName,
        jobTitle: this.properties.jobTitle,
        levelEmployee: this.properties.levelEmployee,
        managerOfEmployee: this.properties.managerOfEmployee,
        statusEmployee: this.properties.statusEmployee,
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
