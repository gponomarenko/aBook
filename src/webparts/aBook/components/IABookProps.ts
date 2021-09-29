import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { IDateTimeFieldValue } from "@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker";

export interface IABookProps {
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
    levelEmployee: number;
    managerCard?: {
      EMail: string;
      Id: number;
      Title: string;
    };          
    managerCardId?: number;
    managerCardStringId?: string;
    managerOfEmployee?: string;
    statusEmployee: string;
    VISA?: string;
}
    // ID: 29
    // Id: 29
    // Title: "Денис Савельев"
    // addressEmployee: null
    // birthdayEmployee: "09/07/2000 00:00:00"
    // employeeCard:
          // Department: "Департамент ИТ аутсорсинга"
          // EMail: "ds@techexpert.onmicrosoft.com"
          // Id: 11
          // JobTitle: "SharePoint Developer"
          // MobilePhone: null
          // Office: "Turkey"
          // Title: "Денис Савельев"
          // WorkPhone: "9379992"
    // employeeCardId: 11
    // employeeCardStringId: "11"
    // fullName: "Денис Савельев"
    // jobTitle: "SharePoint Developer"
    // levelEmployee: null
    // managerCard:
          // EMail: "id@techexpert.onmicrosoft.com"
          // Id: 13
          // Title: "Ирина Держук"
    // managerCardId: 13
    // managerCardStringId: "13"
    // managerOfEmployee: null
    // statusEmployee: "active"