import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { IDateTimeFieldValue } from "@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker";

export interface IABookProps {
  Title: string;
  addressEmployee: string;
  birthdayEmployee: string;
  fullName: string;
  jobTitle: string;
  levelEmployee: string;
  managerOfEmployee: string;
  statusEmployee: string;
}