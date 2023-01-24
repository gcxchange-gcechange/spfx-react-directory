import { WebPartContext } from "@microsoft/sp-webpart-base";
//import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { IUserProperties } from "./IUserProperties";

export interface IPersonaCardProps {
  context: WebPartContext;
  profileProperties: IUserProperties;
  prefLang: string;
}
