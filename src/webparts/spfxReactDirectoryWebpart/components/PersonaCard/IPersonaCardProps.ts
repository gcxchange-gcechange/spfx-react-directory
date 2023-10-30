import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IUserProperties } from "./IUserProperties";

export interface IPersonaCardProps {
  context: WebPartContext;
  profileProperties: IUserProperties;
  prefLang: string;
  activeAccount: any;
  instance: any;
}
