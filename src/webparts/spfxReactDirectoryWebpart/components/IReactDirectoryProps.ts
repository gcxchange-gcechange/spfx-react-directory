import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from "@microsoft/sp-core-library";
export interface IReactDirectoryProps {
  displayMode: DisplayMode;
  context: WebPartContext;
  pageSize?: number;
  prefLang: string;
  hidingUsers: string;
}
