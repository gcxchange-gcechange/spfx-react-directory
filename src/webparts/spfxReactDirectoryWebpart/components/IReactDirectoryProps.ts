import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IReactDirectoryProps {
  context: WebPartContext;
  pageSize: number;
  prefLang: string;
  hidingUsers: string;
}
