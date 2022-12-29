import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IReactDirectoryProps {
  title: string;
  displayMode: DisplayMode;
  context: WebPartContext;
  searchFirstName: boolean;
  updateProperty: (value: string) => void;
  searchProps?: string;
  clearTextSearchProps?: string;
  pageSize?: number;
  prefLang: string;
}
