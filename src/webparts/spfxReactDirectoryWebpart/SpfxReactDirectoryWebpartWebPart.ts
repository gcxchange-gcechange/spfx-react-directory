import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "SpfxReactDirectoryWebpartWebPartStrings";
import DirectoryHook from "./components/DirectoryHook";
import { IReactDirectoryProps } from "./components/IReactDirectoryProps";

export interface ISpfxReactDirectoryWebpartWebPartProps {
  title: string;
  searchProps: string;
  pageSize: number;
  prefLang: string;
  hidingUsers: string;
}

export default class SpfxReactDirectoryWebpartWebPart extends BaseClientSideWebPart<ISpfxReactDirectoryWebpartWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  public render(): void {
    const element: React.ReactElement<IReactDirectoryProps> = React.createElement(DirectoryHook, {
      title: this.properties.title,
      context: this.context,
      displayMode: this.displayMode,
      updateProperty: (value: string) => {
        this.properties.title = value;
      },
      pageSize: this.properties.pageSize,
      prefLang: this.properties.prefLang,
      hidingUsers: this.properties.hidingUsers,
    });

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext().then((context) => {
        let environmentMessage: string = "";
        switch (context.app.host.name) {
          case "Office": // running in Office
            environmentMessage = this.context.isServedFromLocalhost
              ? strings.AppLocalEnvironmentOffice
              : strings.AppOfficeEnvironment;
            break;
          case "Outlook": // running in Outlook
            environmentMessage = this.context.isServedFromLocalhost
              ? strings.AppLocalEnvironmentOutlook
              : strings.AppOutlookEnvironment;
            break;
          case "Teams": // running in Teams
            environmentMessage = this.context.isServedFromLocalhost
              ? strings.AppLocalEnvironmentTeams
              : strings.AppTeamsTabEnvironment;
            break;
          default:
            throw new Error("Unknown host");
        }

        return environmentMessage;
      });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty("--bodyText", semanticColors.bodyText || null);
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty("--linkHovered", semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown("prefLang", {
                  label: "Preferred Language",
                  options: [
                    { key: "account", text: "Account" },
                    { key: "en-us", text: "English" },
                    { key: "fr-fr", text: "Fran??ais" },
                  ],
                }),
                PropertyPaneTextField("hidingUsers", {
                  label: "Users not in serach",
                  description: "Enter the user ids of the users who are not needed in the search separated by '/' ",
                  multiline: true,
                  rows: 10,
                }),
                PropertyPaneSlider("pageSize", {
                  label: "Results per page",
                  showValue: true,
                  max: 20,
                  min: 2,
                  step: 2,
                  value: this.properties.pageSize,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
