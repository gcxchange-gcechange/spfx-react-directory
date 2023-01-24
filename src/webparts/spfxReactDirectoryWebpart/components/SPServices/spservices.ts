import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp, SearchQuery, SearchResults, SortDirection } from "@pnp/sp";

import { ISPServices } from "./ISPServices";

export class spservices implements ISPServices {
  constructor(private context: WebPartContext) {
    sp.setup({
      spfxContext: this.context,
    });
  }

  public async searchUsersNew(
    context: any,
    searchString: string,
    srchQry: string,
    isInitialSearch: boolean,
    hidingUsers: string[],
    pageNumber?: number
  ): Promise<SearchResults> {
    let qrytext: string = "";
    const client = await context.msGraphClientFactory.getClient();
    if (isInitialSearch) qrytext = `FirstName:${searchString}* OR LastName:${searchString}*`;
    else {
      if (srchQry) qrytext = srchQry;
      else {
        if (searchString) qrytext = searchString;
      }
      if (qrytext.length <= 0) qrytext = `*`;
    }

    const searchProperties: string[] = [
      "FirstName",
      "LastName",
      "PreferredName",
      "WorkEmail",
      "OfficeNumber",
      "PictureURL",
      "WorkPhone",
      "MobilePhone",
      "JobTitle",
      "Department",
      "Skills",
      "PastProjects",
      "BaseOfficeLocation",
      "SPS-UserType",
      "GroupId",
    ];
    try {
      let users = await sp.search(<SearchQuery>{
        Querytext: qrytext,
        RowLimit: 500,
        EnableInterleaving: true,
        SelectProperties: searchProperties,
        SourceId: "b09a7990-05ea-4af9-81ef-edfab16c4e31",
        SortList: [{ Property: "FirstName", Direction: SortDirection.Ascending }],
      });
      let n = users.PrimarySearchResults.length;
      if (users && n > 0) {
        for (let index = 0; index < n; index++) {
          let user: any = users.PrimarySearchResults[index];
          if (hidingUsers.indexOf(user.UniqueId) != -1) {
            users.PrimarySearchResults.splice(index, 1);
            n = n - 1;
            index = index - 1;
          } else {
            user = {
              ...user,
              PictureURL: null,
            };
          }
        }
      }
      return users;
    } catch (error) {
      Promise.reject(error);
    }
  }
}
