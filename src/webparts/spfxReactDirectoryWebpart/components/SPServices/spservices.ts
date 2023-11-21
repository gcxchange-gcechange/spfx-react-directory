// import { WebPartContext } from "@microsoft/sp-webpart-base";
// //import { sp, SearchQuery, SearchResults, SortDirection } from "@pnp/sp";
// import { SearchResults, ISearchQuery, SortDirection } from "@pnp/sp/search";

// import { ISPServices } from "./ISPServices";
// import { useMsal, useMsalAuthentication } from "@azure/msal-react";
// import ChatService from "./ChatService";
// import { getClaimsFromStorage } from "../../../../utils/storageUtils";
// import { msalConfig, protectedResources } from "../../../../authConfig";
// import { InteractionType } from '@azure/msal-browser';
// import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import { SearchResults, ISearchQuery, SortDirection } from "@pnp/sp/search";
import { ISPServices } from "./ISPServices";

export class spservices implements ISPServices {
  // constructor(private context: WebPartContext) {
  //   sp.setup({
  //     spfxContext: this.context,
  //   });
  // }
  constructor(private context: WebPartContext) {
    sp.setup({
      spfxContext: {
        pageContext: {
          web: {
            absoluteUrl: this.context.pageContext.web.absoluteUrl,
          },
        },
      },
    });
  }

  public async searchUsersNew(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    context: any,
    searchString: string,
    srchQry: string,
    isInitialSearch: boolean,
    hidingUsers: string[],
    startItem: number,
    endItem: number
    //accessToken: any,
    //activeAccount: any,
  ): Promise<SearchResults> {
    let qrytext: string = "";

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
      const users = await sp.search(<ISearchQuery>{
        Querytext: qrytext,
        StartRow: startItem,
        RowLimit: endItem,
        EnableInterleaving: true,
        SelectProperties: searchProperties,
        SourceId: "b09a7990-05ea-4af9-81ef-edfab16c4e31",
        SortList: [{ Property: "FirstName", Direction: SortDirection.Ascending }],
      });

      let n = users.PrimarySearchResults.length;
      if (users && n > 0) {
        for (let index = 0; index < n; index++) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const user: any = users.PrimarySearchResults[index];
          if (hidingUsers.indexOf(user.UniqueId) !== -1) {
            users.PrimarySearchResults.splice(index, 1);
            n = n - 1;
            index = index - 1;
          }
        }

        const client = await context.msGraphClientFactory.getClient();
        const body = { requests: [] };
        users.PrimarySearchResults.forEach((user) => {
          const requestUrl: string = `/users/${user.UniqueId}/photo/$value`;
          body.requests.push({
            id: user.UniqueId.toString(),
            method: "GET",
            url: requestUrl,
          });
        });
        const response = await client.api("$batch").version("v1.0").post(body);

        response.responses.forEach((r) => {
          if (r.status === 200) {
            users.PrimarySearchResults.map((u, index) => {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              let user: any = users.PrimarySearchResults[index];
              if (r.id === u.UniqueId) {
                user = {
                  ...user,
                  PictureURL: `data:${r.headers["Content-Type"]};base64,${r.body}`,
                };
                users.PrimarySearchResults[index] = user;
              }
            });
          }

          if (r.status !== 200) {
            users.PrimarySearchResults.map((u, index) => {
              if (r.id === u.UniqueId) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                let user: any = users.PrimarySearchResults[index];
                user = {
                  ...user,
                  PictureURL: null,
                };
                users.PrimarySearchResults[index] = user;
              }
            });
          }
        });
      }

      return users;
    } catch (error) {
      Promise.reject(error)
        .then((data) => {
          return data;
        })
        .catch((err) => {
          /* perform error handling if desired */
        });
    }
  }
}
