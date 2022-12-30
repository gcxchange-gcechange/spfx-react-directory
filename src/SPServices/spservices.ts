import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISPServices } from "./ISPServices";
import {sp} from "@pnp/sp";

export class spservices implements ISPServices {
  //constructor(private context: WebPartContext) {
//     sp.setup({
//       spfxContext: this.context,
//     });
//   }

//   public async searchUsersNew(
//     context: unknown,
//     searchString: string,
//     srchQry: string,
//     isInitialSearch: boolean,
//     pageNumber?: number
//   ): Promise<SearchResults> {
//     let qrytext: string = "";
//     let client = await context.msGraphClientFactory.getClient();
//     if (isInitialSearch)
//       qrytext = `FirstName:${searchString}* OR LastName:${searchString}*`;
//     else {
//       if (srchQry) qrytext = srchQry;
//       else {
//         if (searchString) qrytext = searchString;
//       }
//       if (qrytext.length <= 0) qrytext = `*`;
//     }
//     const searchProperties: string[] = [
//       "FirstName",
//       "LastName",
//       "PreferredName",
//       "WorkEmail",
//       "OfficeNumber",
//       "PictureURL",
//       "WorkPhone",
//       "MobilePhone",
//       "JobTitle",
//       "Department",
//       "Skills",
//       "PastProjects",
//       "BaseOfficeLocation",
//       "SPS-UserType",
//       "GroupId",
//     ];
//     try {
//       console.log(qrytext);
//       let users = await sp.search(<SearchQuery>{
//         Querytext: qrytext,
//         RowLimit: 500,
//         EnableInterleaving: true,
//         SelectProperties: searchProperties,
//         SourceId: "b09a7990-05ea-4af9-81ef-edfab16c4e31",
//         SortList: [
//           { Property: "FirstName", Direction: SortDirection.Ascending },
//         ],
//       });
//       if (users && users.PrimarySearchResults.length > 0) {
//         for (
//           let index = 0;
//           index < users.PrimarySearchResults.length;
//           index++
//         ) {
//           let user: any = users.PrimarySearchResults[index];
//           let res = await client
//             .api(`/users/${user.WorkEmail}/photo/$value`)
//             .get()
//             .then(() => {
//               user = {
//                 ...user,
//                 PictureURL: `/_layouts/15/userphoto.aspx?size=M&accountname=${user.WorkEmail}`,
//               };
//               users.PrimarySearchResults[index] = user;
//             })
//             .catch((red) => {
//               user = { ...user, PictureURL: null };
//               users.PrimarySearchResults[index] = user;
//             });
//         }
//       }
//       console.log("users", users);
//       return users;
//     } catch (error) {
//       Promise.reject(error);
//     }
//   }
}