/* eslint-disable @typescript-eslint/no-explicit-any */
export interface ISPServices {
  searchUsersNew(
    context: any,
    searchString: string,
    srchQry: string,
    isInitialSearch: boolean,
    hidingUsers: any,
    startItem: number,
    endItem: number,
    //accessToken: any,
    //activeAccount: any
  );
}
