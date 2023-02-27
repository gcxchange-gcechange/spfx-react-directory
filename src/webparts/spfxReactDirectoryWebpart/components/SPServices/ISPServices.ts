/* eslint-disable @typescript-eslint/no-explicit-any */
export interface ISPServices {
  // searchUsers(searchString: string, searchFirstName: boolean);
  searchUsersNew(
    context: any,
    searchString: string,
    srchQry: string,
    isInitialSearch: boolean,
    hidingUsers: any,
    startItem: number,
    endItem: number
  );
}
