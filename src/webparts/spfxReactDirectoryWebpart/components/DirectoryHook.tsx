import * as React from "react";
import { IReactDirectoryProps } from "./IReactDirectoryProps";
import PersonaCardMain from "./PersonaCard/PersonaCardMain";
import { EventType, PublicClientApplication } from "@azure/msal-browser";
import { msalConfig } from "../../../authConfig";
import { MsalProvider } from "@azure/msal-react";

const pca = new PublicClientApplication(msalConfig);

if (!pca.getActiveAccount() && pca.getAllAccounts().length > 0) {
  pca.setActiveAccount(pca.getActiveAccount());
}

// Optional - This will update account state if a user signs in from another tab or window
pca.enableAccountStorageEvents();

pca.addEventCallback((event: any) => {
    if (event.eventType === EventType.SSO_SILENT_SUCCESS && event.payload.account) {
        const account = event.payload.account;
        pca.setActiveAccount(account);
    }
});

const DirectoryHook: React.FC<IReactDirectoryProps> = (props) => {
  return (
    <MsalProvider instance={pca}>
      <PersonaCardMain 
                       context={props.context}
                       displayMode={props.displayMode}
                       pageSize={props.pageSize}
                       prefLang={props.prefLang}
                       hidingUsers={props.hidingUsers} />
  </MsalProvider>                     
  );
};

export default DirectoryHook;
