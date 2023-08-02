import { PublicClientApplication, EventType } from "@azure/msal-browser";
import { msalConfig, protectedResources } from '../../../../authConfig';
import { InteractionType } from '@azure/msal-browser';
import { useMsal, useMsalAuthentication } from "@azure/msal-react";
import { getClaimsFromStorage } from "../../../../utils/storageUtils";
import ChatService from "../SPServices/ChatService";
import * as React from "react";


const Chats = () => {
    const { instance } = useMsal();
    const activeAccount: any = instance.getActiveAccount();        
    const accountName: string = activeAccount ? activeAccount.username + ' (' + activeAccount.localAccountId + ')' :  'not active';    

    const resource = new URL('https://appsvc-fnc-dev-scw-obo-poc-dotnet001.azurewebsites.net/api/ConnectAsUser').hostname;
    const request = {
        scopes: ['User.Read'],
        account: activeAccount,
        claims: activeAccount && getClaimsFromStorage(`cc.${msalConfig.auth.clientId}.${activeAccount.idTokenClaims.oid}.${resource}`)
                ? window.atob(getClaimsFromStorage(`cc.${msalConfig.auth.clientId}.${activeAccount.idTokenClaims.oid}.${resource}`))
                : undefined,
        sid: ChatService.context.pageContext.legacyPageContext.aadSessionId
    };
    console.log("request", request);

    const { acquireToken, result, error } = useMsalAuthentication(
        InteractionType.Silent, {...request, redirectUri: ''}
    );

    if (error) {
        console.log("Chats error", error);
    }

    return (
        <>
        <div><strong>Chats</strong></div>
        <div>accountName: {accountName}</div>
        <br />
        </>
    );
}

export default Chats