/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { msalConfig } from './authConfig';
import { addClaimsToStorage } from './utils/storageUtils';
import { parseChallenges } from './utils/claimUtils';

/**
 * This method inspects the HTTPS response from a fetch call for the "www-authenticate header"
 * If present, it grabs the claims challenge from the header and store it in the localStorage
 * For more information, visit: https://docs.microsoft.com/en-us/azure/active-directory/develop/claims-challenge#claims-challenge-header-format
 * @param {object} response
 * @returns response
 */
export const handleClaimsChallenge = async (response: Response, apiEndpoint: string | URL, account: { homeAccountId?: string; environment?: string; tenantId?: string; username?: string; localAccountId?: string; name?: string; idToken?: string; idTokenClaims: any; nativeAccountId?: string; }) => {
    console.log("handleClaimsChallenge - response.status", response.status);

    if (response.status === 200) {
        return response.json();
    } else if (response.status === 401) {   // Unauthorized
        if (response.headers.get('WWW-Authenticate')) {
            const authenticateHeader = response.headers.get('WWW-Authenticate');
            const claimsChallenge = parseChallenges(authenticateHeader);

            /**
             * This method stores the claim challenge to the session storage in the browser to be used when acquiring a token.
             * To ensure that we are fetching the correct claim from the storage, we are using the clientId
             * of the application and oid (user’s object id) as the key identifier of the claim with schema
             * cc.<clientId>.<oid>.<resource.hostname>
             */
            addClaimsToStorage(
                `cc.${msalConfig.auth.clientId}.${account.idTokenClaims.oid}.${new URL(apiEndpoint).hostname}`,
                claimsChallenge.claims,
            );

            //console.log("msalConfig.auth.clientId", msalConfig.auth.clientId);
            throw new Error(`claims_challenge_occurred`);
        }

        throw new Error(`Unauthorized: ${response.status}`);
    } else {
        throw new Error(`Something went wrong with the request: ${response.status}`);
    }
};

/**
 * Makes a fetch call to the API endpoint with the access token in the Authorization header
 * @param {string} accessToken 
 * @param {string} apiEndpoint 
 * @returns 
 */
export const callApiWithToken = async (accessToken: string, apiEndpoint: string, account: any, userId: string = '') => {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    apiEndpoint = apiEndpoint + '?userId=' + userId;
    //console.log('apiEndpoint = ' + apiEndpoint);

    const response = await fetch(apiEndpoint, options);
    return handleClaimsChallenge(response, apiEndpoint, account);
};