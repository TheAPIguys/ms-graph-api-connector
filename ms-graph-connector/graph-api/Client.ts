import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import {
    TokenCredentialAuthenticationProvider,
    TokenCredentialAuthenticationProviderOptions,
} from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

export async function InitGraphClient() {
    const tokenCredential = new ClientSecretCredential(
        process.env.TENANT_ID ? process.env.TENANT_ID : '', // TENANT_ID
        process.env.CLIENT_ID ? process.env.CLIENT_ID : '',
        process.env.CLIENT_SECRET ? process.env.CLIENT_SECRET : '',
    );

    // Set your scopes and options for TokenCredential.getToken (Check the ` interface GetTokenOptions` in (TokenCredential Implementation)[https://github.com/Azure/azure-sdk-for-js/blob/master/sdk/core/core-auth/src/tokenCredential.ts])

    const options: TokenCredentialAuthenticationProviderOptions = {
        scopes: ['https://graph.microsoft.com/.default'],
    };

    // Create an instance of the TokenCredentialAuthenticationProvider by passing the tokenCredential instance and options to the constructor
    const authProvider = new TokenCredentialAuthenticationProvider(tokenCredential, options);
    return Client.initWithMiddleware({
        debugLogging: true,
        authProvider: authProvider,
    });
}
