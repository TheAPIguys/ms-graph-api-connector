import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import {
  TokenCredentialAuthenticationProvider,
  TokenCredentialAuthenticationProviderOptions,
} from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

export async function InitGraphClient() {
  const tokenCredential = new ClientSecretCredential(
    '71124398-4b74-4832-b48a-1c01fb46e497', // TENANT_ID  process.env.TENANT_ID ? process.env.TENANT_ID :
    'a4318a83-8311-407d-b571-768aff5588de', // CLIENT_ID process.env.CLIENT_ID ? process.env.CLIENT_ID :
    'rid8Q~hZ13OKFa0HgH5ceYdyCzXv5Nm4HR_A-bLr',
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
