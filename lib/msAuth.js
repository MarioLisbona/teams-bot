import { ConfidentialClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { loadEnvironmentVariables } from "./environment/setupEnvironment.js";

loadEnvironmentVariables();

console.log("msAuth--->", process.env.ROOT_DIRECTORY_NAME);

// MSAL client configuration
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
};

// Get an access token for the Microsoft Graph API
export const getAccessToken = async () => {
  const cca = new ConfidentialClientApplication(msalConfig);

  const authResponse = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });

  return authResponse.accessToken;
};

// Initialize Microsoft Graph client
export const getGraphClient = async () => {
  const accessToken = await getAccessToken();
  ``;
  return Client.init({
    authProvider: (done) => done(null, accessToken),
  });
};
