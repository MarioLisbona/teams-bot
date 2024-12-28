import { ConfidentialClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { loadEnvironmentVariables } from "../environment/setupEnvironment.js";

loadEnvironmentVariables();

// MSAL client configuration
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
};

/**
 * Acquires an access token for the Microsoft Graph API using client credentials flow.
 * @returns {Promise<string>} Promise that resolves to the access token string.
 * @throws {Error} If token acquisition fails due to invalid credentials or network issues.
 * @throws {Error} If MSAL configuration is invalid.
 */
export const getAccessToken = async () => {
  try {
    const cca = new ConfidentialClientApplication(msalConfig);
    const authResponse = await cca.acquireTokenByClientCredential({
      scopes: ["https://graph.microsoft.com/.default"],
    });

    if (!authResponse?.accessToken) {
      throw new Error("No access token received in auth response");
    }

    return authResponse.accessToken;
  } catch (error) {
    console.error("Failed to acquire access token:", {
      error: error.message,
      stack: error.stack,
    });
    throw new Error(`Authentication failed: ${error.message}`);
  }
};

// Initialize Microsoft Graph client
export const getGraphClient = async () => {
  try {
    const accessToken = await getAccessToken();
    return Client.init({
      authProvider: (done) => done(null, accessToken),
    });
  } catch (error) {
    console.error("Failed to initialize Graph client:", error);
    throw error;
  }
};
