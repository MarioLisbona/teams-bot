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
 *
 * @description
 * Initializes a Confidential Client Application using MSAL configuration and
 * requests an access token with Graph API default scope.
 *
 * @throws {Error} When token acquisition fails due to invalid credentials or network issues
 * @throws {Error} When MSAL configuration is invalid or missing required fields
 * @throws {Error} When no access token is received in the authentication response
 * @returns {Promise<string>} Access token for Microsoft Graph API authentication
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

/**
 * Initializes and returns an authenticated Microsoft Graph client instance.
 *
 * @description
 * Creates a Graph client configured with an access token obtained through MSAL.
 * The client is ready to make authenticated requests to Microsoft Graph API endpoints.
 *
 * @throws {Error} When access token acquisition fails
 * @throws {Error} When Graph client initialization fails
 * @throws {Error} When auth provider callback encounters an error
 * @returns {Promise<Client>} Authenticated Microsoft Graph client instance
 */
export const getGraphClient = async () => {
  try {
    const accessToken = await getAccessToken();

    if (!accessToken) {
      throw new Error("Failed to get access token");
    }

    return Client.init({
      authProvider: (done) => {
        try {
          done(null, accessToken);
        } catch (error) {
          console.error("Auth provider callback failed:", error);
          done(error, null);
        }
      },
    });
  } catch (error) {
    console.error("Failed to initialize Graph client:", {
      error: error.message,
      stack: error.stack,
    });
    throw new Error(`Graph client initialization failed: ${error.message}`);
  }
};
