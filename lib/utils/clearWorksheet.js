import { getGraphClient } from "../auth/msAuth.js"; // Import the authentication functions

/**
 * Clears content from specified ranges in an Excel worksheet using Microsoft Graph API.
 *
 * @description
 * Handles worksheet content clearing by:
 * 1. Initializing Microsoft Graph client
 * 2. Processing each range sequentially
 * 3. Using Graph API's clear endpoint for each range
 * 4. Providing detailed logging and error handling
 *
 * The function uses the SharePoint site ID from environment variables
 * and applies the "Contents" clear operation to each range.
 *
 * @param {string} workbookId - Excel workbook identifier in SharePoint
 * @param {string} worksheetName - Name of the target worksheet
 * @param {Array<string>} ranges - Array of Excel range addresses to clear
 *                                (e.g., ["A1:B10", "D1:E10"])
 *
 * @throws {Error} When Graph client initialization fails
 * @throws {Error} When clearing specific ranges fails
 * @throws {Error} When processing multiple ranges fails
 * @returns {Promise<void>} Resolves when all ranges are cleared
 *
 * @example
 * try {
 *   await clearWorksheetRange(
 *     "1234567890",
 *     "Sheet1",
 *     ["A1:B10", "D1:E10"]
 *   );
 * } catch (error) {
 *   console.error("Failed to clear ranges:", error);
 * }
 *
 * @requires Microsoft Graph API
 * @requires SHAREPOINT_SITE_ID environment variable
 */
export const clearWorksheetRange = async (
  workbookId,
  worksheetName,
  ranges
) => {
  try {
    // Initialize the Graph client
    const client = await getGraphClient();

    const clear = {
      applyTo: "Contents",
    };

    try {
      for (const range of ranges) {
        try {
          const apiEndpoint = `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${worksheetName}/range(address='${range}')/clear`;
          await client.api(apiEndpoint).post(clear);
          console.log(
            `Cleared contents of range ${range} in worksheet - ${worksheetName}.`
          );
        } catch (error) {
          console.error(`Failed to clear range ${range}:`, error);
          throw new Error(`Failed to clear range ${range}: ${error.message}`);
        }
      }
    } catch (error) {
      console.error("Failed to process ranges:", error);
      throw new Error(`Failed to clear worksheet ranges: ${error.message}`);
    }
  } catch (error) {
    console.error("Error clearing worksheet range:", error.message);
    console.error("Full error details:", error);
    throw new Error(`Worksheet clearing failed: ${error.message}`);
  }
};
