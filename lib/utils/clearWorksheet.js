import { getGraphClient } from "../auth/msAuth.js"; // Import the authentication functions

/**
 * This function clears the contents of a specified range in a worksheet.
 * @param {string} workbookId - The ID of the workbook.
 * @param {string} worksheetName - The name of the worksheet.
 * @param {Array} ranges - The ranges to clear.
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
            `Cleared contents of range ${range} in worksheet ${worksheetName}.`
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
