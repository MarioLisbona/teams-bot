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
    const client = await getGraphClient(); // Initialize the Graph client

    const clear = {
      applyTo: "Contents", // Specify that we want to clear the contents
    };

    for (const range of ranges) {
      const apiEndpoint = `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${worksheetName}/range(address='${range}')/clear`;
      await client.api(apiEndpoint).post(clear);
      console.log(
        `Cleared contents of range ${range} in worksheet ${worksheetName}.`
      );
    }
  } catch (error) {
    console.error("Error clearing worksheet range:", error.message);
    console.error("Full error details:", error);
  }
};
