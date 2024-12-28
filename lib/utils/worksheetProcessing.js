import { filterRowsForRFICells } from "./utils.js";
import {
  updateRfiCellData,
  prepareRfiCellDataForRfiSpreadsheet,
  getRfiRanges,
  updateExcelData,
  getCellRange,
} from "./utils.js";
import { clearWorksheetRange } from "./clearWorksheet.js";
import { copyFileInOneDrive } from "./copyFile.js";

/**
 * Processes an Excel worksheet to extract and update RFI information.
 * @param {Object} client - The Microsoft Graph API client instance.
 * @param {Function} client.api - Function to make Graph API calls.
 * @param {string} workbookId - The SharePoint workbook identifier.
 * @param {string} sheetName - The name of the worksheet to process.
 * @returns {Promise<Array<Object>>} Array of processed RFI data objects.
 * @property {string} rfi - The RFI text.
 * @property {string} cellReference - The Excel cell reference.
 * @property {string} iid - The item identifier.
 * @property {string} updatedRfi - The processed RFI text.
 * @throws {Error} If client is invalid.
 * @throws {Error} If worksheet access fails.
 * @throws {Error} If data processing fails.
 */
export const processTesting = async (client, workbookId, sheetName) => {
  try {
    // Validate inputs
    if (!client?.api) {
      throw new Error("Invalid Graph API client");
    }
    if (!workbookId) {
      throw new Error("Workbook ID is required");
    }
    if (!sheetName) {
      throw new Error("Sheet name is required");
    }
    if (!process.env.SHAREPOINT_SITE_ID) {
      throw new Error("SharePoint site ID is not configured");
    }

    try {
      // Construct the URL for the Excel file's used range
      const range = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/usedRange`;

      // Fetch worksheet data
      let response;
      try {
        response = await client.api(range).get();
        if (!response?.values) {
          throw new Error("Invalid response format from Graph API");
        }
      } catch (apiError) {
        console.error("Graph API request failed:", apiError);
        throw new Error(`Failed to fetch worksheet data: ${apiError.message}`);
      }

      // Process worksheet data
      try {
        const data = response.values;

        // Filter out empty rows
        const nonEmptyRows = data.filter((row) => {
          if (!Array.isArray(row)) {
            console.warn("Invalid row format found, skipping");
            return false;
          }
          return row.some((cell) => cell !== null && cell !== "");
        });

        console.log({
          "Total rows": data.length,
          "Non-empty rows": nonEmptyRows.length,
          "Workbook ID": workbookId,
        });

        // Extract RFI data
        let rfiCellData;
        try {
          rfiCellData = filterRowsForRFICells(nonEmptyRows);
          console.log("Rows with RFI Data:", rfiCellData.length);

          if (rfiCellData.length === 0) {
            console.warn("No RFI data found in worksheet");
          }
        } catch (filterError) {
          console.error("RFI cell filtering failed:", filterError);
          throw new Error(`Failed to filter RFI cells: ${filterError.message}`);
        }

        // Update RFI data
        try {
          const updatedRfiCellData = await updateRfiCellData(rfiCellData);
          if (!Array.isArray(updatedRfiCellData)) {
            throw new Error("Invalid RFI update result");
          }
          return updatedRfiCellData;
        } catch (updateError) {
          console.error("RFI data update failed:", updateError);
          throw new Error(`Failed to update RFI data: ${updateError.message}`);
        }
      } catch (processError) {
        console.error("Data processing failed:", processError);
        throw new Error(
          `Failed to process worksheet data: ${processError.message}`
        );
      }
    } catch (error) {
      console.error("Worksheet processing failed:", error);
      throw new Error(`Failed to process worksheet: ${error.message}`);
    }
  } catch (error) {
    console.error("Testing process failed:", {
      error: error.message,
      workbookId,
      sheetName,
      stack: error.stack,
    });
    throw new Error(`Testing process failed: ${error.message}`);
  }
};

/**
 * This function updates an Excel spreadsheet with new data.
 * @param {Object} client - The Microsoft Graph client instance.
 * @param {string} workbookId - The unique identifier for the workbook.
 * @param {string} sheetName - The name of the sheet within the workbook.
 * @param {Array} rfiCellData - The array of updated RFI cell data.
 * @returns {Promise<void>} - A promise that resolves when the update is complete.
 */
export const updateRfiWorksheet = async (
  client,
  workbookId,
  sheetName,
  rfiCellData
) => {
  // The array of ranges to clear - only clear cell data, not headings or images
  // This range is based off the RFI Spreadsheet in the main client workbook
  const ranges = ["C14:I34", "C42:I141"];

  // Clear the ranges before updating the RFI spreadsheet
  await clearWorksheetRange(workbookId, sheetName, ranges);

  // Filter groupedData into two arrays: one where projectsAffected.length >= 4, and one where projectsAffected.length < 4
  // This is done to separate the RFI's into general and specific issues
  const generalIssuesRfi =
    rfiCellData.filter((group) => group.projectsAffected.length >= 4).length > 0
      ? rfiCellData.filter((group) => group.projectsAffected.length >= 4)
      : []; // Assign only if at least one group has projectsAffected.length >= 4

  const specificIssuesRfi =
    rfiCellData.filter((group) => group.projectsAffected.length < 4).length > 0
      ? rfiCellData.filter((group) => group.projectsAffected.length < 4)
      : []; // Assign only if at least one group has projectsAffected.length < 4

  // Prepare data for both sets of groups
  const generalIssuesRfiData =
    generalIssuesRfi.length > 0
      ? prepareRfiCellDataForRfiSpreadsheet(generalIssuesRfi)
      : []; // Call only if generalIssuesRfi is not empty

  const specificIssuesRfiData =
    specificIssuesRfi.length > 0
      ? prepareRfiCellDataForRfiSpreadsheet(specificIssuesRfi)
      : []; // Call only if specificIssuesRfi is not empty

  // Define starting row for both cases
  const startRowGeneralIssuesRfi = 14; // Start from row 7 for general issues
  const startRowSpecificIssuesRfi = 42; // Start from row 18 for specific issues

  // Get the ranges for general and specific issues RFI for the update request
  const { rangeForGeneralIssuesRfi, rangeForSpecificIssuesRfi } = getRfiRanges(
    startRowGeneralIssuesRfi,
    startRowSpecificIssuesRfi,
    generalIssuesRfiData.length > 0 ? generalIssuesRfiData.length : 0, // Use length only if generalIssuesRfiData is not empty
    specificIssuesRfiData.length > 0 ? specificIssuesRfiData.length : 0 // Use length only if specificIssuesRfiData is not empty
  );

  console.log("Range for general issues RFI:", rangeForGeneralIssuesRfi);
  console.log("Range for specific issues RFI:", rangeForSpecificIssuesRfi);

  // Construct the URL for the Excel file's using ranges for general and specific issues RFI
  const urlGeneralIssuesRfi = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='${rangeForGeneralIssuesRfi}')`;
  const urlSpecificIssuesRfi = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='${rangeForSpecificIssuesRfi}')`;

  // Prepare the request body with the data to update only if generalIssuesRfiData is not empty
  if (generalIssuesRfiData.length > 0) {
    const requestBodyGeneralIssuesRfi = {
      values: generalIssuesRfiData,
    };

    // Call the new function for general issues RFI
    await updateExcelData(
      client,
      urlGeneralIssuesRfi,
      requestBodyGeneralIssuesRfi,
      "general issues RFI"
    );
  }

  // Prepare the request body with the data to update for specific issues RFI only if specificIssuesRfiData is not empty
  if (specificIssuesRfiData.length > 0) {
    const requestBodySpecificIssuesRfi = {
      values: specificIssuesRfiData,
    };

    // Call the new function for specific issues RFI
    await updateExcelData(
      client,
      urlSpecificIssuesRfi,
      requestBodySpecificIssuesRfi,
      "specific issues RFI"
    );
  }
};

/**
 * This function copies a worksheet to a new spreadsheet.
 * @param {Object} client - The Microsoft Graph client instance.
 * @param {string} sourceWorkbookId - The unique identifier for the source workbook.
 * @param {string} sourceWorksheetName - The name of the source worksheet.
 * @param {string} clientName - The name of the client.
 * @param {string} targetDirectoryId - The unique identifier for the target directory.
 * @returns {Promise<Object>} - A promise that resolves to an object containing the new workbook ID and name.
 */
export const copyWorksheetToClientWorkbook = async (
  client,
  sourceWorkbookId,
  sourceWorksheetName,
  clientName,
  targetDirectoryId
) => {
  const newWorksheetName = "RFI Responses";

  // Get the RFI Client Template workbook ID from the .env file
  const templateWorkbookId = process.env.RFI_CLIENT_TEMPLATE_ID;

  console.log("Template workbook ID:", templateWorkbookId);

  // Pass the targetDirectoryId to copyFileInOneDrive
  const { newWorkbookId, newWorkbookName } = await copyFileInOneDrive(
    templateWorkbookId,
    `RFI Responses - ${clientName}.xlsx`,
    targetDirectoryId
  );

  console.log("New workbook created with ID:", newWorkbookId);

  // Extract the data from the source worksheet
  console.log("Fetching data from source worksheet...");
  const existingData = await client
    .api(
      `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${sourceWorkbookId}/workbook/worksheets/${sourceWorksheetName}/usedRange`
    )
    .get();

  // Check if existingData is valid
  if (!existingData || !existingData.values) {
    throw new Error("No data found in the existing worksheet.");
  }

  const cellValuesData = existingData.values;
  console.log(`Found ${cellValuesData.length} rows of data to copy`);

  // Calculate the cell range for the data
  const newRangeAddress = getCellRange(cellValuesData, "A1", true);
  console.log(`Writing to range: ${newRangeAddress}`);

  // Try to write to the worksheet named "RFI Spreadsheet Template" first
  try {
    await client
      .api(
        `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${newWorkbookId}/workbook/worksheets('${newWorksheetName}')/range(address='${newRangeAddress}')`
      )
      .patch({
        values: cellValuesData,
      });
    console.log(`Data written successfully to ${newWorkbookName}`);
  } catch (error) {
    console.log("Error writing to RFI Spreadsheet Template:", error.message);
    // Try writing to Sheet1 as fallback
    try {
      await client
        .api(
          `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${newWorkbookId}/workbook/worksheets('Sheet1')/range(address='${newRangeAddress}')`
        )
        .patch({
          values: cellValuesData,
        });
      console.log("Data written successfully to Sheet1");
    } catch (fallbackError) {
      console.log("Error writing to Sheet1:", fallbackError.message);
      throw fallbackError;
    }
  }

  return { newWorkbookId, newWorkbookName };
};
