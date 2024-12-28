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
 * Updates an Excel worksheet with RFI data, separating general and specific issues.
 * @param {Object} client - The Microsoft Graph API client instance.
 * @param {Function} client.api - Function to make Graph API calls.
 * @param {string} workbookId - The SharePoint workbook identifier.
 * @param {string} sheetName - The name of the worksheet to update.
 * @param {Array<Object>} rfiCellData - Array of RFI data objects to update.
 * @param {Array<Object>} rfiCellData[].projectsAffected - Array of affected projects.
 * @returns {Promise<void>} Promise that resolves when update is complete.
 * @throws {Error} If client is invalid.
 * @throws {Error} If worksheet access fails.
 * @throws {Error} If data update fails.
 */
export const updateRfiWorksheet = async (
  client,
  workbookId,
  sheetName,
  rfiCellData
) => {
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
    if (!Array.isArray(rfiCellData)) {
      throw new Error("RFI cell data must be an array");
    }
    if (!process.env.SHAREPOINT_SITE_ID) {
      throw new Error("SharePoint site ID is not configured");
    }

    try {
      // Clear existing data
      const ranges = ["C14:I34", "C42:I141"];
      await clearWorksheetRange(workbookId, sheetName, ranges);

      // Process and filter RFI data
      try {
        // Separate general and specific issues
        const generalIssuesRfi = rfiCellData.filter((group) => {
          if (!Array.isArray(group?.projectsAffected)) {
            console.warn("Invalid group structure found, skipping");
            return false;
          }
          return group.projectsAffected.length >= 4;
        });

        const specificIssuesRfi = rfiCellData.filter((group) => {
          if (!Array.isArray(group?.projectsAffected)) {
            console.warn("Invalid group structure found, skipping");
            return false;
          }
          return group.projectsAffected.length < 4;
        });

        // Prepare data for spreadsheet
        const generalIssuesRfiData =
          generalIssuesRfi.length > 0
            ? prepareRfiCellDataForRfiSpreadsheet(generalIssuesRfi)
            : [];

        const specificIssuesRfiData =
          specificIssuesRfi.length > 0
            ? prepareRfiCellDataForRfiSpreadsheet(specificIssuesRfi)
            : [];

        // Calculate ranges
        const startRowGeneralIssuesRfi = 14;
        const startRowSpecificIssuesRfi = 42;

        const { rangeForGeneralIssuesRfi, rangeForSpecificIssuesRfi } =
          getRfiRanges(
            startRowGeneralIssuesRfi,
            startRowSpecificIssuesRfi,
            generalIssuesRfiData.length,
            specificIssuesRfiData.length
          );

        console.log("Update ranges:", {
          general: rangeForGeneralIssuesRfi,
          specific: rangeForSpecificIssuesRfi,
        });

        // Construct URLs for updates
        const baseUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='`;
        const urlGeneralIssuesRfi = `${baseUrl}${rangeForGeneralIssuesRfi}')`;
        const urlSpecificIssuesRfi = `${baseUrl}${rangeForSpecificIssuesRfi}')`;

        // Update general issues
        if (generalIssuesRfiData.length > 0) {
          try {
            await updateExcelData(
              client,
              urlGeneralIssuesRfi,
              { values: generalIssuesRfiData },
              "general issues RFI"
            );
          } catch (error) {
            console.error("Failed to update general issues:", error);
            throw new Error(`General issues update failed: ${error.message}`);
          }
        }

        // Update specific issues
        if (specificIssuesRfiData.length > 0) {
          try {
            await updateExcelData(
              client,
              urlSpecificIssuesRfi,
              { values: specificIssuesRfiData },
              "specific issues RFI"
            );
          } catch (error) {
            console.error("Failed to update specific issues:", error);
            throw new Error(`Specific issues update failed: ${error.message}`);
          }
        }
      } catch (error) {
        console.error("RFI data processing failed:", error);
        throw new Error(`Failed to process RFI data: ${error.message}`);
      }
    } catch (error) {
      console.error("Worksheet update failed:", error);
      throw new Error(`Failed to update worksheet: ${error.message}`);
    }
  } catch (error) {
    console.error("RFI worksheet update failed:", {
      error: error.message,
      workbookId,
      sheetName,
      stack: error.stack,
    });
    throw new Error(`Failed to update RFI worksheet: ${error.message}`);
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
