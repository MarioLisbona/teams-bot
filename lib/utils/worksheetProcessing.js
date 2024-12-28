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
 * Copies a worksheet to a new client workbook in SharePoint.
 * @param {Object} client - The Microsoft Graph client instance.
 * @param {Function} client.api - Function to make Graph API calls.
 * @param {string} sourceWorkbookId - The source workbook identifier.
 * @param {string} sourceWorksheetName - The source worksheet name.
 * @param {string} clientName - The client name for the new workbook.
 * @param {string} targetDirectoryId - The target directory identifier.
 * @returns {Promise<Object>} Object containing new workbook details.
 * @property {string} newWorkbookId - The ID of the created workbook.
 * @property {string} newWorkbookName - The name of the created workbook.
 * @throws {Error} If client is invalid.
 * @throws {Error} If source workbook access fails.
 * @throws {Error} If file copy fails.
 * @throws {Error} If data write fails.
 */
export const copyWorksheetToClientWorkbook = async (
  client,
  sourceWorkbookId,
  sourceWorksheetName,
  clientName,
  targetDirectoryId
) => {
  try {
    // Validate inputs
    if (!client?.api) {
      throw new Error("Invalid Graph API client");
    }
    if (!sourceWorkbookId) {
      throw new Error("Source workbook ID is required");
    }
    if (!sourceWorksheetName) {
      throw new Error("Source worksheet name is required");
    }
    if (!clientName) {
      throw new Error("Client name is required");
    }
    if (!targetDirectoryId) {
      throw new Error("Target directory ID is required");
    }
    if (!process.env.SHAREPOINT_SITE_ID) {
      throw new Error("SharePoint site ID is not configured");
    }
    if (!process.env.RFI_CLIENT_TEMPLATE_ID) {
      throw new Error("RFI client template ID is not configured");
    }

    const newWorksheetName = "RFI Responses";
    const templateWorkbookId = process.env.RFI_CLIENT_TEMPLATE_ID;

    console.log("Starting workbook copy process:", {
      templateId: templateWorkbookId,
      sourceId: sourceWorkbookId,
      targetDir: targetDirectoryId,
    });

    try {
      // Create new workbook from template
      const { newWorkbookId, newWorkbookName } = await copyFileInOneDrive(
        templateWorkbookId,
        `RFI Responses - ${clientName}.xlsx`,
        targetDirectoryId
      );

      if (!newWorkbookId) {
        throw new Error("Failed to create new workbook");
      }

      console.log("New workbook created:", {
        id: newWorkbookId,
        name: newWorkbookName,
      });

      try {
        // Fetch source data
        console.log("Fetching source worksheet data...");
        const existingData = await client
          .api(
            `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${sourceWorkbookId}/workbook/worksheets/${sourceWorksheetName}/usedRange`
          )
          .get();

        if (!existingData?.values) {
          throw new Error("No data found in source worksheet");
        }

        const cellValuesData = existingData.values;
        console.log(`Found ${cellValuesData.length} rows to copy`);

        // Calculate target range
        const newRangeAddress = getCellRange(cellValuesData, "A1", true);
        console.log(`Target range: ${newRangeAddress}`);

        // Try primary worksheet
        try {
          await client
            .api(
              `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${newWorkbookId}/workbook/worksheets('${newWorksheetName}')/range(address='${newRangeAddress}')`
            )
            .patch({
              values: cellValuesData,
            });
          console.log(`Data written to ${newWorksheetName}`);
        } catch (primaryError) {
          console.warn(
            `Failed to write to ${newWorksheetName}, trying Sheet1:`,
            primaryError.message
          );

          // Try fallback worksheet
          try {
            await client
              .api(
                `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${newWorkbookId}/workbook/worksheets('Sheet1')/range(address='${newRangeAddress}')`
              )
              .patch({
                values: cellValuesData,
              });
            console.log("Data written to Sheet1");
          } catch (fallbackError) {
            console.error("Failed to write to Sheet1:", fallbackError);
            throw new Error(
              `Failed to write data to workbook: ${fallbackError.message}`
            );
          }
        }

        return { newWorkbookId, newWorkbookName };
      } catch (error) {
        console.error("Data copy failed:", error);
        throw new Error(`Failed to copy worksheet data: ${error.message}`);
      }
    } catch (error) {
      console.error("Workbook creation failed:", error);
      throw new Error(`Failed to create new workbook: ${error.message}`);
    }
  } catch (error) {
    console.error("Worksheet copy failed:", {
      error: error.message,
      source: sourceWorkbookId,
      client: clientName,
      stack: error.stack,
    });
    throw new Error(`Failed to copy worksheet: ${error.message}`);
  }
};
