import {
  processTesting,
  updateRfiWorksheet,
  copyWorksheetToClientWorkbook,
} from "./worksheetProcessing.js";
import { getGraphClient } from "../auth/msAuth.js";
import { createTeamsUpdate } from "./utils.js";

/**
 * Processes a Testing worksheet and generates an RFI Response workbook.
 *
 * @description
 * Handles the complete RFI processing workflow:
 * 1. Extracts RFI data from Testing worksheet
 * 2. Updates RFI Spreadsheet with extracted data
 * 3. Creates new client-specific RFI Response workbook
 * 4. Provides progress updates through Teams messages
 *
 * @param {Object} context - Teams bot turn context
 * @param {Function} context.sendActivity - Method to send Teams messages
 * @param {Object} selectedFileData - Selected file information
 * @param {string} selectedFileData.id - Workbook ID in SharePoint
 * @param {string} selectedFileData.directoryId - Parent directory ID
 * @param {string} selectedFileData.directoryName - Client directory name
 * @param {string} selectedFileData.name - Original file name
 *
 * @throws {Error} When failing to process Testing worksheet
 * @throws {Error} When failing to update RFI spreadsheet
 * @throws {Error} When failing to create new workbook
 * @returns {Promise<string>} Name of the newly created RFI Response workbook
 *
 * @example
 * processRfiWorksheet(context, {
 *   id: "123",
 *   directoryId: "456",
 *   directoryName: "Client A",
 *   name: "Testing.xlsx"
 * })
 * // Returns "Client A - RFI Response.xlsx"
 */
export const processRfiWorksheet = async (context, selectedFileData) => {
  try {
    // Get the file data and directory name from the context
    const workbookId = selectedFileData.id;
    const directoryId = selectedFileData.directoryId;
    const directoryName = selectedFileData.directoryName;

    console.log("Processing with directory ID:", directoryId);
    console.log("Processing with directory name:", directoryName);

    try {
      // Create a Graph client with caching disabled
      const client = await getGraphClient({ cache: false });

      // The name of the worksheet to find in the client workbook
      const testingSheetName = "Testing";
      const clientName = directoryName;

      await createTeamsUpdate(
        context,
        `Processing Testing worksheet...`,
        "",
        "⏳",
        "default"
      );

      try {
        // Process the testing sheet and return the updated RFI cell data
        const updatedRfiCellData = await processTesting(
          client,
          workbookId,
          testingSheetName
        );

        await createTeamsUpdate(
          context,
          `RFI data processed`,
          "",
          "♺",
          "default"
        );

        // only update the RFI spreadsheet, copy and email if there is RFI data to process
        if (updatedRfiCellData.length > 0) {
          try {
            await createTeamsUpdate(
              context,
              `Updating RFI Spreadsheet...`,
              "",
              "⚙️",
              "default"
            );

            // Update the RFI Spreadsheet worksheet
            await updateRfiWorksheet(
              client,
              workbookId,
              "RFI Spreadsheet",
              updatedRfiCellData
            );

            await createTeamsUpdate(
              context,
              `Copying RFI Spreadsheet to new workbook...`,
              "",
              "🛠️",
              "default"
            );

            // Copy the data to a new workbook
            const { newWorkbookId, newWorkbookName } =
              await copyWorksheetToClientWorkbook(
                client,
                workbookId,
                "RFI Spreadsheet",
                clientName,
                directoryId
              );

            await createTeamsUpdate(
              context,
              `RFI Spreadsheet copied to new workbook:`,
              newWorkbookName,
              "✅",
              "good"
            );

            return newWorkbookName;
          } catch (error) {
            console.error("Failed to update or copy RFI spreadsheet:", error);
            throw new Error(
              `Failed to create RFI response workbook: ${error.message}`
            );
          }
        } else {
          await createTeamsUpdate(
            context,
            `No RFI data found to process`,
            "",
            "ℹ️",
            "default"
          );
        }
      } catch (error) {
        console.error("Failed to process testing worksheet:", error);
        throw error;
      }
    } catch (error) {
      console.error("Failed to initialize processing:", error);
      throw error;
    }
  } catch (error) {
    console.error("Processing error:", error);
    try {
      await createTeamsUpdate(
        context,
        `An error occurred while processing ${selectedFileData.name}: ${error.message}`,
        "",
        "❌",
        "attention"
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }
    return; // Don't rethrow to prevent multiple error messages
  }
};
