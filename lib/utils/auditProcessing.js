import {
  processTesting,
  updateRfiWorksheet,
  copyWorksheetToClientWorkbook,
} from "./worksheetProcessing.js";
import { getGraphClient } from "../auth/msAuth.js";
import { createTeamsUpdate } from "./utils.js";

/**
 * This function processes the Testing worksheet.
 * It processes the Testing worksheet, extracting all RFI data into the RFI Spreadsheet in the same workbook
 * It copys that worksheet to a new RFI Response workbook, ready to be emailed to a client and sends a message to the user.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 * @param {Object} selectedFileData - The data of the selected file.
 * @returns {string} - The name of the new workbook.
 */
export const processTestingWorksheet = async (context, selectedFileData) => {
  // Get the file data and directory name from the context
  const workbookId = selectedFileData.id;
  const directoryId = selectedFileData.directoryId;
  const directoryName = selectedFileData.directoryName;

  console.log("Processing with directory ID:", directoryId);
  console.log("Processing with directory name:", directoryName);

  // Create a Graph client with caching disabled
  const client = await getGraphClient({ cache: false });

  const testingSheetName = "Testing";

  try {
    // Use the directory name as the client name
    const clientName = directoryName;

    // Initial notification
    await createTeamsUpdate(
      context,
      `Processing Testing worksheet in **${selectedFileData.name}**...`,
      "⏳",
      "default"
    );
    // Process the testing sheet and return the updated RFI cell data
    const updatedRfiCellData = await processTesting(
      client,
      workbookId,
      testingSheetName
    );
    await createTeamsUpdate(
      context,
      `RFI data processed for **${selectedFileData.name}**`,
      "♺",
      "default"
    );

    // only update the RFI spreadsheet, copy ane email if there is RFI data to process
    if (updatedRfiCellData.length > 0) {
      await createTeamsUpdate(
        context,
        `Updating RFI Spreadsheet in **${selectedFileData.name}**...`,
        "⚙️",
        "default"
      );

      // Update the RFI Spreadsheet worksheet in the same workbook the Testing worksheet is in
      await updateRfiWorksheet(
        client,
        workbookId,
        "RFI Spreadsheet",
        updatedRfiCellData
      );

      await createTeamsUpdate(
        context,
        `Copying RFI Spreadsheet to new workbook...`,
        "🛠️",
        "default"
      );

      // Copy the data in the updated RFI spreadsheet to a new workbook
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
        `RFI Spreadsheet copied to new workbook: **${newWorkbookName}**`,
        "✅",
        "good"
      );

      return newWorkbookName;
    }
  } catch (error) {
    console.error("Processing error:", error);

    await context.sendActivity({
      type: "message",
      textFormat: "markdown",
      text: `❌ An error occurred while processing **${selectedFileData.name}**. Please try again.`,
    });
  }
};
