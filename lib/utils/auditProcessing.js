import { TurnContext } from "botbuilder";
import {
  processTesting,
  updateRfiWorksheet,
  copyWorksheetToClientWorkbook,
} from "./worksheetProcessing.js";
import { getGraphClient } from "../auth/msAuth.js";

export const processTestingWorksheet = async (context, selectedFileData) => {
  const workbookId = selectedFileData.id;
  const directoryId = selectedFileData.directoryId;
  const directoryName = selectedFileData.directoryName;

  console.log("Processing with directory ID:", directoryId);
  console.log("Processing with directory name:", directoryName);

  // Create a Graph client with caching disabled
  const client = await getGraphClient({ cache: false });

  const testingSheetName = "Testing";

  try {
    // Initial notification
    await context.sendActivity({
      type: "message",
      textFormat: "markdown",
      text: `⏳ Processing Testing worksheet in **${selectedFileData.name}**...`,
    });

    // Use the directory name as the client name
    const clientName = directoryName;

    // Process the testing sheet and return the updated RFI cell data
    const updatedRfiCellData = await processTesting(
      client,
      workbookId,
      testingSheetName
    );
    await context.sendActivity({
      type: "message",
      textFormat: "markdown",
      text: `⚙️ RFI data processed for **${selectedFileData.name}**`,
    });

    // only update the RFI spreadsheet, copy ane email if there is RFI data to process
    if (updatedRfiCellData.length > 0) {
      // Update the RFI Spreadsheet worksheet in the same workbook the Testing worksheet is in
      await updateRfiWorksheet(
        client,
        workbookId,
        "RFI Spreadsheet",
        updatedRfiCellData
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
