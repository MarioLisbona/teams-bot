import { processTestingWorksheet } from "../utils/auditProcessing.js";
import { createTeamsUpdate } from "../utils/utils.js";

export const handleTestingWorkbookSelected = async (context) => {
  // Get the file data and directory name from the context
  const fileData = JSON.parse(context.activity.value.fileChoice);
  const directoryName = context.activity.value.directoryName;

  // Delete the original card
  await context.deleteActivity(context.activity.replyToId);

  // Send a simple message using createTeamsUpdate
  await createTeamsUpdate(context, `📋 Selected file: **${fileData.name}**`);

  for (let i = 0; i < 15; i++) {
    console.log(`Waiting for ${15 - i} seconds...`);
    await new Promise((resolve) => setTimeout(resolve, 1000));
  }

  // Create a combined data object with all necessary information
  const combinedFileData = {
    ...fileData,
    directoryName: directoryName,
  };

  try {
    // Process the Testing worksheet with the combined data
    // All status updates are handled by createTeamsUpdate in processTestingWorksheet
    await processTestingWorksheet(context, combinedFileData);
  } catch (error) {
    console.error("Error processing worksheet:", error);
    await context.sendActivity(
      `❌ Error processing ${fileData.name}: ${error.message}`
    );
  }
};
