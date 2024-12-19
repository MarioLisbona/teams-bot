import { createTeamsUpdate } from "../utils/utils.js";
import { getClientJobDirectories } from "../utils/fileStorageAndRetrieval.js";
import { createClientSelectionCard } from "../utils/adaptiveCards.js";
export const handleProcessSelectJob = async (context) => {
  // parse the selected client directory from the context
  const selectedClientDirectory = JSON.parse(
    context.activity.value.directoryChoice
  );
  // Get the name of the selected client directory
  const selectedDirectoryName = selectedClientDirectory.name;
  const selectedDirectoryID = selectedClientDirectory.id;
  console.log(
    `Directory selected -> ${selectedDirectoryName} ID: ${selectedDirectoryID}`
  );

  try {
    const jobDirectories = await getClientJobDirectories(
      selectedDirectoryName,
      selectedDirectoryID
    );

    // Create the client selection card and return the action "processClientSelected"
    const jobDirectorySelectionCard = createClientSelectionCard(
      jobDirectories,
      "Choose a Job"
    );

    // Send the client selection card to the user
    await context.sendActivity({
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: jobDirectorySelectionCard,
        },
      ],
    });
  } catch (error) {
    if (error.message.includes("Evidence Packs")) {
      await createTeamsUpdate(
        context,
        `⚠️ Unable to process evidence pack: No "Evidence Packs" folder found in **${selectedDirectoryName}**`
      );
    } else {
      // For other errors, rethrow them to be handled by the global error handler
      throw error;
    }
  }
};
