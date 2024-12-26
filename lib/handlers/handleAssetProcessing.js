import { createTeamsUpdate } from "../utils/utils.js";
import { getClientJobDirectories } from "../utils/fileStorageAndRetrieval.js";
import { createClientSelectionCard } from "../utils/adaptiveCards.js";

/**
 * This function handles the user selecting a job to process.
 * It creates a job selection card and sends it to the user.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export const handleProcessClientSelected = async (context) => {
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

  // returns all the directories inside each client's "Evidence Pack" directory
  try {
    const jobDirectories = await getClientJobDirectories(
      selectedDirectoryName,
      selectedDirectoryID
    );

    // Create the Jobs directory selection card and return the action "processJobSelected"
    const jobDirectorySelectionCard = createClientSelectionCard(
      jobDirectories,
      "Choose a Job"
    );

    // Send the Jobs directory selection card to the user
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
        `Unable to process evidence pack: No "Evidence Packs" folder found in **${selectedDirectoryName}**`,
        "",
        "⚠️",
        "warning"
      );
    } else {
      // For other errors, rethrow them to be handled by the global error handler
      throw error;
    }
  }
};

/**
 * This function handles the call to the Evidence Pack Processing Agent.
 * It sends the selected client directory to the agent.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export const handleProcessJobSelectedCallAgent = async (context) => {
  // parse the selected client directory from the context
  const selectedClientDirectory = JSON.parse(
    context.activity.value.directoryChoice
  );
  // Get the name of the selected client directory
  const selectedDirectoryName = selectedClientDirectory.name;
  const selectedDirectoryID = selectedClientDirectory.id;
  console.log(
    `Sending ${selectedDirectoryName} ID: ${selectedDirectoryID} to Evidence Pack Processing Agent`
  );

  await createTeamsUpdate(
    context,
    `Querying Agent...`,
    `Job: ${selectedDirectoryName}`,
    "🤖",
    "default"
  );

  const messageDetails = {
    serviceUrl: context.activity.serviceUrl,
    conversationId: context.activity.conversation.id,
    channelId: context.activity.channelId,
    tenantId: context.activity.channelData?.tenant?.id,
  };

  // TODO: Make a POST request to the Processing Agent
  // pass jobID(selectedDirectoryName), messageDetails

  console.log("Posting to Processing Agent", {
    messageDetails,
    selectedDirectoryName,
  });
};
