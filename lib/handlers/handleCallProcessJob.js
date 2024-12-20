import { createTeamsUpdate } from "../utils/utils.js";

/**
 * This function handles the call to the Evidence Pack Processing Agent.
 * It sends the selected client directory to the agent.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export const handleCallProcessJobAgent = async (context) => {
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
    `Sending ${selectedDirectoryName} to Evidence Pack Processing Agent. `
  );

  try {
    // Extract message details from context
    const messageDetails = {
      serviceUrl: context.activity.serviceUrl,
      conversationId: context.activity.conversation.id,
      channelId: context.activity.channelId,
      tenantId: context.activity.channelData?.tenant?.id,
    };

    // Make POST request to the test endpoint
    const response = await fetch(
      "http://localhost:3978/api/test-teams-message",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          messageDetails,
          images: [
            "https://placecats.com/millie/300/200",
            "https://placecats.com/bella/300/200",
            "https://placecats.com/millie_neo/300/200",
            "https://placecats.com/millie/300/200",
            "https://placecats.com/bella/300/200",
            "https://placecats.com/millie_neo/300/200",
            "https://placecats.com/millie_neo/300/200",
          ],
          message: `Processing results for ${selectedDirectoryName}`,
          delay: 2,
        }),
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to send message: ${response.statusText}`);
    }
  } catch (error) {
    console.error("Error sending message:", error);
  }
};
