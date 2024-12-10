import { createProcessingResponsesCard } from "../utils/adaptiveCards.js";

export async function handleProcessSelectedResponses(context) {
  const selectedFile = JSON.parse(context.activity.value.fileChoice);
  console.log("Processing RFI Client Responses:", selectedFile.name);

  // Update the card to show processing state
  await context.updateActivity({
    type: "message",
    id: context.activity.replyToId,
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: createProcessingResponsesCard(selectedFile.name),
      },
    ],
  });

  // Send message to Teams
  await context.sendActivity("Processing RFI Client Responses");
}
