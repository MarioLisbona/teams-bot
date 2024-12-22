import { createWorkflow1ValidationCard } from "../utils/adaptiveCards.js";

export async function handleValidateWorkflow(
  adapter,
  serviceUrl,
  conversationId,
  channelId,
  tenantId,
  validations
) {
  console.log("handleValidateWorkflow function called");

  const validationCard = createWorkflow1ValidationCard(validations);

  // Create a reference to the conversation
  const conversationReference = {
    channelId: channelId,
    serviceUrl: serviceUrl,
    conversation: { id: conversationId },
    tenantId: tenantId,
  };

  // Use the adapter to continue the conversation and send the card
  await adapter.continueConversation(
    conversationReference,
    async (turnContext) => {
      await turnContext.sendActivity({
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: validationCard,
          },
        ],
      });
    }
  );
}
