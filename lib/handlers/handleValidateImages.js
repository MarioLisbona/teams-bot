export async function handleValidateImages(
  adapter,
  message,
  serviceUrl,
  conversationId,
  channelId,
  tenantId,
  images
) {
  // Helper function to split array into chunks
  function chunk(array, size) {
    const chunked = [];
    for (let i = 0; i < array.length; i += size) {
      chunked.push(array.slice(i, i + size));
    }
    return chunked;
  }

  // Create initial card with images and approve/reject buttons
  const reviewCard = {
    type: "AdaptiveCard",
    version: "1.0",
    body: [
      {
        type: "TextBlock",
        text: message,
        size: "medium",
        weight: "bolder",
      },
      ...chunk(images, 3).map((imageChunk) => ({
        type: "ColumnSet",
        columns: imageChunk.map((url) => ({
          type: "Column",
          width: "stretch",
          items: [
            {
              type: "Image",
              url: url,
              size: "stretch",
              height: "200px",
            },
          ],
        })),
      })),
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "✅ Approve",
        style: "positive",
        data: {
          action: "approve",
          images: images,
          message: message,
        },
      },
      {
        type: "Action.Submit",
        title: "❌ Reject",
        style: "destructive",
        data: {
          action: "reject",
          images: images,
          message: message,
        },
      },
    ],
  };

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
            content: reviewCard,
          },
        ],
      });
    }
  );
}
