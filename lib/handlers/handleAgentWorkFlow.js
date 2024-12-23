import {
  createWorkflow1ValidationCard,
  createValidationProgressCard,
  createValidationCompletionCard,
} from "../utils/adaptiveCards.js";

export async function handleValidateWorkflow(
  adapter,
  serviceUrl,
  conversationId,
  channelId,
  tenantId,
  validations,
  jobId
) {
  console.log("handleValidateWorkflow function called");

  const validationCard = createWorkflow1ValidationCard(validations, jobId);

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

export async function handleValidateWorkflowStep(context) {
  console.log("handleValidateWorkflowStep function called");
  const documentType = context.activity.value.documentType;
  const currentValidations = context.activity.value.currentValidations;
  const jobId = context.activity.value.jobId;
  console.log(`Validation requested for: ${documentType}`);

  // Delete the original validation card
  await context.deleteActivity(context.activity.replyToId);

  // Send an enhanced validation in progress card
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: createValidationProgressCard(documentType),
      },
    ],
  });

  // TODO: Add your validation logic here
  // Simulate validation process with timeout
  await new Promise((resolve) => setTimeout(resolve, 5000)).then(() =>
    console.log("Validation timeout completed")
  );

  // Update validations object with new status
  const updatedValidations = {
    ...currentValidations,
    [documentType]: true,
  };

  // Send enhanced completion message card
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: createValidationCompletionCard(documentType),
      },
    ],
  });

  // Send final updated validation card
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: createWorkflow1ValidationCard(updatedValidations, jobId),
      },
    ],
  });
}

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

export async function handleApproveRejectImages(context) {
  console.log("handleApproveRejectImages function called");
  const action = context.activity.value.action;
  const images = context.activity.value.images;
  const message = context.activity.value.message;

  console.log({ action, images, message });

  // Delete the original review card
  await context.deleteActivity(context.activity.replyToId);

  // Send processing card
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: {
          type: "AdaptiveCard",
          version: "1.0",
          body: [
            {
              type: "Container",
              style: "emphasis",
              items: [
                {
                  type: "ColumnSet",
                  columns: [
                    {
                      type: "Column",
                      width: "auto",
                      verticalContentAlignment: "center",
                      items: [
                        {
                          type: "TextBlock",
                          text: "🔄",
                          size: "extraLarge",
                          spacing: "none",
                        },
                      ],
                    },
                    {
                      type: "Column",
                      width: "stretch",
                      items: [
                        {
                          type: "TextBlock",
                          text: message,
                          weight: "bolder",
                          size: "medium",
                          color: "accent",
                        },
                        {
                          type: "TextBlock",
                          text: "Processing in Progress",
                          spacing: "none",
                          isSubtle: true,
                        },
                      ],
                    },
                  ],
                },
              ],
              padding: "default",
            },
          ],
        },
      },
    ],
  });

  // Simulate processing
  await new Promise((resolve) => setTimeout(resolve, 5000));

  // Send completion card
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: {
          type: "AdaptiveCard",
          version: "1.0",
          body: [
            {
              type: "Container",
              style: action === "approve" ? "good" : "attention",
              items: [
                {
                  type: "ColumnSet",
                  columns: [
                    {
                      type: "Column",
                      width: "auto",
                      verticalContentAlignment: "center",
                      items: [
                        {
                          type: "TextBlock",
                          text: action === "approve" ? "✅" : "❌",
                          size: "extraLarge",
                          spacing: "none",
                        },
                      ],
                    },
                    {
                      type: "Column",
                      width: "stretch",
                      items: [
                        {
                          type: "TextBlock",
                          text: message,
                          weight: "bolder",
                          size: "medium",
                        },
                        {
                          type: "TextBlock",
                          text: `${
                            action === "approve" ? "Approval" : "Rejection"
                          } Complete`,
                          spacing: "none",
                          color: action === "approve" ? "good" : "attention",
                        },
                      ],
                    },
                  ],
                },
              ],
              padding: "default",
            },
          ],
        },
      },
    ],
  });
}
