import {
  createWorkflow1ValidationCard,
  createValidationProgressCard,
  createValidationCompletionCard,
  createReviewImagesCard,
} from "../utils/adaptiveCards.js";

/**
 * Handles the initial workflow validation process by creating and sending a validation card
 * to the specified conversation.
 * @param {Object} adapter - The bot adapter for managing conversations
 * @param {string} serviceUrl - The service URL for the Teams conversation
 * @param {string} conversationId - The unique identifier for the conversation
 * @param {string} channelId - The Teams channel identifier
 * @param {string} tenantId - The Teams tenant identifier
 * @param {Object} validations - The validation states for different document types
 * @param {string} jobId - The unique identifier for the current job
 * @returns {Promise<void>}
 */
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

/**
 * Handles a single validation step in the workflow process. Updates the UI with progress
 * and completion cards, and manages the validation state.
 * @param {Object} context - The turn context from the Teams activity
 * @param {Object} context.activity - The activity object containing validation details
 * @param {Object} context.activity.value - The values passed from the adaptive card
 * @param {string} context.activity.value.documentType - The type of document being validated
 * @param {Object} context.activity.value.currentValidations - Current state of all validations
 * @param {string} context.activity.value.jobId - The unique identifier for the current job
 * @returns {Promise<void>}
 */
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

/**
 * Creates and sends an adaptive card with multiple images for review/approval.
 * @param {Object} adapter - The bot adapter for managing conversations
 * @param {string} message - The message to display with the images
 * @param {string} serviceUrl - The service URL for the Teams conversation
 * @param {string} conversationId - The unique identifier for the conversation
 * @param {string} channelId - The Teams channel identifier
 * @param {string} tenantId - The Teams tenant identifier
 * @param {string[]} images - Array of image URLs to display for review
 * @returns {Promise<void>}
 */
export async function handleValidateImages(
  adapter,
  message,
  serviceUrl,
  conversationId,
  channelId,
  tenantId,
  images
) {
  // Create initial card with images and approve/reject buttons
  const reviewCard = createReviewImagesCard(message, images);

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

/**
 * Handles the approval or rejection of images, updating the UI with processing
 * and completion cards.
 * @param {Object} context - The turn context from the Teams activity
 * @param {Object} context.activity - The activity object containing action details
 * @param {Object} context.activity.value - The values passed from the adaptive card
 * @param {('approve'|'reject')} context.activity.value.action - The action selected by the user
 * @param {string[]} context.activity.value.images - Array of image URLs that were reviewed
 * @param {string} context.activity.value.message - The original message displayed with the images
 * @returns {Promise<void>}
 */
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