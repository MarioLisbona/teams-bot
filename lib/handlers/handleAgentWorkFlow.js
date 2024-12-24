import {
  createWorkflow1ValidationCard,
  createValidationProgressCard,
  createValidationCompletionCard,
  createValidateSignaturesCard,
  createApproveRejectImagesCard,
  createWorkflowProgressCard,
} from "../utils/adaptiveCards.js";
import { createTeamsUpdate } from "../utils/utils.js";

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
 * Creates and sends an adaptive card with multiple images of signatures for review/approval.
 * @param {Object} adapter - The bot adapter for managing conversations
 * @param {string} message - The message to display with the images
 * @param {string} serviceUrl - The service URL for the Teams conversation
 * @param {string} conversationId - The unique identifier for the conversation
 * @param {string} channelId - The Teams channel identifier
 * @param {string} tenantId - The Teams tenant identifier
 * @param {string[]} images - Array of image URLs to display for review
 * @returns {Promise<void>}
 */
export async function handleValidateSignatures(
  adapter,
  message,
  serviceUrl,
  conversationId,
  channelId,
  tenantId,
  images
) {
  // Create the validate signatures card with the images and a user comment input
  const validateSignaturesCard = createValidateSignaturesCard(message, images);

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
            content: validateSignaturesCard,
          },
        ],
      });
    }
  );
}

export async function handleWorkflowProgress(
  adapter,
  serviceUrl,
  conversationId,
  channelId,
  tenantId,
  isCompleted,
  workflowStep,
  jobId
) {
  console.log("handleWorkflowProgress function called");

  const workflowProgressCard = createWorkflowProgressCard(
    jobId,
    workflowStep,
    isCompleted
  );

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
            content: workflowProgressCard,
          },
        ],
      });
    }
  );
}

export async function handleSubmitSigReview(context) {
  console.log("handleSubmitSigReview function called");
  const reviewComment = context.activity.value.reviewComment;
  console.log({ reviewComment });

  // Delete the original review card
  await context.deleteActivity(context.activity.replyToId);

  await createTeamsUpdate(
    context,
    "Posting review to the Agent...",
    "🔄",
    "default"
  );

  // Simulate processing
  await new Promise((resolve) => setTimeout(resolve, 5000));

  await createTeamsUpdate(context, "Review posted to the Agent.", "✅", "good");
}
