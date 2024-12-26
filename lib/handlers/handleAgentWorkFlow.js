import {
  createValidateSignaturesCard,
  createWorkflowProgressCard,
  createHumanWorkflowValidationCard,
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
export async function handleHumanWorkflowValidationUI(
  adapter,
  serviceUrl,
  conversationId,
  channelId,
  tenantId,
  validationsRequired,
  jobId
) {
  const validationCard = createHumanWorkflowValidationCard(
    jobId,
    validationsRequired
  );

  const conversationReference = {
    channelId: channelId,
    serviceUrl: serviceUrl,
    conversation: { id: conversationId },
    tenantId: tenantId,
  };

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
  isComplete,
  workflowStep,
  jobId
) {
  const workflowProgressCard = createWorkflowProgressCard(
    jobId,
    workflowStep,
    isComplete
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

export async function handleHumanValidation(context) {
  const { validationType, currentValidations, completedValidations, jobId } =
    context.activity.value;
  console.log(`Validation requested for step: ${validationType}`);

  // Delete the original validation card
  await context.deleteActivity(context.activity.replyToId);

  await createTeamsUpdate(
    context,
    `Validation in progress for step ${validationType}`,
    "🔄",
    "emphasis"
  );

  // Simulate validation process
  await new Promise((resolve) => setTimeout(resolve, 3000));

  // Update completed validations
  const updatedCompletedValidations = {
    ...completedValidations,
    [validationType]: true,
  };

  await createTeamsUpdate(
    context,
    `Validation completed for step ${validationType}`,
    "✅",
    "good"
  );

  // Send updated checklist
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: createHumanWorkflowValidationCard(
          currentValidations,
          updatedCompletedValidations,
          jobId
        ),
      },
    ],
  });
}
