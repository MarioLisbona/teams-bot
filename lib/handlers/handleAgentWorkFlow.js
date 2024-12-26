import {
  createValidateSignaturesCard,
  createWorkflowProgressNotificationCard,
  createHumanValidationStepsCard,
} from "../utils/adaptiveCards.js";
import { createTeamsUpdate } from "../utils/utils.js";

/**
 * Updates the workflow progress by sending an adaptive card to the conversation.
 *
 * @param {Object} adapter - The bot adapter for managing conversations.
 * @param {string} serviceUrl - The service URL for the Teams conversation.
 * @param {string} conversationId - The unique identifier for the conversation.
 * @param {string} channelId - The Teams channel identifier.
 * @param {string} tenantId - The Teams tenant identifier.
 * @param {boolean} isComplete - Indicates if the workflow step is complete.
 * @param {Object} workflowStep - The workflow step to update.
 * @param {string} jobId - The job identifier for the workflow.
 * @returns {Promise<void>} - A promise that resolves when the workflow progress is updated.
 */
export async function workflowProgressNotification(
  adapter,
  serviceUrl,
  conversationId,
  channelId,
  tenantId,
  isComplete,
  workflowStep,
  jobId
) {
  const workflowProgressCard = createWorkflowProgressNotificationCard(
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
export async function validateSignatures(
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

/**
 * Handles the validation of signatures by processing the review comment and updating the Teams conversation.
 *
 * @param {Object} context - The context object containing the activity from the Teams bot.
 *
 */
export async function handleValidateSignatures(context) {
  const reviewComment = context.activity.value.reviewComment;
  console.log({ reviewComment });

  // Delete the original review card
  await context.deleteActivity(context.activity.replyToId);

  await createTeamsUpdate(
    context,
    "Posting review to the Agent...",
    `"${reviewComment}"`,
    "🔄",
    "emphasis"
  );

  // Simulate processing
  await new Promise((resolve) => setTimeout(resolve, 5000));

  await createTeamsUpdate(
    context,
    "Review posted to the Agent.",
    "",
    "✅",
    "good"
  );
}

/**
 * Handles the initial workflow validation process by creating and sending a validation card
 * to the specified conversation.
 * @param {Object} adapter - The bot adapter for managing conversations
 * @param {string} serviceUrl - The service URL for the Teams conversation
 * @param {string} conversationId - The unique identifier for the conversation
 * @param {string} channelId - The Teams channel identifier
 * @param {string} tenantId - The Teams tenant identifier
 * @param {Object} validationsRequired - The validation states for different document types
 * @param {string} jobId - The unique identifier for the current job
 * @returns {Promise<void>}
 */
export async function humanValidationSteps(
  adapter,
  serviceUrl,
  conversationId,
  channelId,
  tenantId,
  validationsRequired,
  jobId
) {
  const validationCard = createHumanValidationStepsCard(
    validationsRequired,
    {},
    jobId
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
 * This function handles the human validation steps in the workflow.
 * It takes the context as a parameter and performs the necessary actions.
 * @param {Object} context - The context object for the current conversation
 * @returns {Promise<void>}
 */

export async function handleHumanValidationSteps(context) {
  const { validationType, currentValidations, completedValidations, jobId } =
    context.activity.value;
  console.log(`Validation requested for step: ${validationType}`);

  // Delete the original validation card
  await context.deleteActivity(context.activity.replyToId);

  await createTeamsUpdate(
    context,
    `Validation in progress...`,
    `Step: ${validationType}`,
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
    `Validation completed`,
    `Step: ${validationType}`,
    "✅",
    "good"
  );

  // Send updated checklist
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: createHumanValidationStepsCard(
          currentValidations,
          updatedCompletedValidations,
          jobId
        ),
      },
    ],
  });
}
