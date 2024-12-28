import {
  createValidateSignaturesCard,
  createWorkflowProgressNotificationCard,
  createHumanValidationStepsCard,
} from "../utils/adaptiveCards.js";
import { createTeamsUpdate } from "../utils/utils.js";

/**
 * Updates the user with the workflow progress by sending an adaptive card to the conversation.
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
  try {
    // Create the workflow progress card
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
        try {
          await turnContext.sendActivity({
            attachments: [
              {
                contentType: "application/vnd.microsoft.card.adaptive",
                content: workflowProgressCard,
              },
            ],
          });
        } catch (error) {
          console.error("Failed to send activity:", error);
          throw new Error(
            `Failed to send workflow progress notification: ${error.message}`
          );
        }
      }
    );
  } catch (error) {
    console.error("Workflow progress notification failed:", error);
    throw error; // Re-throw to allow handling by the caller
  }
}

/**
 * Creates and sends an adaptive card with multiple images of signatures for review/approval.
 * @param {Object} adapter - The bot adapter for managing conversations.
 * @param {string} message - The message to display with the images.
 * @param {string} serviceUrl - The service URL for the Teams conversation.
 * @param {string} conversationId - The unique identifier for the conversation.
 * @param {string} channelId - The Teams channel identifier.
 * @param {string} tenantId - The Teams tenant identifier.
 * @param {string[]} images - Array of image URLs to display for review.
 * @returns {Promise<void>} - A promise that resolves when the adaptive card is sent.
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
  try {
    // Create the validate signatures card with the images and a user comment input
    const validateSignaturesCard = createValidateSignaturesCard(
      message,
      images
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
        try {
          await turnContext.sendActivity({
            attachments: [
              {
                contentType: "application/vnd.microsoft.card.adaptive",
                content: validateSignaturesCard,
              },
            ],
          });
        } catch (error) {
          console.error("Failed to send signature validation card:", error);
          throw new Error(
            `Failed to send signature validation activity: ${error.message}`
          );
        }
      }
    );
  } catch (error) {
    console.error("Signature validation process failed:", error);
    throw error; // Re-throw to allow handling by the caller
  }
}

/**
 * Processes the review comment and updates the Teams conversation as part of the signature validation process.
 *TODO: POST the user's review comment to the Agent
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export async function handleValidateSignatures(context) {
  try {
    const reviewComment = context.activity.value.reviewComment;

    try {
      // Delete the original review card
      await context.deleteActivity(context.activity.replyToId);
    } catch (error) {
      console.error("Failed to delete original review card:", error);
      // Continue execution even if deletion fails
    }

    try {
      // Notify the user that the review is being posted to the Agent
      await createTeamsUpdate(
        context,
        "Posting review to the Agent...",
        `"${reviewComment}"`,
        "🔄",
        "emphasis"
      );

      // TODO: POST the user's review comment to the Agent
      // Simulate processing
      await new Promise((resolve) => setTimeout(resolve, 5000));

      // Notify the user that the review has been posted to the Agent
      await createTeamsUpdate(
        context,
        "Review posted to the Agent.",
        "",
        "✅",
        "good"
      );
    } catch (error) {
      console.error("Failed to process review:", error);
      // Send error notification to user
      await createTeamsUpdate(
        context,
        "Failed to process review",
        "Please try again later",
        "❌",
        "attention"
      );
      throw error;
    }
  } catch (error) {
    console.error("Signature validation handler failed:", error);
    throw error;
  }
}

/**
 * Initiates the human validation for workflow steps that require it.
 * A validation card is sent to the specified conversation.
 *
 * @param {Object} adapter - The bot adapter responsible for managing conversations.
 * @param {string} serviceUrl - The service URL associated with the Teams conversation.
 * @param {string} conversationId - The unique identifier for the conversation.
 * @param {string} channelId - The identifier for the Teams channel.
 * @param {string} tenantId - The identifier for the Teams tenant.
 * @param {Object} validationsRequired - An object containing the validation states for different document types.
 * @param {string} jobId - The unique identifier for the current job.
 * @returns {Promise<void>} - A promise that resolves when the validation card is sent.
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
  // Create the validation card
  const validationCard = createHumanValidationStepsCard(
    validationsRequired,
    {},
    jobId
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
            content: validationCard,
          },
        ],
      });
    }
  );
}

/**
 * This function handles the human validation steps in the workflow.
 * It is triggered when the user selects a validation button for a step in the validation card.
 *TODO: Trigger the necessary human validation logic for each step and POST the results to the Agent
 * @param {Object} context - The context object for the current conversation
 * @returns {Promise<void>}
 */

export async function handleHumanValidationSteps(context) {
  // Extract the validation type, current validations, completed validations, and job ID from the activity value
  const { validationType, currentValidations, completedValidations, jobId } =
    context.activity.value;

  // Log the validation type
  console.log(`Validation requested for step: ${validationType}`);

  // Delete the original validation card
  await context.deleteActivity(context.activity.replyToId);

  // `Notify the user that the validation is in progress`
  await createTeamsUpdate(
    context,
    `Validation in progress...`,
    `Step: ${validationType}`,
    "🔄",
    "emphasis"
  );

  // *TODO: Trigger the necessary human validation logic for each step and POST the results to the Agent
  // Simulate validation process
  await new Promise((resolve) => setTimeout(resolve, 3000));

  // TODO: Only update the completed validations if the validation is successful

  // Update completed validations
  const updatedCompletedValidations = {
    ...completedValidations,
    [validationType]: true,
  };

  // Create a Teams update to notify the user that the validation has been completed
  // TODO Create update for failed validation
  await createTeamsUpdate(
    context,
    `Validation completed`,
    `Step: ${validationType}`,
    "✅",
    "good"
  );

  // Send updated validation card with the remaining validations to be completed
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
