import {
  createValidateSignaturesCard,
  createWorkflowProgressNotificationCard,
  createHumanValidationStepsCard,
} from "../utils/adaptiveCards.js";
import {
  createConversationReference,
  continueTeamsConversation,
  createTeamsUpdate,
} from "../utils/utils.js";

/**
 * Sends an adaptive card to update the user about workflow progress in a Teams conversation.
 *
 * @description
 * Creates and sends a progress notification card using the Microsoft Teams Bot Framework adapter.
 * The card displays the current workflow step status and related information.
 *
 * @param {Object} adapter - The Bot Framework adapter instance
 * @param {string} serviceUrl - Teams service URL for the conversation
 * @param {string} conversationId - Unique identifier for the Teams conversation
 * @param {string} channelId - Teams channel identifier where the message will be sent
 * @param {string} tenantId - Microsoft Teams tenant identifier
 * @param {boolean} isComplete - Whether the current workflow step is completed
 * @param {Object} workflowStep - Current workflow step information
 * @param {string} workflowStep.name - Name of the workflow step
 * @param {string} jobId - Unique identifier for the current job/process
 *
 * @throws {Error} When failing to send the adaptive card or continue the conversation
 * @returns {Promise<void>} Resolves when the progress card is successfully sent
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
    const conversationReference = createConversationReference(
      conversationId,
      channelId,
      serviceUrl,
      tenantId
    );

    // Use the adapter to continue the conversation and send the card
    await continueTeamsConversation(
      adapter,
      conversationReference,
      workflowProgressCard
    );
  } catch (error) {
    console.error("Workflow progress notification failed:", error);
    throw error; // Re-throw to allow handling by the caller
  }
}

/**
 * Creates and sends an adaptive card for signature validation in Teams.
 *
 * @description
 * Generates a card containing signature images for user review and approval.
 * Allows users to submit comments about the signatures through an input field.
 * Once a user comment is submitted the action "validateSignatures" is triggered
 * and the comment is sent to the Agent
 * @param {Object} adapter - The Bot Framework adapter instance
 * @param {string} message - Display message accompanying the signature images
 * @param {string} serviceUrl - Teams service URL for the conversation
 * @param {string} conversationId - Unique identifier for the Teams conversation
 * @param {string} channelId - Teams channel identifier where the message will be sent
 * @param {string} tenantId - Microsoft Teams tenant identifier
 * @param {string[]} images - Array of signature image URLs to display
 *
 * @throws {Error} When failing to send the signature validation card
 * @returns {Promise<void>} Resolves when the validation card is successfully sent
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
    const conversationReference = createConversationReference(
      conversationId,
      channelId,
      serviceUrl,
      tenantId
    );

    // Use the adapter to continue the conversation and send the card
    await continueTeamsConversation(
      adapter,
      conversationReference,
      validateSignaturesCard
    );
  } catch (error) {
    console.error("Signature validation process failed:", error);
    throw error; // Re-throw to allow handling by the caller
  }
}

/**
 * Processes user's signature validation response and updates the Teams conversation.
 *
 * @description
 * Handles the review comment submitted by the user, removes the original review card,
 * and sends status updates about the processing of the review.
 * The users comment is sent to the Workflow Agent as a message
 *
 * @param {Object} context - Teams bot turn context
 * @param {Object} context.activity - The incoming activity from Teams
 * @param {Object} context.activity.value - Values submitted through the adaptive card
 * @param {string} context.activity.value.reviewComment - User's review comment
 * @param {string} context.activity.replyToId - ID of the message being replied to
 *
 * @throws {Error} When failing to process the review or update the conversation
 * @returns {Promise<void>} Resolves when the review is processed and updates are sent
 *
 * @todo: POST the user's review comment to the Workflow Agent
 */
export async function handleValidateSignatures(context) {
  try {
    const reviewComment = context.activity.value.reviewComment;

    try {
      // Delete the original review card - avoids timeout errors on adaptive card
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
 * Initiates human validation steps for workflow processes in Teams.
 *
 * @description
 * Creates and sends an adaptive card that displays required validation steps
 * and allows users to interact with individual validation requirements by cliciing the validate button
 * The validate button triggers the action "humanValidation" and the validation for that step is triggered
 *
 * @param {Object} adapter - The Bot Framework adapter instance
 * @param {string} serviceUrl - Teams service URL for the conversation
 * @param {string} conversationId - Unique identifier for the Teams conversation
 * @param {string} channelId - Teams channel identifier where the message will be sent
 * @param {string} tenantId - Microsoft Teams tenant identifier
 * @param {Object} validationsRequired - Object containing validation requirements for different document types
 * @param {string} jobId - Unique identifier for the current workflow job
 *
 * @throws {Error} When failing to send the validation steps card
 * @returns {Promise<void>} Resolves when the validation card is successfully sent
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
  try {
    // Create the validation card
    const validationCard = createHumanValidationStepsCard(
      validationsRequired,
      {},
      jobId
    );

    // Create a reference to the conversation
    const conversationReference = createConversationReference(
      conversationId,
      channelId,
      serviceUrl,
      tenantId
    );

    // Use the adapter to continue the conversation and send the card
    await continueTeamsConversation(
      adapter,
      conversationReference,
      validationCard
    );
  } catch (error) {
    console.error("Human validation steps process failed:", error);
    throw error; // Re-throw to allow handling by the caller
  }
}

/**
 * Handles user interactions with human validation step cards.
 *
 * @description
 * Processes user responses to validation steps, updates the validation status,
 * and manages the conversation flow by removing old cards and sending updates.
 *
 * @param {Object} context - Teams bot turn context
 * @param {Object} context.activity - The incoming activity from Teams
 * @param {Object} context.activity.value - Values submitted through the adaptive card
 * @param {string} context.activity.value.validationType - Type of validation being performed
 * @param {Object} context.activity.value.currentValidations - Current validation requirements
 * @param {Object} context.activity.value.completedValidations - Status of completed validations
 * @param {string} context.activity.value.jobId - Associated job identifier
 *
 * @throws {Error} When failing to process validation steps or update the conversation
 * @returns {Promise<void>} Resolves when the validation step is processed and updates are sent
 *
 * @todo: Trigger the necessary human validation logic for each step and POST the results to the Agent
 */
export async function handleHumanValidationSteps(context) {
  try {
    const { validationType, currentValidations, completedValidations, jobId } =
      context.activity.value;

    // Log the validation type
    console.log(`Validation requested for step: ${validationType}`);

    try {
      // Delete the original validation card - avoids timeout errors on adaptive card
      await context.deleteActivity(context.activity.replyToId);
    } catch (error) {
      console.error("Failed to delete original validation card:", error);
      // Continue execution even if deletion fails
    }

    try {
      // Notify the user that the validation is in progress
      await createTeamsUpdate(
        context,
        `Validation in progress...`,
        `Step: ${validationType}`,
        "🔄",
        "emphasis"
      );

      // Simulate validation process
      await new Promise((resolve) => setTimeout(resolve, 500));

      // Update completed validations
      const updatedCompletedValidations = {
        ...completedValidations,
        [validationType]: true,
      };

      try {
        await createTeamsUpdate(
          context,
          `Validation completed`,
          `Step: ${validationType}`,
          "✅",
          "good"
        );

        // Check if all validations are complete
        const allValidationsComplete = Object.keys(currentValidations).every(
          (key) => updatedCompletedValidations[key] === true
        );

        if (allValidationsComplete) {
          try {
            // Show completion message
            await createTeamsUpdate(
              context,
              "All Validation Steps Complete",
              `Job ID: ${jobId}`,
              "🎉",
              "good"
            );
          } catch (completionError) {
            console.error(
              "Failed to send completion notification:",
              completionError
            );
            // Continue execution - the validation was still successful
          }
        } else {
          try {
            // Send updated validation card with remaining steps
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
          } catch (cardError) {
            console.error("Failed to send updated validation card:", cardError);
            throw cardError; // This is important as users need to see remaining steps
          }
        }
      } catch (validationCompleteError) {
        console.error(
          "Failed to send validation complete notification:",
          validationCompleteError
        );
        // Continue execution - the validation was still successful
      }
    } catch (progressError) {
      console.error(
        "Failed to send validation progress notification:",
        progressError
      );
      throw progressError; // This is important as users need to know if the validation started
    }
  } catch (error) {
    console.error("Human validation steps handler failed:", error);
    throw error;
  }
}
