import {
  createWorkflow1ValidationCard,
  createValidationProgressCard,
  createValidationCompletionCard,
} from "../utils/adaptiveCards.js";
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
