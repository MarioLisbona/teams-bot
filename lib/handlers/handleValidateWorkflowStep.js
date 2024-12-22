import { createWorkflow1ValidationCard } from "../utils/adaptiveCards.js";
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
                          text: documentType,
                          weight: "bolder",
                          size: "medium",
                          color: "accent",
                        },
                        {
                          type: "TextBlock",
                          text: "Validation in Progress",
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
        content: {
          type: "AdaptiveCard",
          version: "1.0",
          body: [
            {
              type: "Container",
              style: "good",
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
                          text: "✅",
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
                          text: documentType,
                          weight: "bolder",
                          size: "medium",
                        },
                        {
                          type: "TextBlock",
                          text: "Validation Complete",
                          spacing: "none",
                          color: "good",
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
