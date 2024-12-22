import { handleTeamsCommands } from "./handleTeamsCommands.js";
import { handleSelectClientDirectory } from "./handleSelectClientDirectory.js";
import { handleProcessTestingActionSelected } from "./handleProcessTestingActionSelected.js";
import { handleProcessResponsesActionSelected } from "./handleProcessResponsesActionSelected.js";
import { handleResponseWorkbookSelected } from "./handleResponseWorkbookSelected.js";
import { handleTextMessages } from "./handleTextMessages.js";
import { handleTestingWorkbookSelected } from "./handleTestingWorkbookSelected.js";
import { handleProcessSelectJob } from "./handleProcessSelectJob.js";
import { handleCallProcessJobAgent } from "./handleCallProcessJob.js";
import { createWorkflow1ValidationCard } from "../utils/adaptiveCards.js";

/**
 * This function handles messages from the bot.
 * It checks the type of the activity and performs the appropriate action.
 * If the activity is a message, it strips out the bot mention and processes the user message.
 * If the activity is a command, it calls the appropriate handler function.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */

export async function handleMessages(context) {
  // Handle message activities
  if (context.activity.type === "message") {
    // Strip out the bot mention from the message
    const userMessage = context.activity.text
      ?.replace(/<at>.*?<\/at>/g, "")
      .trim();

    // Create action variable from context
    const action = context.activity.value?.action;
    const value = context.activity.value?.value;

    switch (action || userMessage) {
      // Begins the process inputs / evidence packs workflow
      // Displays client selection card with the client directories
      // Returns the action "processClientSelected" when the user selects a client
      case "p":
        await handleTeamsCommands(context, "process");
        break;

      // User has selected a client from the client selection card
      // Displays Job selection card with the Job directories - Inside the Evidence Pack folder
      // Returns the action "processJobSelected" when the user selects a client
      case "processClientSelected":
        await handleProcessSelectJob(context);
        break;

      // User has selected a client and a Job for the Processing Agent
      // TODO: Make a post request to the Processing Agent, pass jobID
      case "processJobSelected":
        await handleCallProcessJobAgent(context);
        break;

      // Begins the audit workflow
      // Displays client selection card with the client directories
      // Returns the action "auditClientSelected" when the user selects a client
      case "a":
        await handleTeamsCommands(context, "audit");
        break;

      // User has selected a client from the client selection card
      // Updates the card displaying the selected client
      // Displays the Audit Actions card with buttons to process the Testing worksheet or client responses worksheet
      case "auditClientSelected":
        await handleSelectClientDirectory(context);
        break;

      // User has selected the "Process Testing Worksheet" button from the Audit Actions card
      // Displays the file selection card for the Testing worksheet
      // Returns the action "testingWorkbookSelected" when the user selects a file
      case "processTestingActionSelected":
        await handleProcessTestingActionSelected(context);
        break;

      // User has selected a Testing worksheet from the file selection card
      // Processes the Testing worksheet, create RFI Response workbook
      case "testingWorkbookSelected":
        await handleTestingWorkbookSelected(context);
        break;

      // User has selected the "Process Client Responses" button from the Audit Actions card
      // Displays the file selection card for the client responses worksheet
      // Returns the action "processResponsesActionSelected" when the user selects a file
      case "processResponsesActionSelected":
        await handleProcessResponsesActionSelected(context);
        break;

      // User has selected a Responses Workbook from the file selection card
      // Processes the Responses workbook, generate and write auditor notes
      case "responsesWorkbookSelected":
        await handleResponseWorkbookSelected(context);
        break;

      // Handle the approve/reject actions
      case "approve_processing":
        console.log(`Processing approval action: ${value}`);
        console.log("Full context value:", context.activity.value);

        // Immediately update card to show processing state
        await context.updateActivity({
          type: "message",
          id: context.activity.replyToId,
          attachments: [
            {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: {
                type: "AdaptiveCard",
                version: "1.4",
                body: [
                  {
                    type: "TextBlock",
                    text:
                      value === "yes"
                        ? "✅ Processing approval..."
                        : "❌ Processing rejection...",
                    size: "Large",
                    weight: "Bolder",
                    wrap: true,
                  },
                  {
                    type: "TextBlock",
                    text: "Please wait while we process your response...",
                    wrap: true,
                  },
                ],
              },
            },
          ],
        });

        // TODO: make another POST request to /api/test-teams-message
        console.log("Processing approval action:", value);
        await new Promise((resolve) => setTimeout(resolve, 3000)); // Short timeout
        await context.updateActivity({
          type: "message",
          id: context.activity.replyToId,
          attachments: [
            {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: {
                type: "AdaptiveCard",
                version: "1.4",
                body: [
                  {
                    type: "TextBlock",
                    text: `Processing complete. Approved: ${
                      value === "yes" ? "✅" : "❌"
                    }`,
                    size: "Large",
                    weight: "Bolder",
                    wrap: true,
                  },
                ],
              },
            },
          ],
        });
        // with the final result after your processing is complete
        break;

      // Handle validation button clicks
      case "validate":
        const documentType = context.activity.value.documentType;
        const currentValidations = context.activity.value.currentValidations;
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
              content: createWorkflow1ValidationCard(updatedValidations),
            },
          ],
        });
        break;

      // Handle image approval/rejection
      case "approve":
      case "reject":
        const action = context.activity.value.action;
        const images = context.activity.value.images;

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
                                text: `Processing ${action}`,
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
        await new Promise((resolve) => setTimeout(resolve, 2000));

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
                                text: `Images ${
                                  action === "approve" ? "Approved" : "Rejected"
                                }`,
                                weight: "bolder",
                                size: "medium",
                              },
                              {
                                type: "TextBlock",
                                text: `${
                                  action === "approve"
                                    ? "Approval"
                                    : "Rejection"
                                } Complete`,
                                spacing: "none",
                                color:
                                  action === "approve" ? "good" : "attention",
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
        break;

      // Handle any other text messages from the user
      default:
        if (userMessage) {
          await handleTextMessages(context, userMessage);
        }
        break;
    }
  }
}
