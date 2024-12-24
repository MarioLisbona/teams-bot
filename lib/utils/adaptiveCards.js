import { chunk } from "./utils.js";
/**
 * This function creates an updated card to show the user that the RFI processing is complete.
 * @param {Object} selectedFileData - The data of the selected file.
 * @param {string} newWorkbookName - The name of the new workbook.
 * @returns {Object} - The updated card.
 */
export const createUpdatedCard = (selectedFileData, newWorkbookName) => {
  return {
    type: "AdaptiveCard",
    version: "1.0",
    style: "emphasis",
    body: [
      {
        type: "Container",
        style: "emphasis",
        items: [
          {
            type: "TextBlock",
            text: "RFI Processing Complete",
            size: "large",
            weight: "bolder",
            color: "accent",
          },
          {
            type: "TextBlock",
            textFormat: "markdown",
            text: `✅ Processed workbook:\n\n - **${selectedFileData.name}**`,
            wrap: true,
            spacing: "small",
          },
          {
            type: "TextBlock",
            textFormat: "markdown",
            text: `🛠️ Client RFI spreadsheet created:\n\n - **${newWorkbookName}**`,
            wrap: true,
          },
        ],
        padding: "medium",
      },
    ],
  };
};

/**
 * This function creates a file selection card to allow the user to select a file.
 * @param {Array} files - The list of files to be displayed in the card.
 * @param {string} directoryId - The ID of the directory containing the files.
 * @param {string} directoryName - The name of the directory containing the files.
 * @param {string} customSubheading - The subheading to be displayed in the card.
 * @param {string} buttonText - The text to be displayed on the button.
 * @param {string} action - The action to be performed when the button is clicked.
 * @returns {Object} - The file selection card.
 */
export const createFileSelectionCard = (
  files,
  directoryId,
  directoryName,
  customSubheading,
  buttonText,
  action
) => {
  return {
    type: "AdaptiveCard",
    version: "1.0",
    style: "emphasis",
    body: [
      {
        type: "Container",
        style: "emphasis",
        items: [
          {
            type: "TextBlock",
            text: customSubheading || "Process Testing Worksheet",
            size: "large",
            weight: "bolder",
            color: "accent",
          },
          {
            type: "TextBlock",
            textFormat: "markdown",
            text: `Client: **${directoryName}**`,
            wrap: true,
            spacing: "small",
          },
          {
            type: "Input.ChoiceSet",
            id: "fileChoice",
            style: "compact",
            isRequired: true,
            errorMessage: "Please select a file",
            placeholder: "Choose a file",
            choices: files.map((file) => ({
              title: file.name,
              value: JSON.stringify({
                name: file.name,
                id: file.id,
                directoryId: directoryId,
                directoryName: directoryName,
              }),
            })),
          },
        ],
        padding: "medium",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: buttonText || "Process Worksheet",
        data: {
          action: action || "testingWorkbookSelected",
          directoryId: directoryId,
          directoryName: directoryName,
          timestamp: Date.now(),
        },
        style: "positive",
      },
    ],
    $schema: "http://adaptivecards.io/schemas/adaptive-card",
    version: "1.2",
  };
};

/**
 * This function creates a client selection card to allow the user to select a client.
 * @param {Array} directories - The list of directories to be displayed in the card.
 * @param {string} actionText - The text to be displayed on the button.
 * @returns {Object} - The client selection card.
 */
export function createClientSelectionCard(directories, actionText) {
  let actionToSend;
  let selectNoun;
  let cardTitle;

  switch (actionText) {
    case "Process Evidence pack":
      cardTitle = actionText;
      actionToSend = "processClientSelected";
      selectNoun = "client";
      break;
    case "Choose a Job":
      cardTitle = "Process Evidence pack";
      actionToSend = "processJobSelected";
      selectNoun = "Job";
      break;
    case "Audit Workbook":
      cardTitle = actionText;
      actionToSend = "auditClientSelected";
      selectNoun = "client";
      break;
  }
  return {
    type: "AdaptiveCard",
    version: "1.0",
    style: "emphasis",
    body: [
      {
        type: "Container",
        style: "emphasis",
        items: [
          {
            type: "TextBlock",
            text: cardTitle,
            size: "large",
            weight: "bolder",
            color: "accent",
          },
          {
            type: "TextBlock",
            text: `Select a ${selectNoun}`,
            spacing: "small",
          },
          {
            type: "Input.ChoiceSet",
            id: "directoryChoice",
            style: "compact",
            isRequired: true,
            errorMessage: "Please select a directory",
            placeholder: "Choose a directory",
            choices: directories.map((dir) => ({
              title: dir.name,
              value: JSON.stringify({
                id: dir.id,
                name: dir.name,
              }),
            })),
          },
        ],
        padding: "medium",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: `Select ${selectNoun}`,
        data: {
          action: actionToSend,
        },
      },
    ],
  };
}

/**
 * This function creates a card to show the user the selected client.
 * @param {Object} selectedDirectory - The selected client directory.
 * @returns {Object} - The selected client card.
 */
export function createUpdatedClientDirectoryCard(selectedDirectory) {
  return {
    type: "AdaptiveCard",
    version: "1.0",
    style: "emphasis",
    body: [
      {
        type: "Container",
        style: "emphasis",
        items: [
          {
            type: "TextBlock",
            text: "Selected Client:",
            size: "large",
            weight: "bolder",
            color: "accent",
          },
          {
            type: "TextBlock",
            text: selectedDirectory.name,
            color: "good",
            spacing: "small",
          },
        ],
        padding: "medium",
      },
    ],
  };
}

/**
 * This function creates an audit actions card to show the user the available actions.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 * @param {string} selectedDirectoryName - The name of the selected client directory.
 * @returns {Object} - The audit actions card.
 */
export const createAuditActionsCard = (context, selectedDirectoryName) => {
  return {
    type: "AdaptiveCard",
    version: "1.0",
    style: "emphasis",
    body: [
      {
        type: "Container",
        style: "emphasis",
        items: [
          {
            type: "TextBlock",
            text: "Audit Processing Actions",
            size: "large",
            weight: "bolder",
            color: "accent",
          },
          {
            type: "TextBlock",
            textFormat: "markdown",
            text: `Selected Client: **${selectedDirectoryName}**`,
            spacing: "small",
          },
        ],
        padding: "medium",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Process Testing Worksheet",
        data: {
          action: "processTestingActionSelected",
          directoryChoice: context.activity.value.directoryChoice,
        },
      },
      {
        type: "Action.Submit",
        title: "Process Client Responses",
        data: {
          action: "processResponsesActionSelected",
          directoryChoice: context.activity.value.directoryChoice,
        },
      },
    ],
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.2",
  };
};

/**
 * This function creates an updated actions card to show the user the selected action and client.
 * @param {string} selectedDirectoryName - The name of the selected client directory.
 * @param {string} selectedAction - The action selected by the user.
 * @returns {Object} - The updated actions card.
 */
export const createUpdatedActionsCard = (
  selectedDirectoryName,
  selectedAction
) => {
  return {
    type: "AdaptiveCard",
    version: "1.0",
    style: "emphasis",
    body: [
      {
        type: "Container",
        style: "emphasis",
        items: [
          {
            type: "TextBlock",
            text: "Audit Processing Actions",
            size: "large",
            weight: "bolder",
            color: "accent",
          },
          {
            type: "TextBlock",
            textFormat: "markdown",
            text: `Selected Client: **${selectedDirectoryName}**`,
            spacing: "small",
          },
          {
            type: "TextBlock",
            text: `Selected Action: ${selectedAction}`,
            color: "good",
            spacing: "medium",
          },
        ],
        padding: "medium",
      },
    ],
  };
};

/**
 * This function creates a card to show the user the processing of RFI Client Responses.
 * @param {string} fileName - The name of the file being processed.
 * @returns {Object} - The processing responses card.
 */
export function createProcessingResponsesCard(fileName) {
  return {
    type: "AdaptiveCard",
    version: "1.0",
    style: "emphasis",
    body: [
      {
        type: "Container",
        style: "emphasis",
        items: [
          {
            type: "TextBlock",
            text: "Processing RFI Client Responses",
            size: "large",
            weight: "bolder",
            color: "accent",
          },
          {
            type: "TextBlock",
            text: `Selected file: ${fileName}`,
            color: "good",
            wrap: true,
            spacing: "small",
          },
        ],
        padding: "medium",
      },
    ],
  };
}

/**
 * This function creates a help card to show the user the available commands.
 * @returns {Object} - The help card.
 */
export const createHelpCard = () => {
  return {
    type: "AdaptiveCard",
    version: "1.0",
    style: "emphasis",
    body: [
      {
        type: "Container",
        style: "emphasis",
        items: [
          {
            type: "TextBlock",
            text: "👋 Hi!",
            size: "large",
            weight: "bolder",
            color: "accent",
          },
          {
            type: "TextBlock",
            text: "I'm your Audit Assistant. Below are a list of commands available:",
            wrap: true,
            spacing: "small",
          },
          {
            type: "FactSet",
            spacing: "medium",
            facts: [
              {
                title: "process",
                value: "Process an evidence pack for a job",
              },
              {
                title: "audit",
                value: "Begin the audit workflow for a client",
              },
              {
                title: "help",
                value: "Show this help message",
              },
            ],
          },
          {
            type: "TextBlock",
            text: "Prepend any commands with **@els-test-bot**",
            wrap: true,
            spacing: "medium",
            isSubtle: true,
            style: "default",
            weight: "default",
          },
        ],
        padding: "medium",
      },
    ],
  };
};

// export const createProcessingResultsCard = (message, images, chunk) => {
//   return {
//     type: "AdaptiveCard",
//     version: "1.4",
//     $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
//     fallbackText: "Your client doesn't support Adaptive Cards.",
//     speak: message,
//     width: "full",
//     body: [
//       {
//         type: "Container",
//         width: "stretch",
//         // minHeight: "100px",
//         items: [
//           {
//             type: "TextBlock",
//             text: message,
//             size: "Large",
//             weight: "Bolder",
//             wrap: true,
//           },
//         ],
//       },
//       // Split images into groups of 3 and create multiple ColumnSets
//       ...chunk(images, 3).map((imageGroup) => ({
//         type: "ColumnSet",
//         width: "stretch",
//         spacing: "Large",
//         columns: imageGroup.map((url) => ({
//           type: "Column",
//           width: "stretch",
//           items: [
//             {
//               type: "Image",
//               url: url,
//               size: "Large",
//               spacing: "None",
//               horizontalAlignment: "Center",
//             },
//           ],
//         })),
//       })),
//     ],
//     actions: [
//       {
//         type: "Action.Submit",
//         title: "✅ Approve",
//         data: {
//           action: "approve_processing",
//           value: "yes",
//         },
//         style: "positive",
//       },
//       {
//         type: "Action.Submit",
//         title: "❌ Reject",
//         data: {
//           action: "approve_processing",
//           value: "no",
//         },
//         style: "destructive",
//       },
//     ],
//   };
// };

/**
 * This function creates a card to display the workflow validation checklist.
 * @param {Object} validations - Object containing validation results for each document
 * @returns {Object} - The workflow validation card
 */
export const createWorkflow1ValidationCard = (validations, jobId) => {
  const createValidationItem = (key, label, exists, index) => {
    const isProcessing = validations[`${key}Processing`];

    return {
      type: "Container",
      style: index % 2 === 0 ? "default" : "emphasis",
      spacing: "none",
      items: [
        {
          type: "ColumnSet",
          spacing: "none",
          columns: [
            {
              type: "Column",
              width: "stretch",
              verticalContentAlignment: "center",
              items: [
                {
                  type: "TextBlock",
                  text: `${index + 1}. ${label}`,
                  wrap: true,
                  spacing: "none",
                },
              ],
            },
            {
              type: "Column",
              width: "auto",
              verticalContentAlignment: "center",
              horizontalAlignment: "right",
              items: [
                {
                  type: "TextBlock",
                  text: exists ? "✅" : "",
                  spacing: "none",
                  horizontalAlignment: "right",
                },
              ],
            },
            {
              type: "Column",
              width: "auto",
              verticalContentAlignment: "center",
              items: exists
                ? []
                : [
                    {
                      type: "TextBlock",
                      text: isProcessing ? "🔄 Processing 🔄" : "",
                      wrap: true,
                      spacing: "none",
                      weight: "bolder",
                    },
                    {
                      type: "ActionSet",
                      actions: isProcessing
                        ? []
                        : [
                            {
                              type: "Action.Submit",
                              title: "Validate",
                              style: "positive",
                              data: {
                                action: "validate",
                                documentType: key,
                                currentValidations: validations,
                                jobId: jobId,
                              },
                            },
                          ],
                      spacing: "none",
                    },
                  ],
            },
          ],
        },
      ],
      padding: "small",
    };
  };

  return {
    type: "AdaptiveCard",
    version: "1.0",
    style: "emphasis",
    body: [
      {
        type: "Container",
        style: "emphasis",
        items: [
          {
            type: "TextBlock",
            text: "Workflow 1 Validation Checklist",
            size: "large",
            weight: "bolder",
            color: "accent",
          },
          {
            type: "TextBlock",
            text: jobId,
            size: "medium",
            weight: "bolder",
            color: "accent",
          },
          {
            type: "Container",
            spacing: "medium", // Add spacing before the validation items
            items: [
              ...Object.entries({
                nomForm: ["Nom Form", validations.nomForm],
                siteAssessment: [
                  "Site Assessment Report",
                  validations.siteAssessment,
                ],
                taxInvoice: ["Tax Invoice", validations.taxInvoice],
                ccew: ["CCEW", validations.ccew],
                installerDec: [
                  "Installer Declaration",
                  validations.installerDec,
                ],
                coc: ["CoC", validations.coc],
                gtp: ["GTP", validations.gtp],
              }).map(([key, [label, exists]], index) =>
                createValidationItem(key, label, exists, index)
              ),
            ],
          },
        ],
        padding: "medium",
      },
    ],
  };
};

export function createValidationProgressCard(documentType) {
  return {
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
  };
}

export function createValidationCompletionCard(documentType) {
  return {
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
  };
}

export function createValidateSignaturesCard(message, images) {
  const reviewCard = {
    type: "AdaptiveCard",
    version: "1.0",
    body: [
      {
        type: "TextBlock",
        text: "Validate Signatures",
        size: "large",
        weight: "bolder",
        color: "accent",
      },
      {
        type: "TextBlock",
        text: message,
        size: "medium",
        weight: "default",
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
      {
        type: "Input.Text",
        id: "reviewComment",
        placeholder: "Enter your comments here...",
        isMultiline: false,
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Submit",
        data: {
          action: "submitSigReview",
          images: images,
          reviewComment: message,
        },
      },
    ],
  };

  return reviewCard;
}

export function createApproveRejectImagesCard(message, images) {
  const imagesCard = {
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
  };

  return imagesCard;
}

export const createTeamsUpdateCard = (text, emoji, style) => {
  return {
    type: "AdaptiveCard",
    version: "1.0",
    body: [
      {
        type: "Container",
        style: style,
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
                    text: emoji,
                    size: "large",
                    spacing: "none",
                  },
                ],
              },
              {
                type: "Column",
                width: "stretch",
                verticalContentAlignment: "center",
                items: [
                  {
                    type: "TextBlock",
                    text: text,
                    wrap: true,
                    spacing: "small",
                  },
                ],
              },
            ],
          },
        ],
        padding: "default",
      },
    ],
  };
};

export function createWorkflowProgressCard(jobId, workflowStep, isComplete) {
  return {
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
                    text: isComplete ? "✅" : "🔄",
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
                    text: `Job ID: ${jobId} - Workflow Step: ${workflowStep}`,
                    weight: "bolder",
                    size: "medium",
                    color: "accent",
                  },
                  {
                    type: "TextBlock",
                    text: isComplete
                      ? "Validation Complete"
                      : "Validation in Progress",
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
  };
}
