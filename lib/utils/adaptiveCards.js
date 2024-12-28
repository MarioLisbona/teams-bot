import { chunk } from "./utils.js";

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
          action: action || "testingRfiWorkbookSelected",
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
    case "Process RFI":
      cardTitle = actionText;
      actionToSend = "rfiClientSelected";
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
 * This function creates an RFI Actions card to show the user the available actions.
 * Two buttons are displayed for the user to select the action they want to perform.
 * The action "processRfiActionSelected" is returned when the user selects the "Process RFI Worksheet" button.
 * The action "processResponsesActionSelected" is returned when the user selects the "Process Client Responses" button.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 * @param {string} selectedDirectoryName - The name of the selected client directory.
 * @returns {Object} - The RFI Actions card.
 */
export const createRfiActionsCard = (context, selectedDirectoryName) => {
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
            text: "RFI Actions",
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
        title: "Process RFI Worksheet",
        data: {
          action: "processRfiActionSelected",
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
            type: "TextBlock",
            text: "Prepend any commands or Agent prompts with **@els-test-bot**",
            wrap: true,
            spacing: "medium",
            isSubtle: true,
            style: "default",
            weight: "default",
          },
          {
            type: "TextBlock",
            text: "Bot commands",
            size: "medium",
            weight: "bolder",
            color: "accent",
            spacing: "medium",
          },
          {
            type: "FactSet",
            spacing: "medium",
            facts: [
              {
                title: "rfi",
                value: "Process RFI's in Testing Worksheet",
              },
              {
                title: "help",
                value: "Show this help message",
              },
            ],
          },
          {
            type: "TextBlock",
            text: "Workflow Agent",
            size: "medium",
            weight: "bolder",
            color: "accent",
            spacing: "medium",
          },
          {
            type: "TextBlock",
            text: "To trigger the Workflow Agent, use a command similar to the one below:",
            wrap: true,
            spacing: "small",
          },
          {
            type: "TextBlock",
            text: "`@els-test-bot Start the audit for <jobId>`",
            wrap: true,
            spacing: "small",
            fontType: "monospace",
          },
          {
            type: "TextBlock",
            text: "Example: @els-test-bot Start the audit for jobId 123456",
            wrap: true,
            spacing: "small",
          },
        ],
        padding: "medium",
      },
    ],
  };
};

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
          action: "validateSignatures",
          images: images,
          reviewComment: message,
        },
      },
    ],
  };

  return reviewCard;
}

export const createTeamsUpdateCard = (text, userMessage, emoji, style) => {
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
                    size: "extraLarge",
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
                  {
                    type: "TextBlock",
                    text: userMessage,
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
};

export function createWorkflowProgressNotificationCard(
  jobId,
  workflowStep,
  isComplete
) {
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

export const createHumanValidationStepsCard = (
  validationsRequired,
  completedValidations = {},
  jobId
) => {
  const validationLabels = {
    3.4: "Validation Step 3.4",
    3.5: "Validation Step 3.5",
    3.7: "Validation Step 3.7",
    7.6: "Validation Step 7.6",
    7.7: "Validation Step 7.7",
    7.9: "Validation Step 7.9",
    8.6: "Validation Step 8.6",
    8.7: "Validation Step 8.7",
  };

  const createValidationItem = (key, isRequired, index) => {
    const isCompleted = completedValidations[key] === true;

    return {
      type: "Container",
      style: index % 2 === 0 ? "default" : "emphasis",
      spacing: "none",
      items: [
        {
          type: "ColumnSet",
          columns: [
            {
              type: "Column",
              width: "stretch",
              items: [
                {
                  type: "TextBlock",
                  text: `${index + 1}. ${
                    validationLabels[key] || `Step ${key}`
                  }`,
                  wrap: true,
                },
              ],
            },
            {
              type: "Column",
              width: "auto",
              items: [
                {
                  type: "TextBlock",
                  text: isCompleted ? "✅" : "",
                  horizontalAlignment: "right",
                },
              ],
            },
            {
              type: "Column",
              width: "auto",
              items:
                isCompleted || !isRequired
                  ? []
                  : [
                      {
                        type: "ActionSet",
                        actions: [
                          {
                            type: "Action.Submit",
                            title: "Validate",
                            style: "positive",
                            data: {
                              action: "humanValidation",
                              validationType: key,
                              jobId: jobId,
                              currentValidations: validationsRequired,
                              completedValidations: completedValidations,
                            },
                          },
                        ],
                      },
                    ],
            },
          ],
        },
      ],
      padding: "small",
    };
  };

  const validationItems = Object.entries(validationsRequired).map(
    ([key, isRequired], index) => createValidationItem(key, isRequired, index)
  );

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
            text: "Workflow Validation Checklist",
            size: "large",
            weight: "bolder",
            color: "accent",
          },
          {
            type: "TextBlock",
            text: `Job ID: ${jobId}`,
            size: "medium",
            weight: "bolder",
            color: "accent",
          },
          {
            type: "Container",
            spacing: "medium",
            items: validationItems,
          },
        ],
        padding: "medium",
      },
    ],
  };
};
