import { chunk } from "./utils.js";

/**
 * Creates an adaptive card for file selection from a directory.
 *
 * @description
 * Generates a card with:
 * - Custom subheading (defaults to "Process Testing Worksheet")
 * - Client directory name in markdown format
 * - Dropdown list of files with required selection
 * - Submit button with customizable text and action
 *
 * @param {Array<Object>} files - Array of files to display
 * @param {string} files[].name - Display name of the file
 * @param {string} files[].id - Unique identifier for the file
 * @param {string} directoryId - Unique identifier for the parent directory
 * @param {string} directoryName - Display name of the parent directory
 * @param {string} [customSubheading] - Optional custom subheading text
 * @param {string} [buttonText="Process Worksheet"] - Text to display on submit button
 * @param {string} [action="rfiWorksheetSelected"] - Action identifier for form submission
 *
 * @returns {Object} Adaptive card JSON structure with:
 * - Card type and version information
 * - Body containing file selection interface
 * - Submit action with directory context and timestamp
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
          action: action || "rfiWorksheetSelected",
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
 * Creates an adaptive card for client directory selection.
 *
 * @description
 * Generates a card with:
 * - Dynamic title based on action type
 * - Client selection dropdown with required selection
 * - Submit button with contextual text
 *
 * Supports different action types:
 * - "Process RFI": Creates RFI processing selection card
 *
 * @param {Array<Object>} directories - Array of client directories
 * @param {string} directories[].id - Unique identifier for directory
 * @param {string} directories[].name - Display name of directory
 * @param {string} actionText - Action type identifier (e.g., "Process RFI")
 *
 * @returns {Object} Adaptive card JSON structure with:
 * - Card type and version information
 * - Body containing client selection interface
 * - Submit action with appropriate action identifier
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
 * Creates an adaptive card displaying the selected client directory.
 *
 * @description
 * Generates a confirmation card with:
 * - "Selected Client:" header in accent color
 * - Client directory name in success color
 * - Emphasis styling for visual prominence
 *
 * @param {Object} selectedDirectory - Selected client directory information
 * @param {string} selectedDirectory.name - Display name of the selected directory
 *
 * @returns {Object} Adaptive card JSON structure with:
 * - Card type and version information
 * - Container with styled text blocks
 * - Consistent padding and spacing
 *
 * @example
 * createUpdatedClientDirectoryCard({ name: "Client A" })
 * // Returns card showing "Selected Client: Client A"
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
 * Creates an adaptive card displaying available RFI (Request for Information) actions.
 *
 * @description
 * Generates a card with:
 * - "RFI Actions" header in accent color
 * - Selected client name in markdown format
 * - Two action buttons:
 *   1. "Process RFI Worksheet" (triggers "processRfiActionSelected")
 *   2. "Process Client Responses" (triggers "processResponsesActionSelected")
 *
 * @param {Object} context - Teams bot turn context
 * @param {Object} context.activity - The incoming activity from Teams
 * @param {Object} context.activity.value - Values from previous card submission
 * @param {string} context.activity.value.directoryChoice - Selected directory information
 * @param {string} selectedDirectoryName - Display name of selected client directory
 *
 * @returns {Object} Adaptive card JSON structure with:
 * - Card type and schema information
 * - Body containing client context
 * - Action buttons for RFI processing options
 *
 * @example
 * createRfiActionsCard(context, "Client A")
 * // Returns card with RFI action buttons for Client A
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
 * Creates an adaptive card showing selected audit processing action and client.
 *
 * @description
 * Generates a status card with:
 * - "Audit Processing Actions" header in accent color
 * - Selected client name in markdown format
 * - Selected action in success color
 * - Emphasis styling for visual prominence
 *
 * @param {string} selectedDirectoryName - Display name of selected client directory
 * @param {string} selectedAction - Name of selected processing action
 *
 * @returns {Object} Adaptive card JSON structure with:
 * - Card type and version information
 * - Container with styled text blocks
 * - Consistent padding and spacing
 *
 * @example
 * createUpdatedActionsCard("Client A", "Process RFI")
 * // Returns card showing selected client and action status
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
 * Creates an adaptive card displaying bot commands and usage instructions.
 *
 * @description
 * Generates a help card with:
 * - Welcome message with bot introduction
 * - Bot mention instructions
 * - Available commands section:
 *   - "rfi": RFI processing instructions
 *   - "help": Help message display
 * - Workflow Agent section:
 *   - Usage instructions
 *   - Command format example
 *   - Practical usage example
 *
 * Card sections are organized with:
 * - Consistent spacing and padding
 * - Accent colors for section headers
 * - Monospace formatting for command examples
 * - Subtle styling for important notes
 *
 * @returns {Object} Adaptive card JSON structure with:
 * - Card type and version information
 * - Container with organized help content
 * - Styled text blocks and fact sets
 *
 * @example
 * createHelpCard()
 * // Returns comprehensive help card with commands and usage instructions
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
                value: "• Process RFI's in Testing Worksheet",
              },
              {
                title: "",
                value: "• Generate auditor notes for client responses to RFI's",
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

/**
 * Creates an adaptive card for signature validation review.
 *
 * @description
 * Generates a review card with:
 * - "Validate Signatures" header in accent color
 * - Custom message describing the validation task
 * - Grid display of signature images (3 columns)
 * - Comment input field for reviewer feedback
 * - Submit button for validation completion
 *
 * @param {string} message - Description or context for signature validation
 * @param {Array<string>} images - Array of image URLs to display for validation
 *
 * @returns {Object} Adaptive card JSON structure with:
 * - Card type and version information
 * - Image grid using ColumnSet layout
 * - Input field for comments
 * - Submit action with validation data
 *
 * @example
 * createValidateSignaturesCard("Please review these signatures", ["url1", "url2"])
 * // Returns card with signature images and review interface
 */
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

/**
 * Creates an adaptive card for Teams status updates.
 *
 * @description
 * Generates a status update card with:
 * - Two-column layout with emoji and text
 * - Main status message
 * - Optional subtle user message
 * - Customizable styling based on status type
 *
 * @param {string} text - Primary status message
 * @param {string} userMessage - Additional context or details
 * @param {string} emoji - Status indicator emoji (e.g., "✅", "❌", "🔄")
 * @param {string} style - Card style ("default", "emphasis", "good", "attention")
 *
 * @returns {Object} Adaptive card JSON structure with:
 * - Card type and version information
 * - ColumnSet with emoji and message layout
 * - Consistent padding and spacing
 *
 * @example
 * createTeamsUpdateCard("Processing complete", "Job ID: 123", "✅", "good")
 * // Returns status card with success styling
 */
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

/**
 * Creates an adaptive card showing workflow progress status.
 *
 * @description
 * Generates a status card with:
 * - Two-column layout with status emoji and details
 * - Job ID and workflow step information
 * - Visual status indicator (✅ for complete, 🔄 for in progress)
 * - Subtle status message based on completion state
 *
 * @param {string} jobId - Unique identifier for the workflow job
 * @param {string} workflowStep - Current step in the workflow process
 * @param {boolean} isComplete - Whether the workflow step is completed
 *
 * @returns {Object} Adaptive card JSON structure with:
 * - Card type and version information
 * - ColumnSet with status layout
 * - Consistent padding and emphasis styling
 *
 * @example
 * createWorkflowProgressNotificationCard("123", "Data Validation", false)
 * // Returns progress card showing "Data Validation" step in progress
 */
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

/**
 * Creates an adaptive card displaying human validation steps checklist.
 *
 * @description
 * Generates a validation checklist card with:
 * - Workflow validation header with job ID
 * - List of validation steps with:
 *   - Numbered steps with descriptive labels
 *   - Completion status indicators (✅)
 *   - "Validate" buttons for incomplete required steps
 *   - Alternating row styling for readability
 *
 * @param {Object} validationsRequired - Map of validation step IDs to requirement status
 * @param {Object} [completedValidations={}] - Map of validation step IDs to completion status
 * @param {string} jobId - Unique identifier for the workflow job
 *
 * @returns {Object} Adaptive card JSON structure with:
 * - Card type and version information
 * - Container with validation steps list
 * - Dynamic action buttons based on completion status
 *
 * Supported validation steps:
 * - 3.4, 3.5, 3.7: Initial validation phase
 * - 7.6, 7.7, 7.9: Secondary validation phase
 * - 8.6, 8.7: Final validation phase
 *
 */
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
                              validationsRequired: validationsRequired,
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
            // Map over the validation items array and add them to the card
            items: validationItems,
          },
        ],
        padding: "medium",
      },
    ],
  };
};
