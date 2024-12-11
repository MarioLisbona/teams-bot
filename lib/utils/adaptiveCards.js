export const createUpdatedCard = (selectedFileData, newWorkbookName) => {
  return {
    type: "AdaptiveCard",
    body: [
      {
        type: "TextBlock",
        text: "RFI Processing Complete",
        weight: "bolder",
        size: "large",
        color: "good",
      },
      {
        type: "TextBlock",
        textFormat: "markdown",
        text: `✅ Processed workbook:\n\n - **${selectedFileData.name}**`,
        wrap: true,
      },
      {
        type: "TextBlock",
        textFormat: "markdown",
        text: `🛠️ Client RFI spreadsheet created:\n\n - **${newWorkbookName}**`,
        wrap: true,
      },
    ],
    $schema: "http://adaptivecards.io/schemas/adaptive-card",
    version: "1.2",
  };
};

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
    body: [
      {
        type: "TextBlock",
        text: customSubheading || "Process Testing Worksheet",
        weight: "bolder",
        size: "large",
      },
      {
        type: "TextBlock",
        textFormat: "markdown",
        text: `Client: **${directoryName}**`,
        wrap: true,
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
    actions: [
      {
        type: "Action.Submit",
        title: buttonText || "Process Worksheet",
        data: {
          action: action || "selectClientWorkbook",
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

export function createDirectorySelectionCard(directories) {
  const card = {
    type: "AdaptiveCard",
    version: "1.0",
    body: [
      {
        type: "TextBlock",
        text: "Audit Client Workbook",
        weight: "bolder",
        size: "large",
      },
      {
        type: "TextBlock",
        text: "Select a client to process:",
        weight: "lighter",
        size: "medium",
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
    actions: [
      {
        type: "Action.Submit",
        title: "Select Client",
        data: {
          action: "selectDirectory",
        },
      },
    ],
  };

  return card;
}

export function createUpdatedDirectoryCard(selectedDirectory) {
  return {
    type: "AdaptiveCard",
    version: "1.0",
    body: [
      {
        type: "TextBlock",
        text: "Selected Client:",
        weight: "bolder",
      },
      {
        type: "TextBlock",
        text: selectedDirectory.name,
        color: "good",
      },
    ],
    actions: [], // No actions since it's disabled
  };
}

export const createActionsCard = (context, selectedDirectoryName) => {
  const actionsCard = {
    type: "AdaptiveCard",
    body: [
      {
        type: "TextBlock",
        text: "Audit Processing Actions",
        weight: "Bolder",
        size: "Large",
      },
      {
        type: "TextBlock",
        textFormat: "markdown",
        text: `Selected Client: **${selectedDirectoryName}**`,
        size: "Medium",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Process Testing Worksheet",
        data: {
          action: "processTestingWorksheet",
          directoryChoice: context.activity.value.directoryChoice,
        },
      },
      {
        type: "Action.Submit",
        title: "Process Client Responses",
        data: {
          action: "processClientResponses",
          directoryChoice: context.activity.value.directoryChoice,
        },
      },
    ],
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.2",
  };

  return actionsCard;
};

export const createUpdatedActionsCard = (
  selectedDirectoryName,
  selectedAction
) => {
  return {
    type: "AdaptiveCard",
    body: [
      {
        type: "TextBlock",
        text: "Audit Processing Actions",
        weight: "Bolder",
        size: "Large",
      },
      {
        type: "TextBlock",
        textFormat: "markdown",
        text: `Selected Client: **${selectedDirectoryName}**`,
        size: "Medium",
      },
      {
        type: "TextBlock",
        text: `Selected Action: ${selectedAction}`,
        color: "good",
        spacing: "Medium",
      },
    ],
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.2",
  };
};

export function createProcessingResponsesCard(fileName) {
  return {
    type: "AdaptiveCard",
    body: [
      {
        type: "TextBlock",
        text: "Processing RFI Client Responses",
        weight: "bolder",
        size: "large",
      },
      {
        type: "TextBlock",
        color: "good",
        text: `Selected file: ${fileName}`,
        wrap: true,
      },
    ],
    version: "1.2",
  };
}
