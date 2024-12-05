export const createUpdatedCard = (selectedFileData, newWorkbookName) => {
  return {
    type: "AdaptiveCard",
    body: [
      {
        type: "TextBlock",
        text: "RFI Processing Complete",
        weight: "bolder",
      },
      {
        type: "TextBlock",
        textFormat: "markdown",
        text: `✅ Processed workbook: **${selectedFileData.name}**`,
        wrap: true,
      },
      {
        type: "TextBlock",
        textFormat: "markdown",
        text: `🛠️ Client RFI spreadsheet created:\n\n**${newWorkbookName}**`,
        wrap: true,
      },
    ],
    $schema: "http://adaptivecards.io/schemas/adaptive-card",
    version: "1.2",
  };
};

export const createFileSelectionCard = (files, directoryId, directoryName) => {
  return {
    type: "AdaptiveCard",
    body: [
      {
        type: "TextBlock",
        text: "Process Testing Worksheet",
        weight: "bolder",
        size: "medium",
      },
      {
        type: "TextBlock",
        text: `Client: ${directoryName}`,
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
        title: "Process Worksheet",
        data: {
          action: "selectClientWorkbook",
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
        weight: "bolder",
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

export const createUpdatedDirectoryCard = (selectedDirectoryName) => {
  // Create a disabled version of the directory selection card
  const updatedDirectoryCard = {
    type: "AdaptiveCard",
    version: "1.0",
    body: [
      {
        type: "TextBlock",
        text: "Directory selected",
        weight: "bolder",
        size: "large",
      },
      {
        type: "TextBlock",
        text: `Loading files from ${selectedDirectoryName}...`,
        weight: "bolder",
        size: "medium",
        wrap: true,
      },
    ],
  };

  return updatedDirectoryCard;
};

export const createThumbnailCard = (clientName) => {
  const thumbnailCard = {
    type: "ThumbnailCard",
    title: `Testing Worksheet Completed for ${clientName}`,
    text: `The Testing Worksheet for ${clientName} has been completed. it is ready to be processed.`,
    images: [
      {
        url: "https://example.com/thumbnail.png",
      },
    ],
    buttons: [
      {
        type: "messageBack",
        title: "Create RFI Spreadsheet",
        text: "Processing RFI Spreadsheet...",
        displayText: "Creating RFI Spreadsheet...",
        value: {
          action: "createRFI",
          clientName: clientName,
        },
      },
    ],
  };

  return thumbnailCard;
};

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
        text: `Selected Client: ${selectedDirectoryName}`,
        weight: "Bolder",
        size: "Medium",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Process Testing Worksheet",
        data: {
          action: "processTestingWorksheet",
          directoryChoice: context.activity.value.directoryChoice, // Pass through the directory info
        },
      },
      {
        type: "Action.Submit",
        title: "Email RFI Spreadsheet",
        data: {
          action: "emailRFI",
          directoryChoice: context.activity.value.directoryChoice, // Pass through the directory info
        },
      },
    ],
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.2",
  };

  return actionsCard;
};
