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
  console.log("Debug - Creating card with directory info:", {
    directoryId,
    directoryName,
  });

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
        text: `Directory: ${directoryName}`,
        wrap: true,
      },
      {
        type: "Input.ChoiceSet",
        id: "fileChoice",
        style: "compact",
        isRequired: true,
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
        text: "Select a directory:",
        weight: "bolder",
        size: "medium",
      },
      {
        type: "Input.ChoiceSet",
        id: "directoryChoice",
        style: "expanded",
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
        title: "Select Directory",
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
        size: "medium",
      },
      {
        type: "TextBlock",
        text: `Loading files from ${selectedDirectoryName}...`,
        wrap: true,
      },
    ],
  };

  return updatedDirectoryCard;
};
