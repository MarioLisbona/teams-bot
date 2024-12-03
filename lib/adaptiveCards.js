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

export const createFileSelectionCard = (files) => {
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
        text: "Please select the client workbook you would like to process:",
        wrap: true,
      },
      {
        type: "Input.ChoiceSet",
        id: "fileChoice",
        style: "compact",
        isRequired: true,
        choices: files.map((file) => ({
          title: file.name,
          value: JSON.stringify({ name: file.name, id: file.id }),
        })),
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Process Worksheet",
        data: {
          action: "selectClientWorkbook",
          timestamp: Date.now(), // Add timestamp to prevent reuse
        },
        style: "positive",
      },
    ],
    $schema: "http://adaptivecards.io/schemas/adaptive-card",
    version: "1.2",
  };
};
