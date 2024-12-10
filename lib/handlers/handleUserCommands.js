import { getDirectories } from "../oneDrive.js";
import { createDirectorySelectionCard } from "../adaptiveCards.js";

export async function handleAuditCommand(context) {
  const rootDirectoryName = process.env.ROOT_DIRECTORY_NAME;
  const directories = await getDirectories(rootDirectoryName);

  if (!directories || directories.length === 0) {
    await context.sendActivity("No directories found in OneDrive.");
    return;
  }

  const directorySelectionCard = createDirectorySelectionCard(directories);
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: directorySelectionCard,
      },
    ],
  });
}
export async function handleTestCommand(context) {
  console.log("Testing the testing command");
  await context.sendActivity("ECHO-> Testing the testing command");
}
