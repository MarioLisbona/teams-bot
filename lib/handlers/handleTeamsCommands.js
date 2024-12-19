import { getClientDirectories } from "../utils/fileStorageAndRetrieval.js";
import { createClientSelectionCard } from "../utils/adaptiveCards.js";

export async function handleTeamsCommands(context, command) {
  let commandAction;
  switch (command) {
    case "process":
      commandAction = "Process Evidence pack";
      break;
    case "audit":
      commandAction = "Audit Workbook";
      break;
  }
  // Retrieve the client directories from SharePoint
  const rootDirectoryName = process.env.ROOT_DIRECTORY_NAME;
  const clientDirectories = await getClientDirectories(rootDirectoryName);

  if (!clientDirectories || clientDirectories.length === 0) {
    await context.sendActivity(
      "No Client Directories found in SharePoint directory."
    );
    return;
  }

  // Create the client selection card and return the action "processClientSelected"
  const clientDirectorySelectionCard = createClientSelectionCard(
    clientDirectories,
    commandAction
  );

  // Send the client selection card to the user
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: clientDirectorySelectionCard,
      },
    ],
  });
}
