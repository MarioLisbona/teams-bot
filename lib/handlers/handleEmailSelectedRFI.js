import { createUpdatedRFIEmailCard } from "../adaptiveCards.js";

export async function handleEmailSelectedRFI(context) {
  const fileData = JSON.parse(context.activity.value.fileChoice);
  const directoryName = context.activity.value.directoryName;

  // Update the card to show it's being emailed
  const updatedCard = createUpdatedRFIEmailCard(fileData, directoryName);

  await context.updateActivity({
    type: "message",
    id: context.activity.replyToId,
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: updatedCard,
      },
    ],
  });

  // Log detailed file information
  const emailLogMessage = `Emailing RFI file: ${fileData.name} \nID: ${fileData.id}`;
  console.log(emailLogMessage);
  await context.sendActivity(emailLogMessage);
}
