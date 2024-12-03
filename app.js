import express from "express";
import { BotFrameworkAdapter, TurnContext } from "botbuilder";
import dotenv from "dotenv";
import { getFileNamesAndIds } from "./lib/oneDrive.js";
import { createBotAdapter } from "./lib/createBotAdapter.js";
import { getGraphClient } from "./lib/msAuth.js";
import { processTesting } from "./lib/worksheetProcessing.js";
import { processTestingWorksheet } from "./lib/botProcessing.js";
dotenv.config();

const app = express();
app.use(express.json());
const port = process.env.PORT || 3978;

const adapter = await createBotAdapter();

// Simple bot logic
async function handleMessage(context) {
  const userMessage = context.activity.text?.trim();

  try {
    if (
      context.activity.type === "message" &&
      context.activity.value?.action === "selectClientWorkbook"
    ) {
      const selectedFileData = JSON.parse(context.activity.value.fileChoice);

      // Immediately respond to the card interaction
      if (
        context.activity.type === "invoke" ||
        context.activity.name === "adaptiveCard/action"
      ) {
        await context.sendActivity({
          type: "invokeResponse",
          value: {
            statusCode: 200,
            type: "application/vnd.microsoft.activity.message",
            value: "Processing...",
          },
        });
      }

      // Then delegate to the processing function
      const newWorkbookName = await processTestingWorksheet(
        context,
        adapter,
        selectedFileData
      );

      // Update the card to show it's been processed
      const updatedCard = {
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
    } else if (userMessage === "/pt") {
      const files = await getFileNamesAndIds(process.env.ONEDRIVE_ID);

      if (!files || files.length === 0) {
        await context.sendActivity("No workbooks found in OneDrive.");
        return;
      }

      const card = {
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

      await context.sendActivity({
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: card,
          },
        ],
      });
    } else if (context.activity.value?.action === "createRFI") {
      const clientWorkbookId = context.activity.value.clientWorkbookId;
      if (!clientWorkbookId) {
        await context.sendActivity("Error: Missing workbook ID");
        return;
      }

      await context.sendActivity(
        `Starting RFI spreadsheet creation for ${clientWorkbookId}...`
      );

      // Store the conversation reference for later use
      const conversationReference = TurnContext.getConversationReference(
        context.activity
      );
    } else if (userMessage) {
      // Handle other text messages
      if (userMessage.toLowerCase() === "help") {
        await context.sendActivity(
          "Available commands:\n" +
            "• /pt - Process the Testing Worksheet from a client workbook\n" +
            "• help - Show this help message"
        );
      } else {
        await context.sendActivity(`Echo: ${userMessage}`);
      }
    }
  } catch (error) {
    console.error("Handler Error:", error);
    await context.sendActivity(
      "❌ An error occurred while processing your request. Please try again or contact support."
    );
  }
}

app.get("/", (req, res) => {
  res.send("Server is running");
});

let showThumbnailCard = false; // Initial state of the boolean

// POST route to change the boolean value
app.post("/api/showCard", async (req, res) => {
  const { clientName } = req.body; // Added clientName parameter
  showThumbnailCard = true; // Set the boolean to true

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

  // Send the thumbnail card to a specific user or channel
  const message = {
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.thumbnail",
        content: thumbnailCard,
      },
    ],
  };

  // Replace 'YOUR_USER_ID' with the actual user ID or context as needed
  const userId =
    "29:1xG3Q1I-CSlqfIN-rd3oJTcketwzGgjE75Hppzj3B852n2t16FgmTSK-aWI7tgt0oAhpIB101UU_5wU-njL2Lzg"; // Replace with the actual user ID
  const conversationReference = {
    channelId: "msteams", // Replace with the actual channel ID
    serviceUrl:
      "https://smba.trafficmanager.net/au/50a4078b-a9b7-4a68-8223-231f9a012eb3/", // Replace with the actual service URL
    conversation: {
      id: "a:17QE4g2Rlk_JgpWFzQeXPS7BPXlao8YcHetxp1g5BNUU7DI2_7tYFI2JdPFhgReAlDM9eBzFy0fB-8p1M2D03TwWNbLJRtA_z9kalAVVlqrl4bxZubxuSTAqyLSAslqQB",
    }, // Replace with the actual conversation ID
    recipient: { id: userId },
    from: { id: process.env.MICROSOFT_APP_ID }, // Replace with your bot's ID
  };

  // Create a context for sending the message
  const context = await adapter.createContext(conversationReference);
  await context.sendActivity(message);

  res.status(200).send("Thumbnail card sent successfully.");
});

// Webhook endpoint for bot messages
app.post("/api/messages", (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    if (context.activity.type === "message") {
      await handleMessage(context);
    } else {
      await context.sendActivity(`[${context.activity.type}] event detected`);
    }
  });
});

// Start the server
app.listen(port, () => {
  console.log(
    `\nBot is running on http://localhost:${port}/api/messages\nServer is running on http://localhost:${port}/`
  );
});
