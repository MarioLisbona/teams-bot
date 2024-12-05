import express from "express";
import { BotFrameworkAdapter, TurnContext } from "botbuilder";

import { getFileNamesAndIds } from "./lib/oneDrive.js";
import { createBotAdapter } from "./lib/createBotAdapter.js";
import { getGraphClient } from "./lib/msAuth.js";
import { processTesting } from "./lib/worksheetProcessing.js";
import { processTestingWorksheet } from "./lib/botProcessing.js";
import {
  createUpdatedCard,
  createFileSelectionCard,
} from "./lib/adaptiveCards.js";
import { handleMessage } from "./lib/handlers/handleMessage.js";
import { loadEnvironmentVariables } from "./lib/environment/setupEnvironment.js";

loadEnvironmentVariables();

const app = express();
app.use(express.json());
const port = process.env.PORT || 3978;

const adapter = await createBotAdapter();

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
      await handleMessage(adapter, context);
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
