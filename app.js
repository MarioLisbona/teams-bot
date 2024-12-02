import express from "express";
import { BotFrameworkAdapter, TurnContext } from "botbuilder";
import dotenv from "dotenv";

dotenv.config();

const app = express();
app.use(express.json());
const port = process.env.PORT || 3978;

// Create the BotFrameworkAdapter
const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
});

// Error handling
adapter.onTurnError = async (context, error) => {
  console.error(`[onTurnError]: ${error}`);
  await context.sendActivity("Oops, something went wrong!");
};

// Simple bot logic
async function handleMessage(context) {
  console.log("Received activity:", context.activity);

  const userMessage = context.activity.text?.trim();

  if (userMessage === "/process") {
    await context.sendActivity(
      "You triggered the /process command! Testing worksheet being processed..."
    );
    console.log(
      "Process command triggered - Testing worksheet being processed..."
    );
  } else if (
    context.activity.type === "message" &&
    context.activity.value?.action === "selectFile"
  ) {
    const selectedFile = context.activity.value.fileChoice;
    await context.sendActivity(`You selected: ${selectedFile}`);
    console.log(`File selected: ${selectedFile}`);
  } else if (userMessage === "/files") {
    const card = {
      type: "AdaptiveCard",
      body: [
        {
          type: "TextBlock",
          text: "Please select a file:",
        },
        {
          type: "Input.ChoiceSet",
          id: "fileChoice",
          style: "compact",
          choices: [
            { title: "File 1", value: "file1.txt" },
            { title: "File 2", value: "file2.txt" },
            { title: "File 3", value: "file3.txt" },
          ],
        },
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "Submit",
          data: { action: "selectFile" },
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
    const clientName = context.activity.value.clientName;
    processRFISpreadsheet(clientName);
    await context.sendActivity(
      `Starting RFI spreadsheet creation for ${clientName}...`
    );

    // Store the conversation reference for later use
    const conversationReference = TurnContext.getConversationReference(
      context.activity
    );

    // Add 5 second timeout and completion message
    setTimeout(async () => {
      // Create a new context for the delayed message
      const newContext = await adapter.createContext(conversationReference);
      await newContext.sendActivity(
        `RFI spreadsheet creation completed for ${clientName}!`
      );
    }, 5000);
  } else if (userMessage) {
    console.log(`You said: ${userMessage}`);
    await context.sendActivity(`You said: ${userMessage}`);
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

// Add this function to handle the background processing
async function processRFISpreadsheet(clientName) {
  console.log(`Processing RFI spreadsheet for ${clientName}`);

  return new Promise((resolve) => {
    setTimeout(() => {
      // Simulate processing completion
      console.log(`Completed processing for ${clientName}`);
      resolve(true);
    }, 5000);
  });
}
