import express from "express";
import { BotFrameworkAdapter } from "botbuilder";
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
  // console.log("Received activity:", context.activity);

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
  } else if (userMessage) {
    console.log(`You said: ${userMessage}`);
    await context.sendActivity(`You said: ${userMessage}`);
  }
}

app.get("/", (req, res) => {
  res.send("Server is running");
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
