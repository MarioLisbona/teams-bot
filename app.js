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
  const userMessage = context.activity.text.trim();

  if (userMessage === "/daily") {
    await context.sendActivity("You triggered the /daily command!");
    console.log("Daily command triggered");
  } else {
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
