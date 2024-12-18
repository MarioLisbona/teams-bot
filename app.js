import express from "express";
import { createBotAdapter } from "./lib/utils/createBotAdapter.js";
import { handleMessages } from "./lib/handlers/handleMessages.js";
import { loadEnvironmentVariables } from "./lib/environment/setupEnvironment.js";
import { handleTeamsActivity } from "./lib/utils/errorHandling.js";
// Load environment variables
loadEnvironmentVariables();

// Create the express app, JSON middleware and port
const app = express();
app.use(express.json());
const port = process.env.PORT || 3978;

// Create the bot adapter
const adapter = await createBotAdapter();

// Add error handling to adapter
adapter.onTurnError = async (context, error) => {
  console.error(`\n [onTurnError] unhandled error: ${error}`);
  console.error("Error details:", {
    activityType: context.activity.type,
    error: error.stack,
    activity: context.activity,
  });

  // Don't attempt to send messages if the conversation is gone
  if (
    error.message?.includes("Conversation not found") ||
    error.message?.includes("conversation not found") ||
    error.message?.includes("bot is not part of the conversation roster")
  ) {
    console.log("Conversation no longer exists, skipping error message");
    return;
  }

  // For messageDelete and messageUpdate, we don't want to send error messages
  if (
    context.activity.type === "messageDelete" ||
    context.activity.type === "messageUpdate"
  ) {
    console.log(`Skipping error message for ${context.activity.type} activity`);
    return;
  }

  try {
    await context.sendActivity("The bot encountered an error or bug.");
  } catch (err) {
    console.error("Error sending error message:", err);
  }
};

// Home route
app.get("/", (req, res) => {
  res.send("Server is running");
});

// Webhook endpoint for bot messages
app.post("/api/messages", (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    if (context.activity.type === "message") {
      await handleMessages(adapter, context);
    } else {
      await handleTeamsActivity(context);
    }
  });
});

// Start the server
app.listen(port, () => {
  console.log(
    `\nBot is running on http://localhost:${port}/api/messages\nServer is running on http://localhost:${port}/`
  );
});
