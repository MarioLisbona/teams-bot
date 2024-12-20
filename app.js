import express from "express";
import { createBotAdapter } from "./lib/utils/createBotAdapter.js";
import { handleMessages } from "./lib/handlers/handleMessages.js";
import { loadEnvironmentVariables } from "./lib/environment/setupEnvironment.js";
import { handleTeamsActivity } from "./lib/utils/teamsActivity.js";
// Load environment variables
loadEnvironmentVariables();

// Create the express app, JSON middleware and port
const app = express();
app.use(express.json());
const port = process.env.PORT || 3978;

// Create the bot adapter
const adapter = await createBotAdapter();

// Home route
app.get("/", (req, res) => {
  res.send("Server is running");
});

// Webhook endpoint for bot messages
app.post("/api/messages", (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    if (context.activity.type === "message") {
      await handleMessages(context);
    } else {
      await handleTeamsActivity(context);
    }
  });
});

// Test route for sending messages to Teams
app.post("/api/test-teams-message", async (req, res) => {
  try {
    const { serviceUrl, conversationId, channelId, tenantId } =
      req.body.messageDetails;
    const message = req.body.message || "Processing results";
    const images = req.body.images || [];

    // Create an Adaptive Card with larger dimensions
    const card = {
      type: "AdaptiveCard",
      version: "1.4",
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      fallbackText: "Your client doesn't support Adaptive Cards.",
      speak: message,
      width: "full",
      body: [
        {
          type: "Container",
          width: "stretch",
          // minHeight: "100px",
          items: [
            {
              type: "TextBlock",
              text: message,
              size: "Large",
              weight: "Bolder",
              wrap: true,
            },
          ],
        },
        // Split images into groups of 3 and create multiple ColumnSets
        ...chunk(images, 3).map((imageGroup) => ({
          type: "ColumnSet",
          width: "stretch",
          spacing: "Large",
          columns: imageGroup.map((url) => ({
            type: "Column",
            width: "stretch",
            items: [
              {
                type: "Image",
                url: url,
                size: "Large",
                spacing: "None",
                horizontalAlignment: "Center",
              },
            ],
          })),
        })),
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "✅ Approve",
          data: {
            action: "approve_processing",
            value: "yes",
          },
          style: "positive",
        },
        {
          type: "Action.Submit",
          title: "❌ Reject",
          data: {
            action: "approve_processing",
            value: "no",
          },
          style: "destructive",
        },
      ],
    };

    // Helper function to split array into chunks
    function chunk(array, size) {
      const chunked = [];
      for (let i = 0; i < array.length; i += size) {
        chunked.push(array.slice(i, i + size));
      }
      return chunked;
    }

    // Create a reference to the conversation
    const conversationReference = {
      channelId: channelId,
      serviceUrl: serviceUrl,
      conversation: { id: conversationId },
      tenantId: tenantId,
    };

    // Use the adapter to continue the conversation and send the card
    await adapter.continueConversation(
      conversationReference,
      async (turnContext) => {
        if (images.length > 0) {
          // Send as an Adaptive Card
          await turnContext.sendActivity({
            attachments: [
              {
                contentType: "application/vnd.microsoft.card.adaptive",
                content: card,
              },
            ],
          });
        } else {
          // Send as a simple message if no images
          await turnContext.sendActivity(message);
        }
      }
    );

    res.status(200).json({ message: "Message sent successfully" });
  } catch (error) {
    console.error("Error sending test message:", error);
    res.status(500).json({ error: error.message });
  }
});

// Start the server
app.listen(port, () => {
  console.log(
    `\nBot is running on http://localhost:${port}/api/messages\nServer is running on http://localhost:${port}/`
  );
});
