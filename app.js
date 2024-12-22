import express from "express";
import { createBotAdapter } from "./lib/utils/createBotAdapter.js";
import { handleMessages } from "./lib/handlers/handleMessages.js";
import { loadEnvironmentVariables } from "./lib/environment/setupEnvironment.js";
import { handleTeamsActivity } from "./lib/utils/teamsActivity.js";
import {
  createProcessingResultsCard,
  createWorkflow1ValidationCard,
} from "./lib/utils/adaptiveCards.js";
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
    const message = req.body.message || "Please review these images";
    const images = req.body.images || [];

    // Helper function to split array into chunks
    function chunk(array, size) {
      const chunked = [];
      for (let i = 0; i < array.length; i += size) {
        chunked.push(array.slice(i, i + size));
      }
      return chunked;
    }

    // Create initial card with images and approve/reject buttons
    const reviewCard = {
      type: "AdaptiveCard",
      version: "1.0",
      body: [
        {
          type: "TextBlock",
          text: message,
          size: "medium",
          weight: "bolder",
        },
        ...chunk(images, 3).map((imageChunk) => ({
          type: "ColumnSet",
          columns: imageChunk.map((url) => ({
            type: "Column",
            width: "stretch",
            items: [
              {
                type: "Image",
                url: url,
                size: "stretch",
                height: "200px",
              },
            ],
          })),
        })),
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "✅ Approve",
          style: "positive",
          data: {
            action: "approve",
            images: images,
          },
        },
        {
          type: "Action.Submit",
          title: "❌ Reject",
          style: "destructive",
          data: {
            action: "reject",
            images: images,
          },
        },
      ],
    };

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
        await turnContext.sendActivity({
          attachments: [
            {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: reviewCard,
            },
          ],
        });
      }
    );

    res.status(200).json({ message: "Message sent successfully" });
  } catch (error) {
    console.error("Error sending test message:", error);
    res.status(500).json({ error: error.message });
  }
});

// Test route for validation card
app.post("/api/validation", async (req, res) => {
  try {
    const { serviceUrl, conversationId, channelId, tenantId } =
      req.body.messageDetails;
    const validations = req.body.validations || {
      nomForm: false,
      siteAssessment: false,
      taxInvoice: false,
      ccew: false,
      installerDec: false,
      coc: false,
      gtp: false,
    };

    const validationCard = createWorkflow1ValidationCard(validations);

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
        await turnContext.sendActivity({
          attachments: [
            {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: validationCard,
            },
          ],
        });
      }
    );

    res.status(200).json({ message: "Validation card sent successfully" });
  } catch (error) {
    console.error("Error sending validation card:", error);
    res.status(500).json({ error: error.message });
  }
});

// Start the server
app.listen(port, () => {
  console.log(
    `\nBot is running on http://localhost:${port}/api/messages\nServer is running on http://localhost:${port}/`
  );
});
