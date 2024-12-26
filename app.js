import express from "express";
import { createBotAdapter } from "./lib/utils/createBotAdapter.js";
import { handleMessages } from "./lib/handlers/handleTeamsMessages.js";
import { loadEnvironmentVariables } from "./lib/environment/setupEnvironment.js";
import { handleTeamsActivity } from "./lib/utils/teamsActivity.js";
import {
  validateSignatures,
  handleWorkflowProgress,
  handleHumanWorkflowValidationUI,
} from "./lib/handlers/handleAgentWorkFlow.js";

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
app.post("/api/workflow/validate/signatures", async (req, res) => {
  try {
    const { serviceUrl, conversationId, channelId, tenantId } =
      req.body.messageDetails;
    const message = req.body.message || "Please review these images";
    const images = req.body.images || [];

    await new Promise((resolve) => setTimeout(resolve, 3000));

    await validateSignatures(
      adapter,
      message,
      serviceUrl,
      conversationId,
      channelId,
      tenantId,
      images
    );

    res.status(200).json({ message: "Message sent successfully" });
  } catch (error) {
    console.error("Error sending test message:", error);
    res.status(500).json({ error: error.message });
  }
});

// Workflow progress route
app.post("/api/workflow/progress", async (req, res) => {
  try {
    const { serviceUrl, conversationId, channelId, tenantId } =
      req.body.messageDetails;
    const isComplete = req.body.isComplete || false;
    const workflowStep = req.body.workflowStep || {};
    const jobId = req.body.jobId;

    await handleWorkflowProgress(
      adapter,
      serviceUrl,
      conversationId,
      channelId,
      tenantId,
      isComplete,
      workflowStep,
      jobId
    );

    res.status(200).json({ message: "Workflow progress updated successfully" });
  } catch (error) {
    console.error("Error updating workflow progress:", error);
    res.status(500).json({ error: error.message });
  }
});

// Workflow validation route
app.post("/api/workflow/validate", async (req, res) => {
  try {
    const { serviceUrl, conversationId, channelId, tenantId } =
      req.body.messageDetails;
    const validationsRequired = req.body.validationsRequired || {};

    const jobId = req.body.jobId;

    await handleHumanWorkflowValidationUI(
      adapter,
      serviceUrl,
      conversationId,
      channelId,
      tenantId,
      validationsRequired,
      jobId
    );

    // TODO: Add your validation handling logic here
    // For now, just return success
    res.status(200).json({
      message: "Validation request received successfully",
      jobId,
      validationsRequired,
    });
  } catch (error) {
    console.error("Error processing validation request:", error);
    res.status(500).json({ error: error.message });
  }
});

// Start the server
app.listen(port, () => {
  console.log(
    `\nBot is running on http://localhost:${port}/api/messages\nServer is running on http://localhost:${port}/`
  );
});
