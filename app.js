import express from "express";
import { createBotAdapter } from "./lib/utils/createBotAdapter.js";
import { handleMessages } from "./lib/handlers/handleTeamsMessages.js";
import { loadEnvironmentVariables } from "./lib/environment/setupEnvironment.js";
import { handleTeamsActivity } from "./lib/utils/teamsActivity.js";
import {
  handleValidateSignatures,
  handleValidateWorkflow,
  handleWorkflowProgress,
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
app.post("/api/validate/signatures", async (req, res) => {
  try {
    const { serviceUrl, conversationId, channelId, tenantId } =
      req.body.messageDetails;
    const message = req.body.message || "Please review these images";
    const images = req.body.images || [];

    await new Promise((resolve) => setTimeout(resolve, 3000));

    await handleValidateSignatures(
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
    const jobId = req.body.jobId;

    await handleValidateWorkflow(
      adapter,
      serviceUrl,
      conversationId,
      channelId,
      tenantId,
      validations,
      jobId
    );

    res.status(200).json({ message: "Validation card sent successfully" });
  } catch (error) {
    console.error("Error sending validation card:", error);
    res.status(500).json({ error: error.message });
  }
});

// Workflow progress route
app.post("/api/workflow/progress", async (req, res) => {
  try {
    const { serviceUrl, conversationId, channelId, tenantId } =
      req.body.messageDetails;
    const isCompleted = req.body.isCompleted || false;
    const workflowStep = req.body.workflowStep || {};
    const jobId = req.body.jobId;

    await handleWorkflowProgress(
      adapter,
      serviceUrl,
      conversationId,
      channelId,
      tenantId,
      isCompleted,
      workflowStep,
      jobId
    );

    res.status(200).json({ message: "Workflow progress updated successfully" });
  } catch (error) {
    console.error("Error updating workflow progress:", error);
    res.status(500).json({ error: error.message });
  }
});

// Start the server
app.listen(port, () => {
  console.log(
    `\nBot is running on http://localhost:${port}/api/messages\nServer is running on http://localhost:${port}/`
  );
});
