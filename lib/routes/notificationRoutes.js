import express from "express";
import { workflowProgressNotification } from "../handlers/handleAgentWorkFlow.js";

/**
 * Creates Express router for handling workflow progress notifications.
 *
 * @description
 * Sets up a POST route that:
 * 1. Receives workflow progress updates from the Workflow Agent
 * 2. Validates and extracts required notification details
 * 3. Sends notifications to Teams conversations using the Bot Framework adapter
 * 4. Handles errors and provides appropriate HTTP responses
 *
 * @param {Object} adapter - Bot Framework adapter instance
 * @param {Function} adapter.continueConversation - Method to send proactive messages
 * @returns {Object} Express Router configured with notification handling routes
 */
const createNotificationRoutes = (adapter) => {
  const notificationRoutes = express.Router();

  notificationRoutes.post("/notification", async (req, res) => {
    try {
      // Validate required fields in request body
      const { messageDetails, jobId } = req.body;
      if (!messageDetails || !jobId) {
        throw new Error("Missing required notification details");
      }

      const { serviceUrl, conversationId, channelId, tenantId } =
        messageDetails;

      // Validate message details
      if (!serviceUrl || !conversationId || !channelId || !tenantId) {
        throw new Error("Missing required message details");
      }

      const isComplete = req.body.isComplete || false;
      const workflowStep = req.body.workflowStep || {};

      await workflowProgressNotification(
        adapter,
        serviceUrl,
        conversationId,
        channelId,
        tenantId,
        isComplete,
        workflowStep,
        jobId
      );

      res.status(200).json({
        message: "Workflow progress updated successfully",
        jobId: jobId,
      });
    } catch (error) {
      console.error("Error processing workflow notification:", {
        error: error.message,
        body: req.body,
      });

      // Determine appropriate status code based on error
      const statusCode = error.message.includes("Missing required") ? 400 : 500;

      res.status(statusCode).json({
        error: error.message,
        details: "Failed to send workflow notification",
      });
    }
  });

  return notificationRoutes;
};

export default createNotificationRoutes;
