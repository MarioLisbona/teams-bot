import express from "express";
import {
  validateSignatures,
  humanValidationSteps,
} from "../handlers/handleAgentWorkFlow.js";

/**
 * Creates Express router for handling signature and human validation requests.
 *
 * @description
 * Sets up POST routes that:
 * 1. Handle signature validation requests with image review
 * 2. Process human validation steps for workflow tasks
 * 3. Send appropriate Teams notifications using the Bot Framework adapter
 * 4. Handle errors and provide appropriate HTTP responses
 *
 * @param {Object} adapter - Bot Framework adapter instance
 * @param {Function} adapter.continueConversation - Method to send proactive messages
 * @returns {Object} Express Router configured with validation routes
 */
const createValidateRoutes = (adapter) => {
  const validateRoutes = express.Router();

  validateRoutes.post("/validate/signatures", async (req, res) => {
    try {
      // Validate required fields in request body
      const { messageDetails, images } = req.body;
      if (!messageDetails || !Array.isArray(images)) {
        throw new Error("Missing or invalid required fields");
      }

      const { serviceUrl, conversationId, channelId, tenantId } =
        messageDetails;

      // Validate message details
      if (!serviceUrl || !conversationId || !channelId || !tenantId) {
        throw new Error("Missing required message details");
      }

      const message = req.body.message || "Please review these images";

      // Simulate validation process
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

      res.status(200).json({
        message: "Signature validation request sent successfully",
        imageCount: images.length,
      });
    } catch (error) {
      console.error("Error processing signature validation:", {
        error: error.message,
        body: req.body,
      });
      const statusCode = error.message.includes("Missing") ? 400 : 500;
      res.status(statusCode).json({
        error: error.message,
        details: "Failed to process signature validation request",
      });
    }
  });

  validateRoutes.post("/validate/human", async (req, res) => {
    try {
      // Validate required fields in request body
      const { messageDetails, jobId, validationsRequired } = req.body;
      if (!messageDetails || !jobId) {
        throw new Error("Missing required fields");
      }

      const { serviceUrl, conversationId, channelId, tenantId } =
        messageDetails;

      // Validate message details
      if (!serviceUrl || !conversationId || !channelId || !tenantId) {
        throw new Error("Missing required message details");
      }

      await humanValidationSteps(
        adapter,
        serviceUrl,
        conversationId,
        channelId,
        tenantId,
        validationsRequired || {},
        jobId
      );

      res.status(200).json({
        message: "Human validation request sent successfully",
        jobId,
        validationsRequired,
      });
    } catch (error) {
      console.error("Error processing human validation:", {
        error: error.message,
        body: req.body,
      });
      const statusCode = error.message.includes("Missing") ? 400 : 500;
      res.status(statusCode).json({
        error: error.message,
        details: "Failed to process human validation request",
      });
    }
  });

  return validateRoutes;
};

export default createValidateRoutes;
