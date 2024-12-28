import express from "express";
import {
  validateSignatures,
  humanValidationSteps,
} from "../handlers/handleAgentWorkFlow.js";

const createValidateRoutes = (adapter) => {
  const validateRoutes = express.Router();

  // Validate the signatures
  validateRoutes.post("/validate/signatures", async (req, res) => {
    try {
      // Extract the service URL, conversation ID, channel ID, and tenant ID from the message details
      const { serviceUrl, conversationId, channelId, tenantId } =
        req.body.messageDetails;
      // Extract the message and images from the request body
      const message = req.body.message || "Please review these images";
      const images = req.body.images || [];

      // Simulate validation process
      await new Promise((resolve) => setTimeout(resolve, 3000));

      // Validate the signatures
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

  // Human validation route
  validateRoutes.post("/validate/human", async (req, res) => {
    try {
      // Extract the validation type, current validations, completed validations, and job ID from the activity value
      const { serviceUrl, conversationId, channelId, tenantId } =
        req.body.messageDetails;
      const validationsRequired = req.body.validationsRequired || {};
      const jobId = req.body.jobId;

      // Trigger the human validation steps
      await humanValidationSteps(
        adapter,
        serviceUrl,
        conversationId,
        channelId,
        tenantId,
        validationsRequired,
        jobId
      );

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

  return validateRoutes;
};

export default createValidateRoutes;
