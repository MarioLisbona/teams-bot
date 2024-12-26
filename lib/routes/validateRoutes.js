import express from "express";
import {
  validateSignatures,
  humanValidationSteps,
} from "../handlers/handleAgentWorkFlow.js";

const createValidateRoutes = (adapter) => {
  const validateRoutes = express.Router();

  validateRoutes.post("/validate/signatures", async (req, res) => {
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

  validateRoutes.post("/validate/human", async (req, res) => {
    try {
      const { serviceUrl, conversationId, channelId, tenantId } =
        req.body.messageDetails;
      const validationsRequired = req.body.validationsRequired || {};
      const jobId = req.body.jobId;

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
