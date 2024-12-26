import express from "express";
import {
  validateSignatures,
  humanValidationSteps,
} from "../handlers/handleAgentWorkFlow.js";

const validateRoutes = express.Router();

// Signature validation route
validateRoutes.post("/validate/signatures", async (req, res) => {
  try {
    const { serviceUrl, conversationId, channelId, tenantId } =
      req.body.messageDetails;
    const message = req.body.message || "Please review these images";
    const images = req.body.images || [];

    await new Promise((resolve) => setTimeout(resolve, 3000));

    await validateSignatures(
      req.adapter,
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
    const { serviceUrl, conversationId, channelId, tenantId } =
      req.body.messageDetails;
    const validationsRequired = req.body.validationsRequired || {};
    const jobId = req.body.jobId;

    await humanValidationSteps(
      req.adapter,
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

export default validateRoutes;
