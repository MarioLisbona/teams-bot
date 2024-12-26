import express from "express";
import { workflowProgressNotification } from "../handlers/handleAgentWorkFlow.js";

const createNotificationRoutes = (adapter) => {
  const notificationRoutes = express.Router();

  notificationRoutes.post("/notification", async (req, res) => {
    try {
      const { serviceUrl, conversationId, channelId, tenantId } =
        req.body.messageDetails;
      const isComplete = req.body.isComplete || false;
      const workflowStep = req.body.workflowStep || {};
      const jobId = req.body.jobId;

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

      res
        .status(200)
        .json({ message: "Workflow progress updated successfully" });
    } catch (error) {
      console.error("Error updating workflow progress:", error);
      res.status(500).json({ error: error.message });
    }
  });

  return notificationRoutes;
};

export default createNotificationRoutes;
