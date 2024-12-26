import express from "express";
import { updateWorkflowProgress } from "../handlers/handleAgentWorkFlow.js";

const updateRoutes = express.Router();

updateRoutes.post("/update", async (req, res) => {
  try {
    const { serviceUrl, conversationId, channelId, tenantId } =
      req.body.messageDetails;
    const isComplete = req.body.isComplete || false;
    const workflowStep = req.body.workflowStep || {};
    const jobId = req.body.jobId;

    await updateWorkflowProgress(
      req.adapter,
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

export default updateRoutes;
