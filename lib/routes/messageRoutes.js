import express from "express";
import { handleMessages } from "../handlers/handleTeamsMessages.js";
import { handleTeamsActivity } from "../utils/teamsActivity.js";

const createMessageRoutes = (adapter) => {
  const messageRoutes = express.Router();

  messageRoutes.post("/messages", (req, res) => {
    adapter.processActivity(req, res, async (context) => {
      if (context.activity.type === "message") {
        await handleMessages(context);
      } else {
        await handleTeamsActivity(context);
      }
    });
  });

  return messageRoutes;
};

export default createMessageRoutes;
