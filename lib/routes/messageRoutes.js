import express from "express";
import { handleMessages } from "../handlers/handleTeamsMessages.js";
import { handleTeamsActivity } from "../utils/teamsActivity.js";

const messageRoutes = express.Router();

messageRoutes.post("/messages", (req, res) => {
  req.adapter.processActivity(req, res, async (context) => {
    if (context.activity.type === "message") {
      await handleMessages(context);
    } else {
      await handleTeamsActivity(context);
    }
  });
});

export default messageRoutes;
