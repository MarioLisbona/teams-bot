import express from "express";
import { handleMessages } from "../handlers/handleTeamsMessages.js";
import { handleTeamsActivity } from "../utils/teamsActivity.js";

/**
 * Creates Express router for handling Teams bot messages and activities.
 *
 * @description
 * Sets up a POST route that:
 * 1. Processes incoming Teams activities using the Bot Framework adapter
 * 2. Routes messages to appropriate handlers based on activity type
 * 3. Provides error handling for the entire message processing pipeline
 *
 * @param {Object} adapter - Bot Framework adapter instance
 * @param {Function} adapter.processActivity - Method to process Teams activities
 * @returns {Object} Express Router configured with message handling routes
 */
const createMessageRoutes = (adapter) => {
  const messageRoutes = express.Router();

  messageRoutes.post("/messages", (req, res) => {
    try {
      adapter.processActivity(req, res, async (context) => {
        try {
          if (context.activity.type === "message") {
            await handleMessages(context);
          } else {
            await handleTeamsActivity(context);
          }
        } catch (error) {
          console.error("Error processing activity:", error);
          // Attempt to notify user of error
          try {
            await context.sendActivity(
              "❌ An error occurred while processing your request. Please try again."
            );
          } catch (notifyError) {
            console.error("Failed to send error notification:", notifyError);
          }
        }
      });
    } catch (error) {
      console.error("Failed to process Teams activity:", error);
      res.status(500).send("Internal Server Error");
    }
  });

  return messageRoutes;
};

export default createMessageRoutes;
