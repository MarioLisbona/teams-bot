import express from "express";
import { createBotAdapter } from "./lib/utils/createBotAdapter.js";
import { handleMessages } from "./lib/handlers/handleTeamsMessages.js";
import { loadEnvironmentVariables } from "./lib/environment/setupEnvironment.js";
import { handleTeamsActivity } from "./lib/utils/teamsActivity.js";
import validateRoutes from "./lib/routes/validateRoutes.js";
import updateRoutes from "./lib/routes/updateRoutes.js";

// Load environment variables
loadEnvironmentVariables();

// Create the express app, JSON middleware and port
const app = express();
app.use(express.json());
const port = process.env.PORT || 3978;

// Create the bot adapter
const adapter = await createBotAdapter();

// Middleware to inject the adapter
app.use("/api", (req, res, next) => {
  req.adapter = adapter;
  next();
});

// Middleware to inject the adapter for workflow routes
app.use("/api/workflow", (req, res, next) => {
  req.adapter = adapter;
  next();
});

// Home route
app.get("/", (req, res) => {
  res.send("Server is running");
});

// Use the routes
app.use("/api/workflow", validateRoutes);
app.use("/api/workflow", updateRoutes);

// Webhook endpoint for bot messages
app.post("/api/messages", (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    if (context.activity.type === "message") {
      await handleMessages(context);
    } else {
      await handleTeamsActivity(context);
    }
  });
});

// Start the server
app.listen(port, () => {
  console.log(
    `\nBot is running on http://localhost:${port}/api/messages\nServer is running on http://localhost:${port}/`
  );
});
