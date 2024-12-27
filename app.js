import express from "express";
import { createBotAdapter } from "./lib/utils/createBotAdapter.js";
import { loadEnvironmentVariables } from "./lib/environment/setupEnvironment.js";
import createMessageRoutes from "./lib/routes/messageRoutes.js";
import createValidateRoutes from "./lib/routes/validateRoutes.js";
import createNotificationRoutes from "./lib/routes/notificationRoutes.js";

// Load environment variables
loadEnvironmentVariables();

// Create the express app, JSON middleware and port
const app = express();
app.use(express.json());
const port = process.env.PORT || 3978;

// Create the bot adapter
const adapter = await createBotAdapter();

// Home route
app.get("/", (req, res) => {
  res.send("Server is running");
});

// Use the routes
app.use("/api", createMessageRoutes(adapter));
app.use("/api", createValidateRoutes(adapter));
app.use("/api", createNotificationRoutes(adapter));

// Start the server
app.listen(port, () => {
  console.log(
    `\nBot is running on http://localhost:${port}/api/messages\nServer is running on http://localhost:${port}/`
  );
});
