import express from "express";
import { createBotAdapter } from "./lib/utils/createBotAdapter.js";
import { loadEnvironmentVariables } from "./lib/environment/setupEnvironment.js";
import createMessageRoutes from "./lib/routes/messageRoutes.js";
import createValidateRoutes from "./lib/routes/validateRoutes.js";
import createNotificationRoutes from "./lib/routes/notificationRoutes.js";

// Global error handling for uncaught exceptions and unhandled rejections
process.on("uncaughtException", (error) => {
  console.error("Uncaught Exception:", error);
  console.error("Stack:", error.stack);
});

process.on("unhandledRejection", (reason, promise) => {
  console.error("Unhandled Rejection at:", promise);
  console.error("Reason:", reason);
});

console.log("Starting application...");
console.log("Environment:", process.env.NODE_ENV);
console.log("Current working directory:", process.cwd());

try {
  // Load environment variables
  loadEnvironmentVariables();
  console.log("Environment variables loaded successfully");

  // Create the express app, JSON middleware and port
  const app = express();
  app.use(express.json());
  const port = process.env.PORT || 3978;
  const host = process.env.HOST || "localhost";
  const baseUrl = `http://${host}:${port}`;

  // Create the bot adapter
  const adapter = await createBotAdapter();

  // Home route
  app.get("/", (req, res) => {
    res.send("Server is running");
  });

  // Workflow agent route
  app.post("/workflow-agent", (req, res) => {
    const { serviceUrl, conversationId, channelId, tenantId, userMessage } =
      req.body;

    console.log("Workflow Agent Request Received from Teams:", {
      serviceUrl,
      conversationId,
      channelId,
      tenantId,
      userMessage,
    });

    res.status(200).json({ message: "Workflow request received" });
  });

  // Use the routes
  app.use("/api", createMessageRoutes(adapter));
  app.use("/api", createValidateRoutes(adapter));
  app.use("/api", createNotificationRoutes(adapter));

  // Start the server
  app.listen(port, () => {
    console.log(`Bot is running on ${baseUrl}/api/messages`);
    console.log(`Server is running on ${baseUrl}`);
  });
} catch (error) {
  console.error("Startup error:", error);
  console.error("Stack:", error.stack);
  process.exit(1);
}
