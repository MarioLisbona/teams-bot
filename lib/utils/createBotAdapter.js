// Create the BotFrameworkAdapter
import { BotFrameworkAdapter } from "botbuilder";
import { loadEnvironmentVariables } from "../environment/setupEnvironment.js";

loadEnvironmentVariables();

export const createBotAdapter = async () => {
  const adapter = new BotFrameworkAdapter({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD,
  });

  // Error handling
  adapter.onTurnError = async (context, error) => {
    console.error(`[onTurnError]: ${error}`);
    await context.sendActivity("Oops, something went wrong!");
  };

  return adapter;
};
