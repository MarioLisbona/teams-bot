import { ChatOpenAI } from "@langchain/openai";
import { loadEnvironmentVariables } from "../lib/environment/setupEnvironment.js";

// Load environment variables first
loadEnvironmentVariables();

// Clear any Azure-related environment variables
delete process.env.AZURE_OPENAI_API_KEY;
delete process.env.AZURE_OPENAI_API_INSTANCE_NAME;
delete process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME;
delete process.env.AZURE_OPENAI_API_VERSION;
delete process.env.AZURE_OPENAI_API_ENDPOINT;

export const llm = new ChatOpenAI({
  modelName: "gpt-4-turbo-preview",
  temperature: 0,
  apiKey: process.env.OPENAI_API_KEY,
  configuration: {
    baseURL: "https://api.openai.com/v1",
    defaultHeaders: {
      "api-key": process.env.OPENAI_API_KEY,
    },
  },
});
