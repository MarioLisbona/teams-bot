import { ChatOpenAI } from "@langchain/openai";
import { loadEnvironmentVariables } from "../lib/environment/setupEnvironment.js";

// Load environment variables first
loadEnvironmentVariables();

// Clear any Azure-related environment variables
// For some reason, the Azure OpenAI API Key needs to be deleted otherwise
// The OpenAI LLM fails to be built ???
// TODO: Investigate this further
delete process.env.AZURE_OPENAI_API_KEY;

// Only configure Azure in production mode
const isProduction = process.env.NODE_ENV === "production";

console.log(`🤖 Using ${isProduction ? "Azure OpenAI" : "OpenAI"} LLM Model:`, {
  provider: isProduction ? "Azure OpenAI" : "OpenAI",
  model: isProduction
    ? process.env.AZURE_OPENAI_DEPLOYMENT_NAME
    : "gpt-4-turbo-preview",
  environment: process.env.NODE_ENV,
});

export const llm = new ChatOpenAI({
  modelName: isProduction
    ? process.env.AZURE_OPENAI_DEPLOYMENT_NAME
    : "gpt-4-turbo-preview",
  temperature: 0,
  azure: isProduction
    ? {
        apiKey: process.env.AZURE_OPENAI_API_KEY,
        endpoint: process.env.AZURE_OPENAI_API_ENDPOINT,
        deploymentName: process.env.AZURE_OPENAI_DEPLOYMENT_NAME,
        apiVersion: process.env.AZURE_OPENAI_API_VERSION,
      }
    : undefined,
  apiKey: isProduction
    ? process.env.AZURE_OPENAI_API_KEY
    : process.env.OPENAI_API_KEY,
  configuration: isProduction
    ? undefined
    : {
        baseURL: "https://api.openai.com/v1",
        defaultHeaders: {
          "api-key": process.env.OPENAI_API_KEY,
        },
      },
});

export const formatLLMResponse = (result) => {
  return result
    .split("\n")
    .filter((line) => line.trim())
    .join("\n");
};
