import { ChatOpenAI } from "@langchain/openai";
import { initializeAgentExecutorWithOptions } from "langchain/agents";
import {
  listFolders,
  listExcelFiles,
  processTestingWorksheet,
} from "./tools/index.js";
import { loadEnvironmentVariables } from "../lib/environment/setupEnvironment.js";
import { createTeamsUpdate } from "../lib/utils/utils.js";

// Load environment variables first
loadEnvironmentVariables();

// Clear any Azure-related environment variables
delete process.env.AZURE_OPENAI_API_KEY;
delete process.env.AZURE_OPENAI_API_INSTANCE_NAME;
delete process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME;
delete process.env.AZURE_OPENAI_API_VERSION;
delete process.env.AZURE_OPENAI_API_ENDPOINT;

const llm = new ChatOpenAI({
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

async function createTestingProcessingAgent() {
  const executor = await initializeAgentExecutorWithOptions(
    [listFolders, listExcelFiles, processTestingWorksheet],
    llm,
    {
      agentType: "openai-functions", // Changed back to openai-functions
      verbose: true,
      maxIterations: 10,
    }
  );
  return executor;
}

async function runProcessingAgent(userMessage, context) {
  const agent = await createTestingProcessingAgent();

  // Create a wrapper function that will be properly serialized
  const wrappedContext = {
    sendActivity: async (text) => {
      return await context.sendActivity(text);
    },
    turnState: context.turnState || {},
    activity: context.activity || {},
  };

  console.log("Wrapped context:", {
    hasSendActivity: typeof wrappedContext.sendActivity === "function",
    hasTurnState: !!wrappedContext.turnState,
    hasActivity: !!wrappedContext.activity,
  });

  // Store the context in a closure that the tool can access
  global.teamsContext = wrappedContext;

  const result = await agent.invoke({
    input: `You are an assistant in a company that audits energy efficiency installations.
        Complete this task: ${userMessage}.
        Follow these steps strictly in order:
        1. Use listFolders and listExcelFiles to find the workbook with the exact filename
        2. Create an object called selectedFileData with the following properties:
        - id: the id of the file
        - directoryId: the id of the directory the file was found in
        - directoryName: the name of the directory the file was found in
        - name: the name of the file
        3. Use processTestingWorksheet with just the selectedFileData object
        `,
  });
  return result;
}

export async function runProcessing(userMessage, context) {
  console.log("Running processing agent");
  const result = await runProcessingAgent(userMessage, context);
  console.log(result);
  context.sendActivity(JSON.stringify(result.output));
}
