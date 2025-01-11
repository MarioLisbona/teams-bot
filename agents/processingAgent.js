import { ChatOpenAI } from "@langchain/openai";
import { initializeAgentExecutorWithOptions } from "langchain/agents";
import { listFolders, listExcelFiles } from "./tools/index.js";
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
    [listFolders, listExcelFiles],
    llm,
    {
      agentType: "openai-functions", // Changed back to openai-functions
      verbose: true,
      maxIterations: 10,
    }
  );
  return executor;
}

async function runProcessingAgent(auditTask) {
  const agent = await createTestingProcessingAgent();

  const result = await agent.invoke({
    input: `You are an assistant in a company that audits energy efficiency installations.
        Complete this task: ${auditTask}.
        Follow these steps strictly in order:
        1. Use listFolders and listExcelFiles to find the workbook with the exact filename
        2. Create an object called selectedFileData with the following properties:
        - id: the id of the file
        - directoryId: the id of the directory the file was found in
        - directoryName: the name of the directory the file was found in
        - name: the name of the file
        3. Return the selectedFileData object as your response in addition to answering the question.
        `,
  });
  return result;
}

const userMessage =
  "What is the id of the file named 'XXYY - Testing.xlsx'? Can you tell me the ID of the directory the file was found in?";

export async function runProcessing(context, userMessage) {
  console.log("Running processing agent");
  const result = await runProcessingAgent(userMessage);
  console.log(result);
  await createTeamsUpdate(
    context,
    JSON.stringify(result),
    "",
    "🤖",
    "attention"
  );
}
