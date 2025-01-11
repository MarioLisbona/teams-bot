import { initializeAgentExecutorWithOptions } from "langchain/agents";

import { loadEnvironmentVariables } from "../lib/environment/setupEnvironment.js";
import { createTeamsUpdate } from "../lib/utils/utils.js";
import { llm } from "./index.js";
import {
  generateAuditorNotes,
  listFolders,
  listExcelFiles,
} from "./tools/index.js";
// Load environment variables first
loadEnvironmentVariables();

async function createAuditorNotesAgent() {
  const executor = await initializeAgentExecutorWithOptions(
    [listFolders, listExcelFiles, generateAuditorNotes],
    llm,
    {
      agentType: "openai-functions",
      verbose: true,
      maxIterations: 10,
    }
  );
  return executor;
}

async function runAuditorNotesAgent(userMessage, context) {
  const agent = await createAuditorNotesAgent();

  // Create a wrapper function that will be properly serialized
  const wrappedContext = {
    sendActivity: async (text) => {
      return await context.sendActivity(text);
    },
    turnState: context.turnState || {},
    activity: context.activity || {},
  };

  // Store the context in a closure that the tool can access
  global.teamsContext = wrappedContext;

  const result = await agent.invoke({
    input: `You are an AI assistant in a company that audits energy efficiency installations.
        Complete this task: ${userMessage}.
        Follow these steps strictly in order:
        1. Use listFolders and listExcelFiles to find the workbook with the exact filename
          - The filename is the name of the workbook that the user is asking for
          - The workbookId is the id of the workbook that the user is asking for
          - The file name contains the client name which is also the directory name
          For example: "RFI Responses (Back testing Medium) - MLD" - MLD is the client name
        2. Use generateAuditorNotes passing in the context, filename, and workbookId of the located workbook
        `,
  });

  // Format the output for Teams
  const formattedOutput = result.output
    .split("\n")
    .filter((line) => line.trim()) // Remove empty lines
    .map((line) => line.trim()) // Remove extra whitespace
    .join("\n"); // Rejoin with newlines

  await createTeamsUpdate(
    context,
    "Agent Response:",
    formattedOutput,
    "🤖",
    "default"
  );

  return result;
}

export async function runAuditorNotes(userMessage, context) {
  console.log("Running auditor notes agent");
  await createTeamsUpdate(
    context,
    `Querying the auditor notes Agent`,
    userMessage,
    "🤖",
    "default"
  );
  const result = await runAuditorNotesAgent(userMessage, context);
  console.log(result);
}
