import { initializeAgentExecutorWithOptions } from "langchain/agents";

import { loadEnvironmentVariables } from "../lib/environment/setupEnvironment.js";
import { createTeamsUpdate } from "../lib/utils/utils.js";
import { llm, formatLLMResponse } from "./index.js";
import {
  generateAuditorNotes,
  listFolders,
  listExcelFiles,
} from "./tools/index.js";
// Load environment variables first
loadEnvironmentVariables();

async function createAuditorNotesAgent() {
  try {
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
  } catch (error) {
    console.error("Failed to create auditor notes agent:", error);
    throw new Error("Failed to initialize agent executor");
  }
}

async function runAuditorNotesAgent(userMessage, context) {
  try {
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
    const formattedOutput = formatLLMResponse(result.output);

    await createTeamsUpdate(
      context,
      "Agent Response:",
      formattedOutput,
      "🤖",
      "default"
    );

    return result;
  } catch (error) {
    console.error("Error in auditor notes agent:", error);
    await createTeamsUpdate(
      context,
      "Error",
      "Sorry, there was an error processing your request. Please try again later.",
      "❌",
      "error"
    );
    throw error;
  }
}

export async function runAuditorNotes(userMessage, context) {
  try {
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
  } catch (error) {
    console.error("Error in runAuditorNotes:", error);
    // Only send error message if it hasn't been sent by runAuditorNotesAgent
    if (error.message !== "Failed to initialize agent executor") {
      await createTeamsUpdate(
        context,
        "Error",
        "An unexpected error occurred while processing your request.",
        "❌",
        "error"
      );
    }
  }
}
