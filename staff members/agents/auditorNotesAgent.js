import { initializeAgentExecutorWithOptions } from "langchain/agents";

import { loadEnvironmentVariables } from "../../lib/environment/setupEnvironment.js";
import { createTeamsUpdate } from "../../lib/utils/utils.js";
import { llm, formatLLMResponse } from "../index.js";
import {
  generateAuditorNotesTool,
  listFoldersTool,
  listExcelFilesTool,
} from "../tools/index.js";

// function to create the executor agent
async function createAuditorNotesAgent() {
  try {
    const executor = await initializeAgentExecutorWithOptions(
      [listFoldersTool, listExcelFilesTool, generateAuditorNotesTool],
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

// function to invoke the agent with the user message and prompt instructions
async function runAuditorNotesAgent(userMessage, context) {
  try {
    const agent = await createAuditorNotesAgent();

    // Create a wrapper function that will be properly serialized
    // The context is needed to send messages back to Teams from inside the generateAuditorNotes tool
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

        If you are unable to find the file, notify the user that the file was not found.
        `,
    });

    // Format the output for Teams
    const formattedOutput = formatLLMResponse(result.output);

    // Send the formatted Agent output back to Teams
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

// function to run the agent with the user message
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
