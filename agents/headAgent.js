import { initializeAgentExecutorWithOptions } from "langchain/agents";
import {
  listFolders,
  listExcelFiles,
  processTestingWorksheet,
  generateAuditorNotes,
  // ... any other tools
} from "./tools/index.js";
import { createTeamsUpdate } from "../lib/utils/utils.js";
import { llm, formatLLMResponse } from "./index.js";

async function createHeadAgent() {
  try {
    const executor = await initializeAgentExecutorWithOptions(
      [
        listFolders,
        listExcelFiles,
        processTestingWorksheet,
        generateAuditorNotes,
      ],
      llm,
      {
        agentType: "openai-functions",
        verbose: true,
        maxIterations: 10,
      }
    );
    return executor;
  } catch (error) {
    console.error("Failed to create head agent:", error);
    throw new Error("Failed to initialize agent executor");
  }
}

async function runHeadAgent(userMessage, context) {
  try {
    const agent = await createHeadAgent();

    const wrappedContext = {
      sendActivity: async (text) => await context.sendActivity(text),
      turnState: context.turnState || {},
      activity: context.activity || {},
    };

    global.teamsContext = wrappedContext;

    const result = await agent.invoke({
      input: `You are an AI assistant in a company that audits energy efficiency installations.
        Based on the user's message, determine which task needs to be performed and execute it:

        If the message is similar to these example user messages:
        **"Process the rfi in the file <file name>"**
        **"Process the testing worksheet in the file <file name>"**
        **"Process RFI for <file name>"**

        The file name will be in this format: <client name> <file name>.xlsx
        E.g - XXYY Testing.xlsx
        The client name in the file format will be the name of the client directory that the file is in.

        Follow these steps:
        1. Use listFolders and listExcelFiles to find the workbook
        2. Create a selectedFileData object with id, directoryId, directoryName, and name
        3. Use processTestingWorksheet with the selectedFileData object

        If the message is similar to these example user messages:
        **"Generate auditor notes for <file name>"**
        **"Process the rfi client responses in <file name>"**
        **"Process the client responses in <file name>"**

        The file name will be in this format: RFI Responses - <client name>.xlsx
        It may also be in this format for multiple files: RFI Responses - <client name>(number).xlsx
        and sometimes it will be in this format: RFI Responses - size - <client name>.xlsx 
        
        In this last format size is S, M, L, XL, etc.
        E.g - RFI Responses - S - XXYY.xlsx

        The client name in the file format will be the name of the client directory that the file is in.

        Follow these steps:
        1. Use listFolders and listExcelFiles to find the workbook
        2. Use generateAuditorNotes with the context, filename, and workbookId

        User message: ${userMessage}
        
        If you can't find the requested file, notify the user.
        Execute only the relevant task based on the user's message.`,
    });

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
    console.error("Error in head agent:", error);
    await createTeamsUpdate(
      context,
      "Error",
      "Sorry, there was an error processing your request.",
      "❌",
      "error"
    );
    throw error;
  }
}

export async function callHeadAgent(userMessage, context) {
  try {
    console.log("Running head agent");
    await createTeamsUpdate(
      context,
      "Processing Request",
      userMessage,
      "🤖",
      "default"
    );
    const result = await runHeadAgent(userMessage, context);
    console.log(result);
  } catch (error) {
    console.error("Error in handleMessage:", error);
    if (error.message !== "Failed to initialize agent executor") {
      await createTeamsUpdate(
        context,
        "Error",
        "An unexpected error occurred.",
        "❌",
        "error"
      );
    }
  }
}
