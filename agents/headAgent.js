import { initializeAgentExecutorWithOptions } from "langchain/agents";
import {
  listFolders,
  listExcelFiles,
  processTestingWorksheet,
  generateAuditorNotes,
  sendToWorkflowAgentTool,
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
        sendToWorkflowAgentTool,
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

        # Workflow Agent:
        If the message is similar to these example user messages:
        **"Start the audit for job id <job id>"**
        **"Start the audit for job ID <job id>"**
        **"Start the audit for jobID <job id>"**
        **"Start processing the evidence pack documents for job id <job id>"**
        **"Start processing the evidence pack docs for job id <job id>"**

        ### Steps:
        1. Notify the user that the audit job has sent to the workflow agent.
        2. Use sendToWorkflowAgentTool to send the message and context details to the workflow agent.

        # Processing Agent:
        If the message is similar to these example user messages:
        **"Process the rfi in the file <file name>"**
        **"Process the testing worksheet in the file <file name>"**
        **"Process RFI for <file name>"**

        The file name will be in this format: <client name> Testing.xlsx or a similar format with "Testing" in the file name.
        E.g - XXYY Testing.xlsx
        The client name in the file format will be the name of the client directory that the file is located in. 
        In the example above, the client name is XXYY and the directory name is XXYY.

        ### Steps:
        1. Use listFolders and listExcelFiles to find the workbook
        2. Create a selectedFileData object with id, directoryId, directoryName, and name
        3. Use processTestingWorksheet with the selectedFileData object

        > **Note:** If you can't find the requested client directory or exact filename, notify the user and do not process any other files.
        > **Note:** If you cant find the exact but find a similar file, notify the user of the similar filenames and do not process any other files.

        # Auditor Notes Agent:
        If the message is similar to these example user messages:
        **"Generate auditor notes for <file name>"**
        **"Process the rfi client responses in <file name>"**
        **"Process the client responses in <file name>"**

        The file name will be in this format: RFI Responses - <client name>.xlsx
        It may also be in this format for multiple files: RFI Responses - <client name>(number).xlsx
        and sometimes it will be in this format: RFI Responses - size - <client name>.xlsx 
        
        In this last format size is S, M, L, XL, etc.
        E.g - RFI Responses - S - MLD.xlsx
        The client name in the file format will be the name of the client directory that the file is located in.
        In the example above, the client name is MLD and the directory name is MLD.

        ### Steps:
        1. Use listFolders and listExcelFiles to find the workbook
        2. Use generateAuditorNotes with the context, filename, and workbookId

        > **Note:** If you can't find the requested client directory or exact filename, notify the user and do not process any other files.
        > **Note:** If you cant find the exact but find a similar file, notify the user of the similar filenames and do not process any other files.

        # User message: ${userMessage}
        
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
    console.log("Calling head agent");
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
