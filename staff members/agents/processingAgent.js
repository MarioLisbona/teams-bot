import { initializeAgentExecutorWithOptions } from "langchain/agents";
import {
  listFoldersTool,
  listExcelFilesTool,
  processTestingWorksheetTool,
} from "../tools/index.js";
import { createTeamsUpdate } from "../../lib/utils/utils.js";
import { llm, formatLLMResponse } from "../index.js";

// function to create the executor agent
async function createTestingProcessingAgent() {
  try {
    const executor = await initializeAgentExecutorWithOptions(
      [listFoldersTool, listExcelFilesTool, processTestingWorksheetTool],
      llm,
      {
        agentType: "openai-functions",
        verbose: true,
        maxIterations: 10,
      }
    );
    return executor;
  } catch (error) {
    console.error("Failed to create testing processing agent:", error);
    throw new Error("Failed to initialize agent executor");
  }
}

// function to invoke the agent with the user message and prompt instructions

async function runProcessingAgent(userMessage, context) {
  try {
    const agent = await createTestingProcessingAgent();

    // Create a wrapper function that will be properly serialized
    // The context is needed to send messages back to Teams from inside the processTestingWorksheet tool
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
    console.error("Error in processing agent:", error);
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

export async function runProcessing(userMessage, context) {
  try {
    console.log("Running processing agent");
    await createTeamsUpdate(
      context,
      `Querying the RFI Processing Agent`,
      userMessage,
      "🤖",
      "default"
    );
    const result = await runProcessingAgent(userMessage, context);
    console.log(result);
  } catch (error) {
    console.error("Error in runProcessing:", error);
    // Only send error message if it hasn't been sent by runProcessingAgent
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
