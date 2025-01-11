import { ChatOpenAI } from "@langchain/openai";
import { initializeAgentExecutorWithOptions } from "langchain/agents";
import { DynamicStructuredTool } from "@langchain/core/tools";
import { z } from "zod";
import { loadEnvironmentVariables } from "../lib/environment/setupEnvironment.js";
import {
  listFoldersInDirectory,
  getFileNamesAndIds,
} from "../lib/utils/fileStorageAndRetrieval.js";

// Load environment variables first
loadEnvironmentVariables();

// Debug: Check OpenAI API key
console.log("OpenAI Configuration Check:", {
  hasKey: !!process.env.OPENAI_API_KEY,
  keyPrefix: process.env.OPENAI_API_KEY?.substring(0, 5) + "...",
});

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

// Debug: Confirm LLM configuration
console.log("LLM Configuration:", {
  temperature: llm.temperature,
  modelName: "gpt-4-turbo-preview",
  hasKey: !!process.env.OPENAI_API_KEY,
});

const locateWorkbook = new DynamicStructuredTool({
  name: "locateWorkbook",
  description: "Locate a workbook in the user's directory",
  schema: z.object({
    workbookName: z.string().describe("The name of the workbook to locate"),
  }),
  func: async ({ workbookName }) => {
    return `Workbook ${workbookName} located`;
  },
});

const locateSheet = new DynamicStructuredTool({
  name: "locateSheet",
  description: "Locate a sheet in a workbook",
  schema: z.object({
    sheetName: z.string().describe("The name of the sheet to locate"),
    workbookName: z.string().describe("The name of the workbook"),
  }),
  func: async ({ sheetName, workbookName }) => {
    return `Sheet ${sheetName} located in workbook ${workbookName}`;
  },
});

const processTestingWorksheet = new DynamicStructuredTool({
  name: "processTestingWorksheet",
  description: "Process a testing worksheet",
  schema: z.object({
    sheetName: z.string().describe("The name of the sheet to process"),
    workbookName: z.string().describe("The name of the workbook"),
  }),
  func: async ({ sheetName, workbookName }) => {
    return `Processing sheet ${sheetName} in workbook ${workbookName}`;
  },
});

export const listFolders = new DynamicStructuredTool({
  name: "listFolders",
  description: "List all folders in a specified SharePoint directory",
  schema: z.object({
    folderId: z
      .string()
      .optional()
      .default("")
      .describe(
        "Optional folder ID to list subfolders. If not provided, lists root folders"
      ),
  }),
  func: async () => {
    try {
      const folders = await listFoldersInDirectory();

      if (!folders || folders.length === 0) {
        return "No folders found in the specified directory.";
      }

      const folderList = folders
        .map((f) => `- ${f.name} (ID: ${f.id})`)
        .join("\n");
      return `Found ${folders.length} folders:\n${folderList}`;
    } catch (error) {
      return `Error listing folders: ${error.message}`;
    }
  },
});

export const listExcelFiles = new DynamicStructuredTool({
  name: "listExcelFiles",
  description: "List all Excel files in a specified SharePoint directory",
  schema: z.object({
    folderId: z
      .string()
      .optional()
      .default("")
      .describe(
        "Optional folder ID to list Excel files from. If not provided, lists from root folder"
      ),
  }),
  func: async ({ folderId }) => {
    try {
      const files = await getFileNamesAndIds(folderId);

      if (!files || files.length === 0) {
        return "No Excel files found in the specified directory.";
      }

      const fileList = files.map((f) => `- ${f.name} (ID: ${f.id})`).join("\n");
      return `Found ${files.length} Excel files:\n${fileList}`;
    } catch (error) {
      return `Error listing Excel files: ${error.message}`;
    }
  },
});

async function createTestingProcessingAgent() {
  const executor = await initializeAgentExecutorWithOptions(
    [listFolders, listExcelFiles],
    llm,
    {
      agentType: "openai-functions", // Changed back to openai-functions
      verbose: true,
      maxIterations: 5,
    }
  );
  return executor;
}

// async function runProcessingAgent(auditTask) {
//   const agent = await createTestingProcessingAgent();

//   // Extract workbook info from user message if provided, or use defaults
//   const workbookInfo = {
//     clientName: "Example Client",
//     auditType: "Testing",
//     date: new Date().toISOString().split("T")[0],
//   };

//   const result = await agent.invoke({
//     input: `Complete this audit task: ${auditTask}.
//     Follow these steps strictly in order:
//     1. Find the correct workbook and sheet for the client
//      - An example structure of the filename will be: "${workbookInfo.clientName} - ${workbookInfo.auditType} - ${workbookInfo.date}.xlsx"
//     2. Use locateWorkbook to find the workbook with the exact filename
//     3. Use locateSheet to find the sheet, passing both sheet name and workbook name
//     4. Use processTestingWorksheet to process the sheet, using the workbook name found
//     5. Return the final report as your response.

//     Context:
//     - Workbook Name: "${workbookInfo.clientName} - ${workbookInfo.auditType} - ${workbookInfo.date}.xlsx"
//     - Expected Sheet: "Testing"

//     Important: After generating the report, conclude the task and return the results.`,
//     workbookInfo: workbookInfo, // Pass workbook info as context
//   });
//   return result;
// }
async function runProcessingAgent(auditTask) {
  const agent = await createTestingProcessingAgent();

  const result = await agent.invoke({
    input: `Complete this task: ${auditTask}. 
    Follow these steps strictly in order:
    1. Find the client directory by using listFolders to get all the folders in the user's directory.
    2. Find the file by using listFiles function passing the client folder id
    
    This is naming structure of the Testing files:
    - "<client> - Testing.xlsx"
`,
  });
  return result;
}

export async function runProcessing() {
  console.log("Running processing agent");
  const result = await runProcessingAgent(
    "What is the id of the file named 'XXYY - Testing small.xlsx'?"
  );
  console.log(result);
}

runProcessing();
