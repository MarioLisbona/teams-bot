import { DynamicStructuredTool } from "@langchain/core/tools";
import { z } from "zod";
import {
  listFoldersInDirectory,
  getFileNamesAndIds,
} from "../../lib/utils/fileStorageAndRetrieval.js";
import { processRfiWorksheet } from "../../lib/utils/auditProcessing.js";
import { processAuditorNotes } from "../../lib/handlers/handleAuditWorkbook.js";
import { sendToWorkflowAgent } from "../index.js";
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

export const processTestingWorksheet = new DynamicStructuredTool({
  name: "processTestingWorksheet",
  description:
    "Process a Testing worksheet and generate an RFI Response workbook",
  schema: z.object({
    selectedFileData: z
      .object({
        id: z.string().describe("The SharePoint ID of the workbook"),
        directoryId: z
          .string()
          .describe("The SharePoint ID of the parent directory"),
        directoryName: z
          .string()
          .describe("The name of the directory directory"),
        name: z.string().describe("The name of the workbook file"),
      })
      .describe("Object containing file and directory information"),
  }),
  func: async ({ selectedFileData }, runManager) => {
    try {
      // Get the Teams context from the global scope
      const context = global.teamsContext;

      if (!context || typeof context.sendActivity !== "function") {
        console.error("No valid Teams context found:", context);
        throw new Error("Teams context not properly initialized");
      }

      const result = await processRfiWorksheet(context, selectedFileData);

      if (!result) {
        return "Processing completed. No RFI data found to process.";
      }

      return `Successfully created new RFI Response workbook: ${result}`;
    } catch (error) {
      console.error("Processing error:", error);
      return `Error processing RFI worksheet: ${error.message}`;
    }
  },
});

export const generateAuditorNotes = new DynamicStructuredTool({
  name: "generateAuditorNotes",
  description: "Process an RFI Response workbook and generate auditor notes",
  schema: z.object({
    selectedFileData: z
      .object({
        id: z.string().describe("The SharePoint ID of the workbook"),
        name: z.string().describe("The name of the workbook file"),
      })
      .describe("Object containing file information"),
  }),
  func: async ({ selectedFileData }, runManager) => {
    try {
      // Get the Teams context from the global scope
      const context = global.teamsContext;

      if (!context || typeof context.sendActivity !== "function") {
        console.error("No valid Teams context found:", context);
        throw new Error("Teams context not properly initialized");
      }

      console.log("Processing auditor notes for:", {
        filename: selectedFileData.name,
        workbookId: selectedFileData.id,
      });

      await processAuditorNotes(
        context,
        selectedFileData.name,
        selectedFileData.id
      );

      return `Successfully generated and added auditor notes to: ${selectedFileData.name}`;
    } catch (error) {
      console.error("Error generating auditor notes:", error);
      return `Error generating auditor notes: ${error.message}`;
    }
  },
});

export const sendToWorkflowAgentTool = new DynamicStructuredTool({
  name: "sendToWorkflow",
  description: "Send message and context details to the workflow agent",
  schema: z.object({
    message: z.string().describe("The message to send to the workflow agent"),
  }),
  func: async ({ message }) => {
    try {
      // Get the Teams context from the global scope
      const context = global.teamsContext;

      if (!context || !context.activity) {
        console.error("No valid Teams context found:", context);
        throw new Error("Teams context not properly initialized");
      }

      const contextDetails = {
        serviceUrl: context.activity.serviceUrl,
        conversationId: context.activity.conversation.id,
        channelId: context.activity.channelId,
        tenantId: context.activity.conversation.tenantId,
      };

      await sendToWorkflowAgent(contextDetails, message);
      return `Successfully sent message to workflow agent`;
    } catch (error) {
      console.error("Error sending to workflow agent:", error);
      return `Failed to send to workflow agent: ${error.message}`;
    }
  },
});
