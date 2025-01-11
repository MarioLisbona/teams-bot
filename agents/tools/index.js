import { DynamicStructuredTool } from "@langchain/core/tools";
import { z } from "zod";
import {
  listFoldersInDirectory,
  getFileNamesAndIds,
} from "../../lib/utils/fileStorageAndRetrieval.js";
import { processRfiWorksheet } from "../../lib/utils/auditProcessing.js";

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
        directoryName: z.string().describe("The name of the client directory"),
        name: z.string().describe("The name of the workbook file"),
      })
      .describe("Object containing file and directory information"),
  }),
  func: async ({ selectedFileData }, runManager) => {
    try {
      let context = runManager?.context;

      if (!context?.sendActivity) {
        console.warn(
          "No valid Teams context available, using fallback messaging"
        );
        context = {
          sendActivity: async (message) => {
            console.log("Teams Message (fallback):", message);
            await runManager?.handleText(message);
            return { id: "fallback-message-id" };
          },
        };
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

// Update the tools export
export const tools = [
  listFolders,
  listExcelFiles,
  processRfiWorksheet,
  // ... other tools
];
