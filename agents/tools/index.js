import { DynamicStructuredTool } from "@langchain/core/tools";
import { z } from "zod";
import {
  listFoldersInDirectory,
  getFileNamesAndIds,
} from "../../lib/utils/fileStorageAndRetrieval.js";

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
