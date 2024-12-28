import { getGraphClient } from "../auth/msAuth.js";

/**
 * Checks for file existence in a SharePoint directory using Microsoft Graph API.
 *
 * @description
 * Performs file existence check by:
 * 1. Initializing Microsoft Graph client
 * 2. Retrieving directory contents
 * 3. Comparing file names for exact match
 * 4. Providing detailed error handling
 *
 * The function uses the SharePoint site ID from environment variables
 * and performs a case-sensitive filename comparison.
 *
 * @param {string} folderId - SharePoint directory/folder identifier
 * @param {string} fileName - Exact name of file to check for
 *
 * @throws {Error} When Graph client initialization fails
 * @throws {Error} When directory contents cannot be retrieved
 * @returns {Promise<boolean>} True if file exists, false otherwise
 *
 * @example
 * try {
 *   const exists = await checkFileExists(
 *     "folder123",
 *     "document.xlsx"
 *   );
 *   console.log(exists ? "File exists" : "File not found");
 * } catch (error) {
 *   console.error("File check failed:", error);
 * }
 *
 * @requires Microsoft Graph API
 * @requires SHAREPOINT_SITE_ID environment variable
 */
const checkFileExists = async (folderId, fileName) => {
  try {
    const client = await getGraphClient();

    try {
      const response = await client
        .api(
          `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${folderId}/children`
        )
        .get();

      return response.value.some((file) => file.name === fileName);
    } catch (error) {
      console.error("Failed to check directory contents:", error);
      throw new Error(
        `Failed to check directory for file "${fileName}": ${error.message}`
      );
    }
  } catch (error) {
    console.error("Graph client initialization failed:", error);
    throw new Error(
      `Failed to initialize client for file check: ${error.message}`
    );
  }
};

/**
 * Generates a unique filename by appending incremental counters.
 *
 * @description
 * Creates unique filenames by:
 * 1. Checking if base filename exists
 * 2. Appending counter before file extension if needed
 * 3. Incrementing counter until unique name is found
 * 4. Limiting attempts to prevent infinite loops
 *
 * Format: "filename (1).ext", "filename (2).ext", etc.
 *
 * @param {string} folderId - SharePoint directory/folder identifier
 * @param {string} baseName - Original filename to make unique
 *
 * @throws {Error} When file existence check fails
 * @throws {Error} When maximum attempts (100) are reached
 * @returns {Promise<string>} Unique filename guaranteed not to exist
 *
 * @example
 * const uniqueName = await generateUniqueFilename(
 *   "folder123",
 *   "report.xlsx"
 * );
 * // Returns "report.xlsx" or "report (1).xlsx" if original exists
 */
const generateUniqueFilename = async (folderId, baseName) => {
  try {
    let uniqueName = baseName;
    let counter = 1;
    let maxAttempts = 100; // Prevent infinite loops
    let attempts = 0;

    while (attempts < maxAttempts) {
      try {
        const exists = await checkFileExists(folderId, uniqueName);
        if (!exists) {
          return uniqueName;
        }
        // Generate a new name with a counter
        uniqueName = `${baseName.replace(/(\.[\w\d]+)$/, ` (${counter})$1`)}`; // Append counter before the file extension
        counter++;
        attempts++;
      } catch (error) {
        console.error("Failed to check filename:", error);
        throw new Error(
          `Failed to generate unique filename for "${baseName}": ${error.message}`
        );
      }
    }
    throw new Error(
      `Maximum filename generation attempts (${maxAttempts}) reached for "${baseName}"`
    );
  } catch (error) {
    console.error("Filename generation failed:", error);
    throw new Error(`Unable to generate unique filename: ${error.message}`);
  }
};

/**
 * Copies a file to a target directory in SharePoint/OneDrive.
 *
 * @description
 * Handles file copying workflow:
 * 1. Generates unique destination filename
 * 2. Initiates asynchronous copy operation
 * 3. Waits for copy completion
 * 4. Verifies and retrieves new file details
 * 5. Provides detailed error handling
 *
 * Uses Microsoft Graph API for all operations and includes
 * a 2-second delay to ensure copy completion.
 *
 * @param {string} fileId - Source file identifier
 * @param {string} baseFileName - Desired name for copied file
 * @param {string} targetDirectoryId - Destination directory ID (defaults to "root")
 *
 * @throws {Error} When filename generation fails
 * @throws {Error} When copy operation fails
 * @throws {Error} When new file verification fails
 * @returns {Promise<Object>} Object containing:
 *   - newWorkbookId: ID of copied file
 *   - newWorkbookName: Final name of copied file
 *
 * @example
 * try {
 *   const { newWorkbookId, newWorkbookName } = await copyFileInOneDrive(
 *     "file123",
 *     "document.xlsx",
 *     "folder456"
 *   );
 *   console.log(`File copied as: ${newWorkbookName}`);
 * } catch (error) {
 *   console.error("Copy failed:", error);
 * }
 *
 * @requires Microsoft Graph API
 * @requires SHAREPOINT_SITE_ID environment variable
 */
export const copyFileInOneDrive = async (
  fileId,
  baseFileName,
  targetDirectoryId
) => {
  try {
    const client = await getGraphClient();
    const folderId = targetDirectoryId || "root";

    console.log("Copying file to directory:", folderId);

    try {
      // Generate a unique filename
      const newWorkbookName = await generateUniqueFilename(
        folderId,
        baseFileName
      );

      try {
        // Step 1: Copy the RFI Client Template workbook to the specified directory
        const copiedWorkbook = await client
          .api(
            `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${fileId}/copy`
          )
          .post({
            parentReference: {
              id: folderId,
            },
            name: newWorkbookName,
          });

        // The copy operation is asynchronous, wait a moment for it to complete
        await new Promise((resolve) => setTimeout(resolve, 2000));

        try {
          // Step 2: Get the ID of the newly created file from the specified directory
          const files = await client
            .api(
              `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${folderId}/children`
            )
            .filter(`name eq '${newWorkbookName}'`)
            .get();

          if (!files.value || files.value.length === 0) {
            throw new Error(
              `Newly created file "${newWorkbookName}" not found`
            );
          }

          const newWorkbookId = files.value[0].id;
          console.log(`File copied successfully. New ID: ${newWorkbookId}`);

          return { newWorkbookId, newWorkbookName };
        } catch (error) {
          console.error("Failed to verify copied file:", error);
          throw new Error(
            `Failed to locate newly copied file: ${error.message}`
          );
        }
      } catch (error) {
        console.error("Failed to copy file:", error);
        throw new Error(
          `Failed to copy file "${baseFileName}": ${error.message}`
        );
      }
    } catch (error) {
      console.error("Failed to generate unique filename:", error);
      throw new Error(`Failed to generate filename for copy: ${error.message}`);
    }
  } catch (error) {
    console.error("Error in copyFileInOneDrive:", error);
    throw new Error(`File copy operation failed: ${error.message}`);
  }
};
