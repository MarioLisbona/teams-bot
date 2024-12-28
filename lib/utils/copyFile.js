import { getGraphClient } from "../auth/msAuth.js";

/**
 * This function checks if a file with the same name exists in a specified directory.
 * @param {string} folderId - The ID of the directory.
 * @param {string} fileName - The name of the file.
 * @returns {boolean} - True if the file exists, false otherwise.
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
 * Generates a unique filename by appending a counter if the base name already exists.
 * @param {string} folderId - The ID of the directory to check for existing files.
 * @param {string} baseName - The original filename to base the unique name on.
 * @returns {Promise<string>} - A promise that resolves to a unique filename.
 * @throws {Error} If file checking fails or maximum attempts are reached.
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
 * This function copies a file in OneDrive.
 * @param {string} fileId - The ID of the file to copy.
 * @param {string} baseFileName - The name of the file to copy.
 * @param {string} targetDirectoryId - The ID of the directory to copy the file to.
 * @returns {Object} - The ID and name of the new file.
 * @throws {Error} If file copying fails or the new file cannot be found.
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
