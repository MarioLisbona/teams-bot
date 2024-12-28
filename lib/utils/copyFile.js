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

// Function to generate a unique filename
const generateUniqueFilename = async (folderId, baseName) => {
  let uniqueName = baseName;
  let counter = 1;

  // Check if the base name already exists
  while (await checkFileExists(folderId, uniqueName)) {
    // Generate a new name with a counter
    uniqueName = `${baseName.replace(/(\.[\w\d]+)$/, ` (${counter})$1`)}`; // Append counter before the file extension
    counter++;
  }

  return uniqueName;
};

/**
 * This function copies a file in OneDrive.
 * @param {string} fileId - The ID of the file to copy.
 * @param {string} baseFileName - The name of the file to copy.
 * @param {string} targetDirectoryId - The ID of the directory to copy the file to.
 * @returns {Object} - The ID and name of the new file.
 */
export const copyFileInOneDrive = async (
  fileId,
  baseFileName,
  targetDirectoryId
) => {
  const client = await getGraphClient();
  const folderId = targetDirectoryId || "root";

  console.log("Copying file to directory:", folderId);

  // Generate a unique filename
  const newWorkbookName = await generateUniqueFilename(folderId, baseFileName);

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

    // Step 2: Get the ID of the newly created file from the specified directory
    const files = await client
      .api(
        `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${folderId}/children`
      )
      .filter(`name eq '${newWorkbookName}'`)
      .get();

    if (!files.value || files.value.length === 0) {
      throw new Error("Newly created file not found");
    }

    const newWorkbookId = files.value[0].id;
    console.log(`File copied successfully. New ID: ${newWorkbookId}`);

    return { newWorkbookId, newWorkbookName };
  } catch (error) {
    console.error("Error in copyFileInOneDrive:", error);
    throw error;
  }
};
