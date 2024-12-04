import { getGraphClient, getAccessToken } from "./msAuth.js";
import { getFileIdByName } from "./oneDrive.js";
import axios from "axios";

// Function to check if a file with the same name exists
const checkFileExists = async (userId, folderId, fileName) => {
  const client = await getGraphClient();
  const response = await client
    .api(`/users/${userId}/drive/items/${folderId}/children`)
    .get();

  return response.value.some((file) => file.name === fileName);
};

// Function to generate a unique filename
const generateUniqueFilename = async (userId, folderId, baseName) => {
  let uniqueName = baseName;
  let counter = 1;

  // Check if the base name already exists
  while (await checkFileExists(userId, folderId, uniqueName)) {
    // Generate a new name with a counter
    uniqueName = `${baseName.replace(/(\.[\w\d]+)$/, ` (${counter})$1`)}`; // Append counter before the file extension
    counter++;
  }

  return uniqueName;
};

// Function to copy a file in OneDrive
export const copyFileInOneDrive = async (
  userId,
  fileId,
  baseFileName,
  targetDirectoryId
) => {
  const client = await getGraphClient();
  const folderId = targetDirectoryId || "root";

  console.log("Copying file to directory:", folderId);

  // Generate a unique filename
  const newWorkbookName = await generateUniqueFilename(
    userId,
    folderId,
    baseFileName
  );

  try {
    // Step 1: Copy the RFI Client Template workbook to the specified directory
    const copiedWorkbook = await client
      .api(`/users/${userId}/drive/items/${fileId}/copy`)
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
      .api(`/users/${userId}/drive/items/${folderId}/children`)
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
