import { getGraphClient, getAccessToken } from "./msAuth.js";
import { getFileIdByName } from "./oneDrive.js";

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
export const copyFileInOneDrive = async (userId, fileId, baseFileName) => {
  const client = await getGraphClient();
  const folderId = "root";

  // Generate a unique filename
  const newWorkbookName = await generateUniqueFilename(
    userId,
    folderId,
    baseFileName
  );

  // Step 1: Copy the RFI Client Template workbook
  const copiedWorkbook = await client
    .api(`/users/${userId}/drive/items/${fileId}/copy`)
    .post({
      parentReference: {
        id: folderId,
      },
      name: newWorkbookName,
    });

  // Get the ID of the newly copied workbook
  const newWorkbookId = await getFileIdByName(
    process.env.ONEDRIVE_ID,
    newWorkbookName
  );

  // Step 2: Rename the worksheet to "RFI Responses"
  await client
    .api(
      `/users/${userId}/drive/items/${newWorkbookId}/workbook/worksheets/RFI Spreadsheet Template`
    )
    .patch({
      name: "RFI Responses",
    });

  console.log(`File copied to new file: ${newWorkbookName}`);

  return { newWorkbookId, newWorkbookName };
};
