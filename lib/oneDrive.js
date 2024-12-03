import axios from "axios";
import { getAccessToken } from "./msAuth.js";

async function listFilesInOneDrive(driveId) {
  const accessToken = await getAccessToken();
  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );
  return response.data.value; // Returns the list of files
}

export async function getFileNamesAndIds(driveId) {
  const files = await listFilesInOneDrive(driveId);
  return files.map((file) => ({
    id: file.id,
    name: file.name,
  }));
}

export async function getFileIdByName(driveId, fileName) {
  const files = await getFileNamesAndIds(driveId);
  const file = files.find((file) => file.name === fileName);
  return file ? file.id : null; // Returns the file ID or null if not found
}

// const files = await listFilesInOneDrive(process.env.ONEDRIVE_ID);
// console.log(files);

// const response = await getFileNamesAndIds(process.env.ONEDRIVE_ID);
// console.log(response);

// const fileId = await getFileIdByName(
//   process.env.ONEDRIVE_ID,
//   "els-testing-client-XXXX.xlsx"
// );
// console.log(fileId);
