import axios from "axios";
import { getAccessToken } from "./msAuth.js";

async function listFilesInOneDrive(driveId) {
  try {
    const accessToken = await getAccessToken();
    const url = `https://graph.microsoft.com/v1.0/drives/${process.env.ONEDRIVE_ID}/root/children`;

    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });
    return response.data.value;
  } catch (error) {
    console.error("Error listing files:", {
      status: error.response?.status,
      statusText: error.response?.statusText,
      data: error.response?.data,
      message: error.message,
      url: error.config?.url,
    });
    throw error;
  }
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

export async function getDirectories(driveId) {
  try {
    const accessToken = await getAccessToken();
    const url = `https://graph.microsoft.com/v1.0/drives/${process.env.ONEDRIVE_ID}/root/children`;

    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    const folders = response.data.value.filter((item) => item.folder);
    return folders.map((item) => ({
      id: item.id,
      name: item.name,
    }));
  } catch (error) {
    console.error("Error getting directories:", {
      status: error.response?.status,
      statusText: error.response?.statusText,
      data: error.response?.data,
      message: error.message,
      url: error.config?.url,
    });
    throw error;
  }
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
