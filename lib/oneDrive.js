import axios from "axios";
import { getAccessToken } from "./msAuth.js";

async function listFilesInOneDrive(folderId = null) {
  try {
    const accessToken = await getAccessToken();
    const url = folderId
      ? `https://graph.microsoft.com/v1.0/drives/${process.env.ONEDRIVE_ID}/items/${folderId}/children`
      : `https://graph.microsoft.com/v1.0/drives/${process.env.ONEDRIVE_ID}/root/children`;

    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    return response.data.value.filter(
      (file) => file.name.endsWith(".xlsx") || file.name.endsWith(".xls")
    );
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

export async function getFileNamesAndIds(folderId = null) {
  const files = await listFilesInOneDrive(folderId);
  return files.map((file) => ({
    name: file.name,
    id: file.id,
  }));
}

export async function getFileIdByName(fileName) {
  const files = await getFileNamesAndIds();
  const file = files.find((file) => file.name === fileName);
  return file ? file.id : null;
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
