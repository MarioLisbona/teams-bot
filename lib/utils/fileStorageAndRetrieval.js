import axios from "axios";
import { getAccessToken } from "../auth/msAuth.js";

async function listExcelFilesInDirectory(folderId = null) {
  try {
    const accessToken = await getAccessToken();
    const url = folderId
      ? `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${folderId}/children`
      : `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/root/children`;

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
  const files = await listExcelFilesInDirectory(folderId);
  return files.map((file) => ({
    name: file.name,
    id: file.id,
  }));
}

export async function getClientDirectories(directoryName) {
  try {
    // Get the access token
    const accessToken = await getAccessToken();

    // First, get the root directory contents using SharePoint site ID
    const rootResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/root/children`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      }
    );

    // Find the target folder in the root directory
    const targetFolder = rootResponse.data.value.find(
      (item) => item.name === directoryName && item.folder
    );

    if (!targetFolder) {
      throw new Error(`Folder "${directoryName}" not found`);
    }

    // Then get the contents of that folder using the SharePoint drive
    const folderContents = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${targetFolder.id}/children`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      }
    );

    // Filter the folder contents to only include folders
    const folders = folderContents.data.value.filter((item) => item.folder);

    // Return the folders with their id and name
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

export async function getClientJobDirectories(directoryName, subDirectoryID) {
  console.log(directoryName, subDirectoryID);
  try {
    const accessToken = await getAccessToken();
    // First, get the root directory contents
    const rootResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${subDirectoryID}/children`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      }
    );
    console.log({ rootResponse });
    // Find the "Evidence Packs" folder
    const evidencePacksFolder = rootResponse.data.value.find(
      (item) => item.name === "Evidence Packs" && item.folder
    );
    if (!evidencePacksFolder) {
      throw new Error(
        `"Evidence Packs" folder not found in "${directoryName}"`
      );
    }
    // Get the contents of the Evidence Packs folder
    const evidencePacksContents = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${evidencePacksFolder.id}/children`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      }
    );
    // Filter and return only folders
    const folders = evidencePacksContents.data.value.filter(
      (item) => item.folder
    );
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
