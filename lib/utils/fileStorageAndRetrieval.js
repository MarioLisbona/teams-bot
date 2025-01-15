import axios from "axios";
import { getAccessToken } from "../auth/msAuth.js";
import { loadEnvironmentVariables } from "../environment/setupEnvironment.js";

loadEnvironmentVariables();
/**
 * Retrieves simplified Excel file information from a SharePoint directory.
 *
 * @description
 * Gets basic file information by:
 * 1. Retrieving complete file list from directory
 * 2. Extracting essential properties (name and ID)
 * 3. Providing simplified data structure
 *
 * @param {string} [folderId=null] - SharePoint folder ID (null for root folder)
 *
 * @throws {Error} When file listing fails
 * @throws {Error} When data mapping fails
 * @returns {Promise<Array<Object>>} Array of simplified file objects:
 *   - name: Excel file name
 *   - id: SharePoint file identifier
 *
 * @example
 * const files = await getFileNamesAndIds("folder123");
 * // Returns: [{ name: "report.xlsx", id: "file123" }, ...]
 */
async function listExcelFilesInDirectory(folderId = null) {
  try {
    // Get access token
    let accessToken;
    try {
      accessToken = await getAccessToken();
      if (!accessToken) {
        throw new Error("Failed to obtain access token");
      }
    } catch (error) {
      console.error("Access token retrieval failed:", error);
      throw new Error(`Authentication failed: ${error.message}`);
    }

    // Construct URL
    const url = folderId
      ? `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${folderId}/children`
      : `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/root/children`;

    try {
      // Make API request
      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      });

      // Filter Excel files
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
      throw new Error(`Failed to list Excel files: ${error.message}`);
    }
  } catch (error) {
    console.error("Directory listing failed:", error);
    throw new Error(`Failed to list directory contents: ${error.message}`);
  }
}

/**
 * This function gets the names and IDs of all Excel files in a specified directory.
 * @param {string} folderId - The ID of the directory.
 * @returns {Array} - The list of Excel files with their names and IDs.
 * @throws {Error} If file listing fails or data mapping fails.
 */
export async function getFileNamesAndIds(folderId = null) {
  try {
    const files = await listExcelFilesInDirectory(folderId);

    try {
      return files.map((file) => ({
        name: file.name,
        id: file.id,
      }));
    } catch (error) {
      console.error("Failed to map file data:", error);
      throw new Error(`Failed to process file information: ${error.message}`);
    }
  } catch (error) {
    console.error("Failed to get file names and IDs:", error);
    throw new Error(`Failed to retrieve file list: ${error.message}`);
  }
}

/**
 * Retrieves client directories from a specified SharePoint folder.
 *
 * @description
 * Directory retrieval process:
 * 1. Authenticates with Microsoft Graph
 * 2. Locates target directory in root
 * 3. Retrieves all subdirectories
 * 4. Filters for folder-type items
 * 5. Extracts essential directory information
 *
 * @param {string} directoryName - Name of target directory in root
 *
 * @throws {Error} When authentication fails
 * @throws {Error} When root directory access fails
 * @throws {Error} When target directory not found
 * @throws {Error} When subdirectory listing fails
 * @returns {Promise<Array<Object>>} Array of directory objects:
 *   - id: SharePoint folder identifier
 *   - name: Folder display name
 *
 * @example
 * const dirs = await getClientDirectories("Clients");
 * // Returns: [{ name: "Client A", id: "dir123" }, ...]
 *
 * @requires axios
 * @requires SHAREPOINT_SITE_ID environment variable
 */
export async function getClientDirectories(directoryName) {
  try {
    // Get the access token
    let accessToken;
    try {
      accessToken = await getAccessToken();
      if (!accessToken) {
        throw new Error("Failed to obtain access token");
      }
    } catch (error) {
      console.error("Access token retrieval failed:", error);
      throw new Error(`Authentication failed: ${error.message}`);
    }

    try {
      // Get the root directory contents using SharePoint site ID
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
        throw new Error(
          `Folder "${directoryName}" not found in root directory`
        );
      }

      try {
        // Get the contents of the target folder
        const folderContents = await axios.get(
          `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${targetFolder.id}/children`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              "Content-Type": "application/json",
            },
          }
        );

        try {
          // Filter and map the folder contents
          const folders = folderContents.data.value.filter(
            (item) => item.folder
          );
          return folders.map((item) => ({
            id: item.id,
            name: item.name,
          }));
        } catch (error) {
          console.error("Failed to process folder contents:", error);
          throw new Error(
            `Failed to process directory information: ${error.message}`
          );
        }
      } catch (error) {
        console.error("Failed to get folder contents:", error);
        throw new Error(
          `Failed to retrieve contents of "${directoryName}": ${error.message}`
        );
      }
    } catch (error) {
      console.error("Failed to access root directory:", error);
      throw new Error(
        `Failed to access SharePoint root directory: ${error.message}`
      );
    }
  } catch (error) {
    console.error("Error getting directories:", {
      status: error.response?.status,
      statusText: error.response?.statusText,
      data: error.response?.data,
      message: error.message,
      url: error.config?.url,
    });
    throw new Error(`Directory listing failed: ${error.message}`);
  }
}

export async function listFoldersInDirectory() {
  const folderId = "017CTZANUXE4AIIFUFLJB2QUYGMAOJQ7OM";
  console.log("folderId", folderId);
  try {
    const accessToken = await getAccessToken();
    if (!accessToken) {
      throw new Error("Failed to obtain access token");
    }

    const url = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${folderId}/children`;

    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    // Filter folders and map to only id and name
    const folders = response.data.value
      .filter((item) => item.folder !== undefined)
      .map((folder) => ({
        id: folder.id,
        name: folder.name,
      }));

    // Log the simplified folder list
    console.log("Found folders:", folders);

    return folders;
  } catch (error) {
    console.error("Directory listing failed:", error);
    throw new Error(`Failed to list directory contents: ${error.message}`);
  }
}
