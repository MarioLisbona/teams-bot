import axios from "axios";
import { getAccessToken } from "../auth/msAuth.js";

/**
 * Lists Excel files in a SharePoint directory using Microsoft Graph API.
 *
 * @description
 * Retrieves Excel files by:
 * 1. Obtaining Microsoft Graph access token
 * 2. Constructing appropriate API endpoint URL:
 *    - For specific folder: /drive/items/{folderId}/children
 *    - For root folder: /drive/root/children
 * 3. Making authenticated API request
 * 4. Filtering for Excel files (.xlsx, .xls)
 *
 * Includes comprehensive error logging for:
 * - Authentication failures
 * - API request failures
 * - HTTP response details
 *
 * @param {string} [folderId=null] - SharePoint folder ID (null for root folder)
 *
 * @throws {Error} When access token retrieval fails
 * @throws {Error} When API request fails
 * @throws {Error} When folder contents cannot be retrieved
 * @returns {Promise<Array<Object>>} Array of Excel file objects containing:
 *   - name: Filename with extension
 *   - id: SharePoint file identifier
 *   - Additional file metadata from Graph API
 *
 * @requires axios
 * @requires SHAREPOINT_SITE_ID environment variable
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
 * This function gets the names and IDs of all directories in a specified directory.
 * @param {string} directoryName - The name of the directory to search in.
 * @returns {Promise<Array>} - Array of objects containing directory ids and names.
 * @throws {Error} If access token retrieval fails.
 * @throws {Error} If root directory listing fails.
 * @throws {Error} If target folder is not found.
 * @throws {Error} If folder contents cannot be retrieved.
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

/**
 * This function gets the names and IDs of all directories in the "Evidence Packs" folder.
 * @param {string} directoryName - The name of the parent directory.
 * @param {string} subDirectoryID - The ID of the subdirectory to search in.
 * @returns {Promise<Array>} - Array of objects containing directory ids and names.
 * @throws {Error} If access token retrieval fails.
 * @throws {Error} If subdirectory contents cannot be retrieved.
 * @throws {Error} If Evidence Packs folder is not found.
 * @throws {Error} If folder contents cannot be processed.
 */
export async function getClientJobDirectories(directoryName, subDirectoryID) {
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

    try {
      // Get the subdirectory contents
      const rootResponse = await axios.get(
        `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${subDirectoryID}/children`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );

      // Find the "Evidence Packs" folder
      const evidencePacksFolder = rootResponse.data.value.find(
        (item) => item.name === "Evidence Packs" && item.folder
      );
      if (!evidencePacksFolder) {
        throw new Error(
          `"Evidence Packs" folder not found in "${directoryName}"`
        );
      }

      try {
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

        try {
          // Filter and map the folders
          const folders = evidencePacksContents.data.value.filter(
            (item) => item.folder
          );
          return folders.map((item) => ({
            id: item.id,
            name: item.name,
          }));
        } catch (error) {
          console.error("Failed to process folder contents:", error);
          throw new Error(
            `Failed to process Evidence Packs contents: ${error.message}`
          );
        }
      } catch (error) {
        console.error("Failed to get Evidence Packs contents:", error);
        throw new Error(
          `Failed to retrieve Evidence Packs contents: ${error.message}`
        );
      }
    } catch (error) {
      console.error("Failed to access subdirectory:", error);
      throw new Error(
        `Failed to access directory "${directoryName}": ${error.message}`
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
    throw new Error(`Failed to get job directories: ${error.message}`);
  }
}
