import { CardFactory } from "botbuilder";
import { getFileNamesAndIds } from "./fileStorageAndRetrieval.js";
import {
  createFileSelectionCard,
  createTeamsUpdateCard,
} from "./adaptiveCards.js";
import { analyseAcpResponsePrompt } from "./prompts.js";
import { azureGptQuery } from "./azureGpt.cjs";
import { knowledgeBase } from "./acpResponsesKb.js";

/**
 * This function converts a column index to a letter (e.g., 0 -> A, 4 -> E, 26 -> AA).
 * @param {number} colIndex - The index of the column.
 * @returns {string} - The letter corresponding to the column index.
 */
function getColumnLetter(colIndex) {
  let letter = "";
  while (colIndex >= 0) {
    letter = String.fromCharCode((colIndex % 26) + 65) + letter;
    colIndex = Math.floor(colIndex / 26) - 1;
  }
  return letter;
}

/**
 * This function filters and selects the rows for RFI cells.
 * @param {Array} data - The data to filter.
 * @returns {Array} - The filtered data.
 */
export function filterRowsForRFICells(data) {
  // Filter cells to include those that contain the substring "RFI" from rows 3 onwards
  const filteredRows = data
    .slice(2) // Start from row 3 (index 2)
    .map((row, rowIndex) => {
      // Get iid from column C (index 2)
      const iid = row[2] || null;

      return row
        .map((cell, colIndex) => {
          // Skip column AJ (index 35) for the "RFI" check
          if (colIndex === 35) return null;

          // Check if cell is a string before calling includes
          if (typeof cell === "string" && cell.includes("RFI")) {
            // Determine the cell reference, e.g., E5
            const cellReference = `${getColumnLetter(colIndex)}${rowIndex + 3}`;
            return { rfi: cell, cellReference, iid };
          }
          return null;
        })
        .filter((cell) => cell !== null); // Keep only cells that contain "RFI"
    })
    .filter((row) => row.length > 0); // Remove empty rows after filtering

  return filteredRows;
}

/**
 * This function groups data into an object with the unique rfi, and projectsAffected array.
 * @param {Array} filteredRows - The filtered rows.
 * @returns {Array} - The grouped data.
 */
export function groupByRFI(filteredRows) {
  const groupedData = {}; // Initialize an empty object to store grouped data

  filteredRows.forEach((row) => {
    row.forEach(({ rfi, cellReference, iid }) => {
      if (!groupedData[rfi]) {
        groupedData[rfi] = {
          rfi,
          projectsAffected: [], // Initialize an array to store projects affected by this RFI
        };
      }

      // Add the cellReference and iid to the projectsAffected array for this rfi
      groupedData[rfi].projectsAffected.push({ cellReference, iid });
    });
  });

  // Convert the grouped data to an array of objects if needed
  return Object.values(groupedData);
}

/**
 * This function updates the RFI data by prepending the rfi data with "The auditor noted that"
 * then adding an action item based on the context.
 * @param {Array} filteredRows - The filtered rows.
 * @returns {Array} - The updated RFI rows.
 */
export async function updateRfiCellData(filteredRows) {
  // Groups all common rfi's into an object with the rfi, and projectsAffected array
  const groupedData = groupByRFI(filteredRows);

  // Extract all RFI attributes into an array
  const allRfiAttributes = groupedData.map((group) => group.rfi);

  // Transform RFI text - remove "RFI - " and add action item based on context
  const transformedRfiAttributes = allRfiAttributes.map((rfiText) => {
    // Remove "RFI - " and trim any extra spaces
    const cleanText = rfiText.replace(/^RFI\s*-\s*/i, "").trim();

    // Add prefix and determine appropriate action item based on context
    let actionItem = "Can you please clarify?"; // default action

    if (
      cleanText.toLowerCase().includes("not been uploaded") ||
      cleanText.toLowerCase().includes("not been provided")
    ) {
      actionItem = "Can you please provide this documentation?";
    } else if (
      cleanText.toLowerCase().includes("invoice") ||
      cleanText.toLowerCase().includes("declaration")
    ) {
      actionItem = "Can you please review and provide clarification?";
    }

    return `The auditor noted that ${cleanText}. ${actionItem}`;
  });

  // Add the transformedRfi attribute to each object in the groupedData array
  groupedData.forEach((group, index) => {
    group.updatedRfi = transformedRfiAttributes[index];
  });

  return groupedData;
}

/**
 * This function prepares data for updating Google Sheets.
 * @param {Array} data - The data to prepare.
 * @returns {Array} - The prepared data.
 */
export function prepareRfiCellDataForRfiSpreadsheet(data) {
  // Create an array of arrays with the rfi and projectsAffected values
  // To be used to update the RFI Spreadsheet
  return data.map((group) => {
    const row = [];

    // Place 'rfi' value in column A
    row.push(group.updatedRfi);

    // Ensure 'projectsAffected' is an array of the correct values (e.g., iid)
    const projectsAffected = group.projectsAffected
      .map((project) => project.iid)
      .join(", "); // Extract iid and join into a string
    row.push(projectsAffected);

    return row;
  });
}

/**
 * This function gets the ranges for general and specific issues RFI.
 * @param {number} startRowGeneral - The start row for general issues.
 * @param {number} startRowSpecific - The start row for specific issues.
 * @param {number} generalDataLength - The length of the general data.
 * @param {number} specificDataLength - The length of the specific data.
 * @returns {Object} - The ranges for general and specific issues RFI.
 */
export const getRfiRanges = (
  startRowGeneral,
  startRowSpecific,
  generalDataLength,
  specificDataLength
) => {
  const rangeForGeneralIssuesRfi = `C${startRowGeneral}:D${
    startRowGeneral + generalDataLength - 1
  }`;
  const rangeForSpecificIssuesRfi = `C${startRowSpecific}:D${
    startRowSpecific + specificDataLength - 1
  }`;

  return { rangeForGeneralIssuesRfi, rangeForSpecificIssuesRfi };
};

/**
 * This function updates the Excel spreadsheet for a given range and data.
 * @param {Object} client - The client object.
 * @param {string} url - The URL to update.
 * @param {Object} requestBody - The request body.
 * @param {string} issueType - The type of issue.
 */
export const updateExcelData = async (client, url, requestBody, issueType) => {
  try {
    // Send a PATCH request to update the data in the specified range
    await client.api(url).patch(requestBody);
    console.log(`Excel spreadsheet updated successfully for ${issueType}.`);
  } catch (error) {
    // Log the error if the update fails
    console.error(`Error updating Excel data for ${issueType}:`, error.message);
    console.error("Full error details for", issueType, ":", error);
    // Log the request details for better debugging
    console.error("Request URL for", issueType, ":", url);
    console.error(
      "Request Body for",
      issueType,
      ":",
      JSON.stringify(requestBody)
    );
  }
};

/**
 * This function calculates the cell range based on the data array with an offset.
 * @param {Array} data - The data to calculate the range for.
 * @param {string} startCell - The starting cell.
 * @param {boolean} preserveOriginalPositions - Whether to preserve the original positions.
 * @returns {string} - The cell range.
 */
export const getCellRange = (
  data,
  startCell = "A1",
  preserveOriginalPositions = false
) => {
  const numRows = data.length;
  const numCols = data[0] ? data[0].length : 0;

  if (!preserveOriginalPositions) {
    // For template copying, use the exact positions
    const endCell =
      String.fromCharCode(startCell.charCodeAt(0) + numCols - 1) +
      (parseInt(startCell.slice(1)) + numRows - 1);

    return `${startCell}:${endCell}`;
  }

  // For RFI data placement, use the offset positions
  const startColChar = "B"; // Always start from column B
  const startRow = 5; // Start from row 5 to shift data up by 9 rows
  const endCell =
    String.fromCharCode(startColChar.charCodeAt(0) + numCols - 1) +
    (startRow + numRows - 1);

  return `${startColChar}${startRow}:${endCell}`;
};

/**
 * This function handles the directory selection and creates a file selection card.
 * @param {Object} context - The context object.
 * @param {string} selectedDirectoryId - The ID of the selected directory.
 * @param {Object} options - The options object.
 * @returns {Promise<void>} - A promise that resolves when the directory selection is handled.
 */
export async function handleDirectorySelection(
  context,
  selectedDirectoryId,
  options = {}
) {
  try {
    const files = await getFileNamesAndIds(selectedDirectoryId);
    const selectedDirectory = JSON.parse(
      context.activity.value.directoryChoice
    );
    const directoryName = selectedDirectory.name;

    // Filter files if filterPattern is provided
    const filteredFiles = options.filterPattern
      ? files.filter((file) => file.name.includes(options.filterPattern))
      : files;

    const card = createFileSelectionCard(
      filteredFiles,
      selectedDirectoryId,
      directoryName,
      options.customSubheading,
      options.buttonText,
      options.action
    );
    await context.sendActivity({
      attachments: [CardFactory.adaptiveCard(card)],
    });
  } catch (error) {
    console.error("Error handling directory selection:", error);
    await context.sendActivity(
      "Error retrieving files from the selected directory."
    );
  }
}

/**
 * This function extracts the RFI response data from the RFI spreadsheet,
 * used to update the RFI spreadsheet with the generated auditorNotes.
 * @param {Array} data - The data to extract.
 * @param {Array} ranges - The ranges to extract.
 * @returns {Array} - The extracted data.
 */
export const extractRfiResponseData = (data, ranges) => {
  return ranges.flatMap(
    ([start, end]) =>
      data
        .slice(start - 1, end)
        .filter((row) => row[0] || row[2]) // Filter rows with data in columns C or E
        .map((row) => ({
          rfiNumber: row[0] || "", // Column A
          issuesIdentified: row[1] || "", // Column C
          acpResponse: row[3] || "", // Column E
        }))
        .filter((obj) => obj.issuesIdentified && obj.acpResponse) // Only keep objects where both values are non-empty
  );
};

/**
 * This function prepares data for batch update.
 * @param {Array} updatedResponseData - The updated response data.
 * @returns {Object} - The prepared data.
 */
export const prepareDataForBatchUpdate = (updatedResponseData) => {
  // Split the data into general and specific arrays
  const generalIssues = updatedResponseData
    .filter((response) => response.rfiNumber.startsWith("G."))
    .sort(
      (a, b) => parseInt(a.rfiNumber.slice(2)) - parseInt(b.rfiNumber.slice(2))
    );

  const specificIssues = updatedResponseData
    .filter((response) => response.rfiNumber.startsWith("S."))
    .sort(
      (a, b) => parseInt(a.rfiNumber.slice(2)) - parseInt(b.rfiNumber.slice(2))
    );

  // Create arrays for each section with the correct size
  const generalArray = Array(21).fill([""]); // F14 to F34 (21 rows)
  const specificArray = Array(100).fill([""]); // F42 to F141 (100 rows)

  // Fill in the general issues
  generalIssues.forEach((item) => {
    const index = parseInt(item.rfiNumber.slice(2)) - 1; // G.1 goes to index 0
    if (index >= 0 && index < 21) {
      generalArray[index] = [item.auditorNotes || ""];
    }
  });

  // Fill in the specific issues
  specificIssues.forEach((item) => {
    const index = parseInt(item.rfiNumber.slice(2)) - 1; // S.1 goes to index 0
    if (index >= 0 && index < 100) {
      specificArray[index] = [item.auditorNotes || ""];
    }
  });

  return {
    generalArray, // For writing to rows F14:F34 in RFI Spreadsheet
    specificArray, // For writing to rows F42:F141 in RFI Spreadsheet
  };
};

/**
 * This function creates a Teams update.
 * @param {Object} context - The context object.
 * @param {string} text - The text to update.
 * @returns {Promise<void>} - A promise that resolves when the update is created.
 */
export const createTeamsUpdate = async (
  context,
  text,
  userMessage,
  emoji,
  style
) => {
  const teamsUpdateCard = createTeamsUpdateCard(
    text,
    userMessage,
    emoji,
    style
  );

  await context.sendActivity({
    type: "message",
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: teamsUpdateCard,
      },
    ],
  });
};

/**
 * This function processes client responses in batches.
 * This is used to generate the auditorNotes for the RFI spreadsheet
 * while avoiding the rate limit of the Azure GPT API.
 * @param {Object} context - The context object.
 * @param {Array} processedClientResponses - The processed client responses.
 * @returns {Array} - The all responses.
 */
export const batchProcessClientResponses = async (
  context,
  processedClientResponses
) => {
  const batchSize = 6;
  const allResponses = [];

  for (let i = 0; i < processedClientResponses.length; i += batchSize) {
    const clientResponsesBatch = processedClientResponses.slice(
      i,
      i + batchSize
    );

    const prompt = analyseAcpResponsePrompt(
      knowledgeBase,
      clientResponsesBatch
    );

    // Update progress
    await createTeamsUpdate(
      context,
      `Processing client responses batch ${
        Math.floor(i / batchSize) + 1
      }/${Math.ceil(processedClientResponses.length / batchSize)}...`,
      "",
      "💭",
      "default"
    );

    // Generate the response from Azure GPT for this batch
    const azureResponse = await azureGptQuery(prompt);
    const batchResults = JSON.parse(azureResponse);
    allResponses.push(...batchResults);
  }

  return allResponses;
};

// Helper function to split array into chunks
export function chunk(array, size) {
  const chunked = [];
  for (let i = 0; i < array.length; i += size) {
    chunked.push(array.slice(i, i + size));
  }
  return chunked;
}
