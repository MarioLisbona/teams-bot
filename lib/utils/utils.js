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
 * Converts a zero-based column index to an Excel-style column letter.
 * @param {number} colIndex - The zero-based index of the column (e.g., 0 for 'A', 25 for 'Z', 26 for 'AA').
 * @returns {string} The Excel-style column letter(s) (e.g., 'A', 'B', 'AA').
 * @throws {Error} If the input is not a valid number or is negative.
 * @example
 * getColumnLetter(0)  // returns 'A'
 * getColumnLetter(25) // returns 'Z'
 * getColumnLetter(26) // returns 'AA'
 */
function getColumnLetter(colIndex) {
  try {
    // Validate input
    if (typeof colIndex !== "number") {
      throw new Error("Column index must be a number");
    }

    if (colIndex < 0) {
      throw new Error("Column index cannot be negative");
    }

    let letter = "";
    try {
      while (colIndex >= 0) {
        letter = String.fromCharCode((colIndex % 26) + 65) + letter;
        colIndex = Math.floor(colIndex / 26) - 1;
      }

      if (!letter) {
        throw new Error("Failed to generate column letter");
      }

      return letter;
    } catch (error) {
      console.error("Column letter generation failed:", error);
      throw new Error(
        `Failed to convert index ${colIndex} to letter: ${error.message}`
      );
    }
  } catch (error) {
    console.error("Column index processing failed:", error);
    throw new Error(`Invalid column index: ${error.message}`);
  }
}

/**
 * Filters and processes Excel data to extract RFI information from specific cells.
 * @param {Array<Array<any>>} data - The raw Excel data as a 2D array.
 * @param {number} data[][].length - Each row should contain cell values.
 * @returns {Array<Array<Object>>} Filtered array of RFI objects with cell references.
 * @throws {Error} If data is invalid or processing fails.
 */
export function filterRowsForRFICells(data) {
  try {
    // Validate input
    if (!Array.isArray(data)) {
      throw new Error("Input must be an array");
    }

    try {
      // Filter cells to include those that contain the substring "RFI" from rows 3 onwards
      const filteredRows = data
        .slice(2) // Start from row 3 (index 2)
        .map((row, rowIndex) => {
          if (!Array.isArray(row)) {
            throw new Error(
              `Invalid row at index ${rowIndex + 2}: must be an array`
            );
          }

          try {
            // Get iid from column C (index 2)
            const iid = row[2] || null;

            return row
              .map((cell, colIndex) => {
                try {
                  // Skip column AJ (index 35) for the "RFI" check
                  if (colIndex === 35) return null;

                  // Check if cell is a string before calling includes
                  if (typeof cell === "string" && cell.includes("RFI")) {
                    // Determine the cell reference, e.g., E5
                    const cellReference = `${getColumnLetter(colIndex)}${
                      rowIndex + 3
                    }`;
                    return { rfi: cell, cellReference, iid };
                  }
                  return null;
                } catch (error) {
                  console.error(
                    `Error processing cell at row ${
                      rowIndex + 3
                    }, column ${colIndex}:`,
                    error
                  );
                  return null; // Skip problematic cells instead of failing entirely
                }
              })
              .filter((cell) => cell !== null); // Keep only cells that contain "RFI"
          } catch (error) {
            console.error(`Error processing row ${rowIndex + 3}:`, error);
            return []; // Return empty row for problematic rows
          }
        })
        .filter((row) => row.length > 0); // Remove empty rows after filtering

      return filteredRows;
    } catch (error) {
      console.error("Failed to process rows:", error);
      throw new Error(`Row processing failed: ${error.message}`);
    }
  } catch (error) {
    console.error("RFI cell filtering failed:", error);
    throw new Error(`Failed to filter RFI cells: ${error.message}`);
  }
}

/**
 * Groups RFI data by unique RFI identifiers and collects associated project information.
 * @param {Array<Array<Object>>} filteredRows - Nested array of RFI objects.
 * @param {string} filteredRows[][].rfi - The RFI identifier.
 * @param {string} filteredRows[][].cellReference - The Excel cell reference.
 * @param {string} filteredRows[][].iid - The project identifier.
 * @returns {Array<Object>} Array of grouped RFI data with project associations.
 * @throws {Error} If input is invalid or processing fails.
 */
export function groupByRFI(filteredRows) {
  try {
    // Validate input
    if (!Array.isArray(filteredRows)) {
      throw new Error("Input must be an array");
    }

    const groupedData = {}; // Initialize an empty object to store grouped data

    try {
      filteredRows.forEach((row, rowIndex) => {
        if (!Array.isArray(row)) {
          throw new Error(`Invalid row at index ${rowIndex}: must be an array`);
        }

        row.forEach(({ rfi, cellReference, iid }, colIndex) => {
          // Validate required properties
          if (!rfi) {
            throw new Error(
              `Missing RFI at row ${rowIndex}, column ${colIndex}`
            );
          }

          if (!groupedData[rfi]) {
            groupedData[rfi] = {
              rfi,
              projectsAffected: [], // Initialize array for projects affected by this RFI
            };
          }

          // Add the cellReference and iid to the projectsAffected array
          groupedData[rfi].projectsAffected.push({
            cellReference: cellReference || "",
            iid: iid || "",
          });
        });
      });

      // Convert the grouped data to an array of objects
      return Object.values(groupedData);
    } catch (error) {
      console.error("Failed to process RFI data:", error);
      throw new Error(`RFI data processing failed: ${error.message}`);
    }
  } catch (error) {
    console.error("RFI grouping failed:", error);
    throw new Error(`Failed to group RFI data: ${error.message}`);
  }
}

/**
 * Updates RFI data by prepending text and adding contextual action items.
 * @param {Array} filteredRows - Array of RFI rows to process.
 * @returns {Array} Array of processed RFI data with updated text.
 * @throws {Error} If RFI processing fails.
 */
export function updateRfiCellData(filteredRows) {
  try {
    // Validate input
    if (!Array.isArray(filteredRows)) {
      throw new Error("Input must be an array");
    }

    // Groups all common rfi's into an object
    const groupedData = groupByRFI(filteredRows);

    // Extract all RFI attributes
    const allRfiAttributes = groupedData.map((group) => {
      if (!group?.rfi) {
        throw new Error("Invalid RFI group structure");
      }
      return group.rfi;
    });

    // Transform RFI text
    const transformedRfiAttributes = allRfiAttributes.map((rfiText) => {
      if (typeof rfiText !== "string") {
        throw new Error("RFI text must be a string");
      }

      const cleanText = rfiText.replace(/^RFI\s*-\s*/i, "").trim();
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

    // Update the grouped data
    groupedData.forEach((group, index) => {
      group.updatedRfi = transformedRfiAttributes[index];
    });

    return groupedData;
  } catch (error) {
    console.error("RFI cell data update failed:", error);
    throw new Error(`Failed to update RFI data: ${error.message}`);
  }
}

/**
 * Prepares RFI data for Google Sheets by formatting it into a 2D array structure.
 * @param {Array<Object>} data - Array of RFI group objects to prepare.
 * @param {string} data[].updatedRfi - The updated RFI text.
 * @param {Array<Object>} data[].projectsAffected - Array of affected projects.
 * @param {string} data[].projectsAffected[].iid - Project identifier.
 * @returns {Array<Array<string>>} 2D array formatted for Google Sheets.
 * @throws {Error} If data structure is invalid or processing fails.
 */
export function prepareRfiCellDataForRfiSpreadsheet(data) {
  try {
    // Validate input
    if (!Array.isArray(data)) {
      throw new Error("Input must be an array");
    }

    return data.map((group, index) => {
      try {
        const row = [];

        // Validate group structure
        if (!group || typeof group !== "object") {
          throw new Error(`Invalid group at index ${index}`);
        }

        // Validate and add updatedRfi
        if (!group.updatedRfi) {
          throw new Error(`Missing updatedRfi at index ${index}`);
        }
        row.push(group.updatedRfi);

        // Validate and process projectsAffected
        if (!Array.isArray(group.projectsAffected)) {
          throw new Error(`Invalid projectsAffected at index ${index}`);
        }

        try {
          const projectsAffected = group.projectsAffected
            .map((project) => {
              if (!project || !project.iid) {
                console.warn(`Missing iid in project at index ${index}`);
                return "";
              }
              return project.iid;
            })
            .filter(Boolean)
            .join(", ");

          row.push(projectsAffected);
        } catch (error) {
          console.error(`Error processing projects at index ${index}:`, error);
          row.push(""); // Add empty string if projects processing fails
        }

        return row;
      } catch (error) {
        console.error(`Error processing group at index ${index}:`, error);
        return ["Error processing RFI", ""]; // Return placeholder row for failed groups
      }
    });
  } catch (error) {
    console.error("Failed to prepare RFI data:", error);
    throw new Error(`Failed to prepare data for spreadsheet: ${error.message}`);
  }
}

/**
 * Calculates Excel range strings for general and specific RFI issues.
 * @param {number} startRowGeneral - The starting row number for general issues.
 * @param {number} startRowSpecific - The starting row number for specific issues.
 * @param {number} generalDataLength - The number of general issue rows.
 * @param {number} specificDataLength - The number of specific issue rows.
 * @returns {Object} Object containing range strings for both issue types.
 * @property {string} rangeForGeneralIssuesRfi - Range string for general issues (e.g., 'C14:D20').
 * @property {string} rangeForSpecificIssuesRfi - Range string for specific issues (e.g., 'C42:D50').
 * @throws {Error} If any input is invalid or range calculation fails.
 */
export const getRfiRanges = (
  startRowGeneral,
  startRowSpecific,
  generalDataLength,
  specificDataLength
) => {
  try {
    // Validate inputs are numbers and positive
    if (typeof startRowGeneral !== "number" || startRowGeneral <= 0) {
      throw new Error("Start row for general issues must be a positive number");
    }
    if (typeof startRowSpecific !== "number" || startRowSpecific <= 0) {
      throw new Error(
        "Start row for specific issues must be a positive number"
      );
    }
    if (typeof generalDataLength !== "number" || generalDataLength < 0) {
      throw new Error("General data length must be a non-negative number");
    }
    if (typeof specificDataLength !== "number" || specificDataLength < 0) {
      throw new Error("Specific data length must be a non-negative number");
    }

    try {
      // Calculate ranges
      const rangeForGeneralIssuesRfi = `C${startRowGeneral}:D${
        startRowGeneral + generalDataLength - 1
      }`;
      const rangeForSpecificIssuesRfi = `C${startRowSpecific}:D${
        startRowSpecific + specificDataLength - 1
      }`;

      // Validate generated ranges
      if (!rangeForGeneralIssuesRfi.match(/^C\d+:D\d+$/)) {
        throw new Error("Invalid general issues range format");
      }
      if (!rangeForSpecificIssuesRfi.match(/^C\d+:D\d+$/)) {
        throw new Error("Invalid specific issues range format");
      }

      return { rangeForGeneralIssuesRfi, rangeForSpecificIssuesRfi };
    } catch (error) {
      console.error("Range calculation failed:", error);
      throw new Error(`Failed to calculate ranges: ${error.message}`);
    }
  } catch (error) {
    console.error("RFI range generation failed:", error);
    throw new Error(`Failed to generate RFI ranges: ${error.message}`);
  }
};

/**
 * Updates Excel spreadsheet data for a specific range and issue type.
 * @param {Object} client - The Microsoft Graph API client instance.
 * @param {string} url - The API endpoint URL for the update.
 * @param {Object} requestBody - The request payload containing the update data.
 * @param {string} issueType - The type of issue being updated (e.g., 'general', 'specific').
 * @returns {Promise<void>} Promise that resolves when update is complete.
 * @throws {Error} If client is invalid.
 * @throws {Error} If URL is invalid.
 * @throws {Error} If request body is invalid.
 * @throws {Error} If API request fails.
 */
export const updateExcelData = async (client, url, requestBody, issueType) => {
  try {
    // Validate inputs
    if (!client || typeof client.api !== "function") {
      throw new Error("Invalid client object");
    }

    if (!url || typeof url !== "string") {
      throw new Error("Invalid URL");
    }

    if (!requestBody || typeof requestBody !== "object") {
      throw new Error("Invalid request body");
    }

    if (!issueType || typeof issueType !== "string") {
      throw new Error("Invalid issue type");
    }

    try {
      // Send a PATCH request to update the data in the specified range
      await client.api(url).patch(requestBody);
      console.log(`Excel spreadsheet updated successfully for ${issueType}.`);
    } catch (apiError) {
      console.error(`API request failed for ${issueType}:`, {
        message: apiError.message,
        status: apiError.statusCode,
        details: apiError.body,
      });
      throw new Error(`Excel update failed: ${apiError.message}`);
    }
  } catch (error) {
    // Log comprehensive error details
    console.error(`Error updating Excel data for ${issueType}:`, {
      error: error.message,
      url: url,
      requestBody: JSON.stringify(requestBody, null, 2),
      stack: error.stack,
    });

    // Rethrow with context
    throw new Error(`Failed to update ${issueType} data: ${error.message}`);
  }
};

/**
 * Calculates Excel cell range based on data dimensions and positioning options.
 * @param {Array<Array<any>>} data - 2D array of data to calculate range for.
 * @param {string} [startCell="A1"] - Starting cell reference (e.g., "A1", "B2").
 * @param {boolean} [preserveOriginalPositions=false] - If true, uses offset positions for RFI data.
 * @returns {string} Excel-style cell range (e.g., "A1:C5").
 * @throws {Error} If data is invalid or cell range calculation fails.
 * @example
 * getCellRange([["a", "b"], ["c", "d"]], "A1", false) // returns "A1:B2"
 * getCellRange([["a", "b"]], "B2", true) // returns "B5:C5"
 */
export const getCellRange = (
  data,
  startCell = "A1",
  preserveOriginalPositions = false
) => {
  try {
    // Validate input data
    if (!Array.isArray(data)) {
      throw new Error("Input data must be an array");
    }

    if (data.length === 0) {
      throw new Error("Input data cannot be empty");
    }

    // Validate startCell format
    if (!/^[A-Z]\d+$/.test(startCell)) {
      throw new Error('Invalid start cell format. Expected format: e.g., "A1"');
    }

    try {
      const numRows = data.length;
      const numCols = data[0] ? data[0].length : 0;

      if (numCols === 0) {
        throw new Error("Data must contain at least one column");
      }

      if (!preserveOriginalPositions) {
        try {
          // For template copying, use the exact positions
          const endCell =
            String.fromCharCode(startCell.charCodeAt(0) + numCols - 1) +
            (parseInt(startCell.slice(1)) + numRows - 1);

          const range = `${startCell}:${endCell}`;
          if (!/^[A-Z]\d+:[A-Z]\d+$/.test(range)) {
            throw new Error("Invalid range format generated");
          }

          return range;
        } catch (error) {
          console.error("Failed to calculate template range:", error);
          throw new Error(
            `Template range calculation failed: ${error.message}`
          );
        }
      }

      try {
        // For RFI data placement, use the offset positions
        const startColChar = "B"; // Always start from column B
        const startRow = 5; // Start from row 5 to shift data up by 9 rows
        const endCell =
          String.fromCharCode(startColChar.charCodeAt(0) + numCols - 1) +
          (startRow + numRows - 1);

        const range = `${startColChar}${startRow}:${endCell}`;
        if (!/^[A-Z]\d+:[A-Z]\d+$/.test(range)) {
          throw new Error("Invalid range format generated");
        }

        return range;
      } catch (error) {
        console.error("Failed to calculate RFI range:", error);
        throw new Error(`RFI range calculation failed: ${error.message}`);
      }
    } catch (error) {
      console.error("Range calculation failed:", error);
      throw new Error(`Failed to calculate cell range: ${error.message}`);
    }
  } catch (error) {
    console.error("Cell range generation failed:", error);
    throw new Error(`Failed to generate cell range: ${error.message}`);
  }
};

/**
 * Handles the directory selection and creates a file selection card.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 * @param {Object} context.activity - The activity object.
 * @param {Object} context.activity.value - The activity value containing directory choice.
 * @param {string} selectedDirectoryId - The ID of the selected directory.
 * @param {Object} [options={}] - Optional configuration for file filtering and display.
 * @param {string} [options.filterPattern] - Pattern to filter files by name.
 * @param {string} [options.customSubheading] - Custom subheading for the card.
 * @param {string} [options.buttonText] - Custom text for the action button.
 * @param {string} [options.action] - Custom action for the card.
 * @returns {Promise<void>} Promise that resolves when the directory selection is handled.
 * @throws {Error} If context or directory ID is invalid.
 * @throws {Error} If file retrieval fails.
 * @throws {Error} If card creation or sending fails.
 */
export async function handleDirectorySelection(
  context,
  selectedDirectoryId,
  options = {}
) {
  try {
    // Validate inputs
    if (!context || !context.activity?.value?.directoryChoice) {
      throw new Error("Invalid context or missing directory choice");
    }

    if (!selectedDirectoryId) {
      throw new Error("Directory ID is required");
    }

    let files;
    try {
      // Get files from directory
      files = await getFileNamesAndIds(selectedDirectoryId);
      if (!Array.isArray(files)) {
        throw new Error("Invalid files response");
      }
    } catch (error) {
      console.error("File retrieval failed:", error);
      throw new Error(`Failed to get files: ${error.message}`);
    }

    try {
      // Parse directory choice
      const selectedDirectory = JSON.parse(
        context.activity.value.directoryChoice
      );
      if (!selectedDirectory?.name) {
        throw new Error("Invalid directory data");
      }
      const directoryName = selectedDirectory.name;

      // Filter files if pattern is provided
      const filteredFiles = options.filterPattern
        ? files.filter((file) => {
            if (!file?.name) {
              console.warn("Invalid file object found, skipping");
              return false;
            }
            return file.name.includes(options.filterPattern);
          })
        : files;

      try {
        // Create and send card
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
        console.error("Card creation or sending failed:", error);
        throw new Error(`Failed to create or send card: ${error.message}`);
      }
    } catch (error) {
      console.error("Directory processing failed:", error);
      throw new Error(`Failed to process directory: ${error.message}`);
    }
  } catch (error) {
    console.error("Directory selection handling failed:", {
      error: error.message,
      directoryId: selectedDirectoryId,
      options: JSON.stringify(options),
    });

    // Attempt to notify user of error
    try {
      await context.sendActivity(
        `Error retrieving files from the selected directory: ${error.message}`
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }

    throw new Error(`Directory selection failed: ${error.message}`);
  }
}

/**
 * Extracts RFI response data from the RFI spreadsheet for auditor notes generation.
 * @param {Array<Array<string>>} data - The spreadsheet data as a 2D array.
 * @param {Array<[number, number]>} ranges - Array of [start, end] row ranges to extract.
 * @returns {Array<Object>} Array of RFI response objects.
 * @property {string} rfiNumber - The RFI identifier from column A.
 * @property {string} issuesIdentified - The issues from column C.
 * @property {string} acpResponse - The response from column E.
 * @throws {Error} If data structure is invalid or extraction fails.
 */
export const extractRfiResponseData = (data, ranges) => {
  try {
    // Validate inputs
    if (!Array.isArray(data)) {
      throw new Error("Input data must be an array");
    }
    if (!Array.isArray(ranges)) {
      throw new Error("Ranges must be an array");
    }

    return ranges.flatMap(([start, end], rangeIndex) => {
      try {
        // Validate range values
        if (typeof start !== "number" || typeof end !== "number") {
          throw new Error(`Invalid range format at index ${rangeIndex}`);
        }
        if (start <= 0 || end <= 0) {
          throw new Error(
            `Range values must be positive at index ${rangeIndex}`
          );
        }
        if (start > end) {
          throw new Error(
            `Start cannot be greater than end at index ${rangeIndex}`
          );
        }
        if (end > data.length) {
          throw new Error(
            `Range end exceeds data length at index ${rangeIndex}`
          );
        }

        try {
          return data
            .slice(start - 1, end)
            .filter((row, rowIndex) => {
              if (!Array.isArray(row)) {
                console.warn(
                  `Invalid row at index ${rowIndex + start - 1}, skipping`
                );
                return false;
              }
              return row[0] || row[2]; // Filter rows with data in columns A or C
            })
            .map((row, rowIndex) => {
              try {
                const rfiData = {
                  rfiNumber: row[0] || "", // Column A
                  issuesIdentified: row[1] || "", // Column C
                  acpResponse: row[3] || "", // Column E
                };

                // Validate required fields
                if (!rfiData.issuesIdentified || !rfiData.acpResponse) {
                  console.warn(
                    `Missing required data at row ${rowIndex + start}`
                  );
                  return null;
                }

                return rfiData;
              } catch (error) {
                console.error(
                  `Error processing row ${rowIndex + start}:`,
                  error
                );
                return null;
              }
            })
            .filter(Boolean); // Remove null entries
        } catch (error) {
          console.error(`Error processing range ${start}-${end}:`, error);
          return [];
        }
      } catch (error) {
        console.error(`Error validating range at index ${rangeIndex}:`, error);
        return [];
      }
    });
  } catch (error) {
    console.error("RFI response data extraction failed:", error);
    throw new Error(`Failed to extract RFI response data: ${error.message}`);
  }
};

/**
 * Prepares RFI data for batch update by organizing into general and specific issue arrays.
 * @param {Array<Object>} updatedResponseData - Array of RFI response objects to process.
 * @param {string} updatedResponseData[].rfiNumber - The RFI identifier (e.g., "G.1", "S.1").
 * @param {string} updatedResponseData[].auditorNotes - The auditor's notes for the RFI.
 * @returns {Object} Organized arrays ready for spreadsheet update.
 * @property {Array<Array<string>>} generalArray - Array for general issues (rows F14:F34).
 * @property {Array<Array<string>>} specificArray - Array for specific issues (rows F42:F141).
 * @throws {Error} If input data is invalid or processing fails.
 */
export const prepareDataForBatchUpdate = (updatedResponseData) => {
  try {
    // Validate input
    if (!Array.isArray(updatedResponseData)) {
      throw new Error("Input must be an array");
    }

    try {
      // Split and validate the data into general and specific arrays
      const generalIssues = updatedResponseData
        .filter((response) => {
          if (!response?.rfiNumber) {
            console.warn("Found response without RFI number, skipping");
            return false;
          }
          return response.rfiNumber.startsWith("G.");
        })
        .sort((a, b) => {
          try {
            return (
              parseInt(a.rfiNumber.slice(2)) - parseInt(b.rfiNumber.slice(2))
            );
          } catch (error) {
            console.error("Error sorting general issues:", error);
            return 0;
          }
        });

      const specificIssues = updatedResponseData
        .filter((response) => {
          if (!response?.rfiNumber) {
            console.warn("Found response without RFI number, skipping");
            return false;
          }
          return response.rfiNumber.startsWith("S.");
        })
        .sort((a, b) => {
          try {
            return (
              parseInt(a.rfiNumber.slice(2)) - parseInt(b.rfiNumber.slice(2))
            );
          } catch (error) {
            console.error("Error sorting specific issues:", error);
            return 0;
          }
        });

      // Create arrays with correct sizes
      const generalArray = Array(21).fill([""]); // F14 to F34 (21 rows)
      const specificArray = Array(100).fill([""]); // F42 to F141 (100 rows)

      try {
        // Process general issues
        generalIssues.forEach((item) => {
          try {
            const index = parseInt(item.rfiNumber.slice(2)) - 1;
            if (index >= 0 && index < 21) {
              generalArray[index] = [item.auditorNotes || ""];
            } else {
              console.warn(
                `General issue index out of range: ${item.rfiNumber}`
              );
            }
          } catch (error) {
            console.error(
              `Error processing general issue ${item.rfiNumber}:`,
              error
            );
          }
        });

        // Process specific issues
        specificIssues.forEach((item) => {
          try {
            const index = parseInt(item.rfiNumber.slice(2)) - 1;
            if (index >= 0 && index < 100) {
              specificArray[index] = [item.auditorNotes || ""];
            } else {
              console.warn(
                `Specific issue index out of range: ${item.rfiNumber}`
              );
            }
          } catch (error) {
            console.error(
              `Error processing specific issue ${item.rfiNumber}:`,
              error
            );
          }
        });

        return {
          generalArray,
          specificArray,
        };
      } catch (error) {
        console.error("Error processing issues:", error);
        throw new Error(`Failed to process issues: ${error.message}`);
      }
    } catch (error) {
      console.error("Error preparing batch data:", error);
      throw new Error(`Failed to prepare batch data: ${error.message}`);
    }
  } catch (error) {
    console.error("Batch update preparation failed:", error);
    throw new Error(
      `Failed to prepare data for batch update: ${error.message}`
    );
  }
};

/**
 * Creates and sends an adaptive card update message in Teams.
 * @param {Object} context - The Teams activity context.
 * @param {Function} context.sendActivity - Function to send the activity to Teams.
 * @param {string} text - The main text content for the update.
 * @param {string} [userMessage=""] - Optional additional user message.
 * @param {string} [emoji="💬"] - Optional emoji to display with the update.
 * @param {string} [style="default"] - Optional card style ("default", "warning", "error", etc.).
 * @returns {Promise<void>} Promise that resolves when the update is sent.
 * @throws {Error} If context is invalid.
 * @throws {Error} If card creation fails.
 * @throws {Error} If sending activity fails.
 */
export const createTeamsUpdate = async (
  context,
  text,
  userMessage = "",
  emoji = "💬",
  style = "default"
) => {
  try {
    // Validate inputs
    if (!context || typeof context.sendActivity !== "function") {
      throw new Error("Invalid context: missing sendActivity function");
    }

    if (!text || typeof text !== "string") {
      throw new Error("Text content is required and must be a string");
    }

    try {
      // Create the Teams update card
      const teamsUpdateCard = createTeamsUpdateCard(
        text,
        userMessage,
        emoji,
        style
      );

      if (!teamsUpdateCard) {
        throw new Error("Failed to create Teams update card");
      }

      try {
        // Send the activity to Teams
        await context.sendActivity({
          type: "message",
          attachments: [
            {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: teamsUpdateCard,
            },
          ],
        });
      } catch (sendError) {
        console.error("Failed to send Teams activity:", sendError);
        throw new Error(`Activity sending failed: ${sendError.message}`);
      }
    } catch (cardError) {
      console.error("Card creation failed:", cardError);
      throw new Error(`Card creation failed: ${cardError.message}`);
    }
  } catch (error) {
    console.error("Teams update creation failed:", {
      error: error.message,
      text: text,
      userMessage: userMessage,
      style: style,
    });
    throw new Error(`Failed to create Teams update: ${error.message}`);
  }
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
