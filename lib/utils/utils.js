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
 * Converts numeric column index to Excel column reference.
 *
 * @description
 * Converts zero-based indices to Excel-style column references:
 * - Single letter: A-Z for indices 0-25
 * - Multiple letters: AA-ZZ for indices 26+
 * Uses recursive division by 26 for multi-letter columns
 *
 * Validation includes:
 * - Input type checking (must be number)
 * - Negative value prevention
 * - Result verification
 *
 * @param {number} colIndex - Zero-based column index
 *
 * @throws {Error} When input is not a number
 * @throws {Error} When input is negative
 * @throws {Error} When letter generation fails
 * @returns {string} Excel column reference
 *
 * @example
 * getColumnLetter(0)   // returns "A"
 * getColumnLetter(25)  // returns "Z"
 * getColumnLetter(26)  // returns "AA"
 * getColumnLetter(701) // returns "ZZ"
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
 * Extracts RFI information from Excel worksheet data.
 *
 * @description
 * Processing workflow:
 * 1. Skips first two rows (headers)
 * 2. For each remaining row:
 *    - Extracts IID from column C
 *    - Searches cells for "RFI" substring
 *    - Generates Excel references for matches
 *    - Skips column AJ (index 35)
 * 3. Filters out empty results
 *
 * Cell processing includes:
 * - Type validation for string content
 * - Excel reference generation
 * - IID association
 * - Error handling per cell
 *
 * @param {Array<Array<any>>} data - Raw worksheet data matrix
 * @param {any} data[][] - Cell values (strings, numbers, etc.)
 *
 * @throws {Error} When input is not an array
 * @throws {Error} When row format is invalid
 * @throws {Error} When cell processing fails
 * @returns {Array<Array<Object>>} Processed RFI data:
 *   - rfi: RFI content string
 *   - cellReference: Excel cell reference (e.g., "E5")
 *   - iid: Associated IID value
 *
 * @example
 * const rfiData = filterRowsForRFICells([
 *   ["Header1", "Header2"],
 *   ["SubHeader1", "SubHeader2"],
 *   ["Data", "RFI: Review needed", "IID123"]
 * ]);
 * // Returns [[{ rfi: "RFI: Review needed",
 * //            cellReference: "B3",
 * //            iid: "IID123" }]]
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
 * Groups RFI entries by identifier and collects project associations.
 *
 * @description
 * Grouping process:
 * 1. Creates unique RFI groups
 * 2. Collects project references for each RFI
 * 3. Maintains cell references and IIDs
 * 4. Validates data integrity
 *
 * @param {Array<Array<Object>>} filteredRows - Nested RFI data structure
 * @param {Object} filteredRows[][] - Individual RFI entries
 * @param {string} filteredRows[][].rfi - RFI identifier/content
 * @param {string} filteredRows[][].cellReference - Excel cell location
 * @param {string} filteredRows[][].iid - Project identifier
 *
 * @throws {Error} When input structure is invalid
 * @throws {Error} When required properties are missing
 * @returns {Array<Object>} Grouped RFI data:
 *   - rfi: RFI identifier
 *   - projectsAffected: Array of associated projects
 *     - cellReference: Excel location
 *     - iid: Project identifier
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
 * Enhances RFI content with context and action items.
 *
 * @description
 * Processing steps:
 * 1. Groups common RFIs
 * 2. Cleans RFI text
 * 3. Determines appropriate action items based on content:
 *    - Documentation requests
 *    - Invoice/declaration reviews
 *    - General clarifications
 * 4. Formats final text with context
 *
 * @param {Array} filteredRows - RFI entries to process
 *
 * @throws {Error} When input is not an array
 * @throws {Error} When RFI group structure is invalid
 * @throws {Error} When RFI text is invalid
 * @returns {Array<Object>} Processed RFI data with:
 *   - Original RFI properties
 *   - updatedRfi: Enhanced RFI text with context
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
 * Formats RFI data for spreadsheet export.
 *
 * @description
 * Formatting process:
 * 1. Validates input structure
 * 2. Creates rows for each RFI group
 * 3. Combines project identifiers
 * 4. Handles missing or invalid data gracefully
 *
 * Generated columns:
 * 1. Enhanced RFI text with context
 * 2. Comma-separated list of project IDs
 *
 * @param {Array<Object>} data - RFI groups to format
 * @param {string} data[].updatedRfi - Enhanced RFI text
 * @param {Array<Object>} data[].projectsAffected - Project references
 * @param {string} data[].projectsAffected[].iid - Project identifier
 *
 * @throws {Error} When input structure is invalid
 * @throws {Error} When required properties are missing
 * @returns {Array<Array<string>>} Spreadsheet-ready 2D array:
 *   - [0]: Enhanced RFI text
 *   - [1]: Project IDs (comma-separated)

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
 * Calculates Excel range references for RFI data sections.
 *
 * @description
 * Generates range strings by:
 * 1. Validating input parameters
 * 2. Calculating range boundaries
 * 3. Formatting Excel-style references
 * 4. Validating generated ranges
 *
 * Range format: 'C{startRow}:D{endRow}'
 * - Column C: First column of range
 * - Column D: Last column of range
 * - Rows: Calculated based on start row and data length
 *
 * @param {number} startRowGeneral - First row for general issues
 * @param {number} startRowSpecific - First row for specific issues
 * @param {number} generalDataLength - Number of general issue rows
 * @param {number} specificDataLength - Number of specific issue rows
 *
 * @throws {Error} When any input is not a positive number
 * @throws {Error} When data lengths are negative
 * @throws {Error} When range calculation fails
 * @returns {Object} Range references:
 *   - rangeForGeneralIssuesRfi: Range for general issues
 *   - rangeForSpecificIssuesRfi: Range for specific issues
 *
 * @example
 * const ranges = getRfiRanges(14, 42, 7, 9);
 * // Returns:
 * // {
 * //   rangeForGeneralIssuesRfi: "C14:D20",
 * //   rangeForSpecificIssuesRfi: "C42:D50"
 * // }
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
 * Updates Excel ranges via Microsoft Graph API.
 *
 * @description
 * Update process:
 * 1. Validates input parameters
 * 2. Sends PATCH request to Graph API
 * 3. Handles response and errors
 * 4. Provides detailed error logging
 *
 * Error handling includes:
 * - Client validation
 * - URL format checking
 * - Request body verification
 * - API error details
 * - Comprehensive error logging
 *
 * @param {Object} client - Graph API client
 * @param {Function} client.api - Graph API request method
 * @param {string} url - Graph API endpoint
 * @param {Object} requestBody - Update payload
 * @param {string} issueType - Issue category identifier
 *
 * @throws {Error} When client is invalid/missing
 * @throws {Error} When URL is invalid/missing
 * @throws {Error} When request body is invalid
 * @throws {Error} When API request fails
 * @returns {Promise<void>} Resolves on successful update
 *
 * @example
 * await updateExcelData(
 *   graphClient,
 *   "/sites/{site}/drive/items/{file}/workbook/range",
 *   { values: [["Data1", "Data2"]] },
 *   "general"
 * );
 *
 * @requires Microsoft Graph API client
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
 * Generates Excel cell ranges based on data dimensions.
 *
 * @description
 * Range calculation modes:
 * 1. Template mode (preserveOriginalPositions=false):
 *    - Uses provided start cell
 *    - Calculates end cell based on data dimensions
 *    - Maintains relative positioning
 *
 * 2. RFI mode (preserveOriginalPositions=true):
 *    - Fixed start at column B, row 5
 *    - Calculates end cell based on data size
 *    - Applies 9-row upward shift
 *
 * Validation includes:
 * - Data array structure and content
 * - Start cell format (letter + number)
 * - Generated range format
 * - Column count verification
 *
 * @param {Array<Array<any>>} data - 2D array to calculate range for
 * @param {string} [startCell="A1"] - Starting cell reference
 * @param {boolean} [preserveOriginalPositions=false] - Use RFI positioning
 *
 * @throws {Error} When data array is invalid/empty
 * @throws {Error} When start cell format is invalid
 * @throws {Error} When range calculation fails
 * @returns {string} Excel range (e.g., "A1:C5")
 *
 * @example
 * // Template mode
 * getCellRange([["A", "B"], ["C", "D"]], "B2")
 * // Returns "B2:C3"
 *
 * // RFI mode
 * getCellRange([["A", "B"]], "A1", true)
 * // Returns "B5:C5"
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
 * Processes directory selection and generates file selection interface.
 *
 * @description
 * Complete workflow:
 * 1. Input validation
 *    - Context object structure
 *    - Directory ID presence
 *    - Directory choice data
 *
 * 2. File retrieval and processing
 *    - Fetches directory contents
 *    - Validates file list structure
 *    - Applies optional filtering
 *
 * 3. Card generation and display
 *    - Creates adaptive card
 *    - Configures custom elements
 *    - Handles sending to Teams
 *
 * Error handling includes:
 * - Input validation failures
 * - File retrieval issues
 * - Card creation problems
 * - Teams communication errors
 *
 * @param {Object} context - Teams bot context
 * @param {Object} context.activity - Bot activity data
 * @param {Object} context.activity.value - Selection values
 * @param {string} selectedDirectoryId - Target directory ID
 * @param {Object} [options={}] - Configuration options
 * @param {string} [options.filterPattern] - File name filter
 * @param {string} [options.customSubheading] - Card subheading
 * @param {string} [options.buttonText] - Action button text
 * @param {string} [options.action] - Card action type
 *
 * @throws {Error} When context/directory validation fails
 * @throws {Error} When file retrieval fails
 * @throws {Error} When card creation/sending fails
 * @returns {Promise<void>}
 *
 * @requires CardFactory from botbuilder
 * @requires createFileSelectionCard from adaptiveCards
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
      await createTeamsUpdate(
        context,
        `Error retrieving files from the selected directory: ${error.message}`,
        "",
        "❌",
        "attention"
      );
    } catch (notifyError) {
      console.error("Failed to send error notification:", notifyError);
    }

    throw new Error(`Directory selection failed: ${error.message}`);
  }
}

/**
 * Extracts RFI data from specified spreadsheet ranges.
 *
 * @description
 * Processing workflow:
 * 1. Input validation
 *    - Data array structure
 *    - Range array format
 *    - Range value boundaries
 *
 * 2. Data extraction per range:
 *    - Validates range boundaries
 *    - Processes rows within range
 *    - Extracts required columns:
 *      * Column A: RFI number
 *      * Column C: Issues identified
 *      * Column E: ACP response
 *
 * 3. Data validation:
 *    - Row structure verification
 *    - Required field presence
 *    - Data format checking
 *
 * Error handling:
 * - Invalid range values
 * - Missing required data
 * - Malformed row structure
 * - Range boundary violations
 *
 * @param {Array<Array<string>>} data - Spreadsheet data matrix
 * @param {Array<[number, number]>} ranges - Row ranges to process
 *
 * @throws {Error} When data structure is invalid
 * @throws {Error} When ranges are invalid
 * @returns {Array<Object>} Processed RFI data:
 *   - rfiNumber: Identifier from column A
 *   - issuesIdentified: Issues from column C
 *   - acpResponse: Response from column E
 *
 * @example
 * const rfiData = extractRfiResponseData(
 *   spreadsheetData,
 *   [[1, 5], [10, 15]]
 * );
 * // Returns [{
 * //   rfiNumber: "RFI-001",
 * //   issuesIdentified: "Missing documentation",
 * //   acpResponse: "Documents provided"
 * // }, ...]
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
 * Organizes RFI responses for spreadsheet batch updates.
 *
 * @description
 * Processing workflow:
 * 1. Separates responses into general (G.*) and specific (S.*) issues
 * 2. Sorts each category by numeric identifier
 * 3. Creates fixed-size arrays:
 *    - General: 21 rows (F14:F34)
 *    - Specific: 100 rows (F42:F141)
 * 4. Maps responses to correct array positions
 *
 * Validation includes:
 * - RFI number format and presence
 * - Array index boundaries
 * - Data structure integrity
 *
 * @param {Array<Object>} updatedResponseData - RFI responses
 * @param {string} updatedResponseData[].rfiNumber - "G.n" or "S.n" format
 * @param {string} updatedResponseData[].auditorNotes - Evaluation notes
 *
 * @throws {Error} When input structure is invalid
 * @throws {Error} When processing fails
 * @returns {Object} Formatted arrays:
 *   - generalArray: 21-row array for F14:F34
 *   - specificArray: 100-row array for F42:F141
 *
 * @example
 * const data = prepareDataForBatchUpdate([
 *   { rfiNumber: "G.1", auditorNotes: "Approved" },
 *   { rfiNumber: "S.1", auditorNotes: "Needs review" }
 * ]);
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
 * Creates and sends Teams adaptive card updates.
 *
 * @description
 * Card creation process:
 * 1. Validates input parameters
 * 2. Generates adaptive card with:
 *    - Main text content
 *    - Optional user message
 *    - Configurable emoji
 *    - Style variations
 * 3. Sends via Teams activity
 *
 * Styles supported:
 * - "default": Standard update
 * - "warning": Warning message
 * - "error": Error notification
 *
 * @param {Object} context - Teams context
 * @param {Function} context.sendActivity - Message sender
 * @param {string} text - Primary message
 * @param {string} [userMessage=""] - Secondary message
 * @param {string} [emoji="💬"] - Message emoji
 * @param {string} [style="default"] - Card style
 *
 * @throws {Error} When context is invalid
 * @throws {Error} When card creation fails
 * @throws {Error} When sending fails
 * @returns {Promise<void>}
 *
 * @example
 * await createTeamsUpdate(
 *   context,
 *   "Processing complete",
 *   "All items updated",
 *   "✅",
 *   "default"
 * );
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
 * Processes client responses in rate-limited batches.
 *
 * @description
 * Processing workflow:
 * 1. Divides responses into 6-item batches
 * 2. For each batch:
 *    - Generates analysis prompt
 *    - Sends progress updates
 *    - Calls Azure GPT API
 *    - Parses and validates results
 * 3. Handles errors per batch
 *
 * Error handling:
 * - Continues processing on batch failures
 * - Sends warning notifications
 * - Maintains partial results
 *
 * @param {Object} context - Teams context
 * @param {Function} context.sendActivity - Update sender
 * @param {Array<Object>} processedClientResponses - Client data
 * @param {string} processedClientResponses[].rfiNumber - RFI ID
 * @param {string} processedClientResponses[].response - Client text
 *
 * @throws {Error} When batch processing fails
 * @throws {Error} When API calls fail
 * @throws {Error} When parsing fails
 * @returns {Promise<Array>} Processed responses
 *
 * @requires analyseAcpResponsePrompt
 * @requires azureGptQuery
 */
export const batchProcessClientResponses = async (
  context,
  processedClientResponses
) => {
  try {
    // Validate inputs
    if (!context || typeof context.sendActivity !== "function") {
      throw new Error("Invalid context object");
    }
    if (!Array.isArray(processedClientResponses)) {
      throw new Error("processedClientResponses must be an array");
    }

    const batchSize = 6;
    const allResponses = [];
    const totalBatches = Math.ceil(processedClientResponses.length / batchSize);

    for (let i = 0; i < processedClientResponses.length; i += batchSize) {
      try {
        // Process batch
        const clientResponsesBatch = processedClientResponses.slice(
          i,
          i + batchSize
        );

        try {
          // Generate prompt
          const prompt = analyseAcpResponsePrompt(
            knowledgeBase,
            clientResponsesBatch
          );

          // Update progress
          await createTeamsUpdate(
            context,
            `Processing client responses batch ${
              Math.floor(i / batchSize) + 1
            }/${totalBatches}...`,
            "",
            "💭",
            "default"
          );

          try {
            // Generate and parse Azure GPT response
            const azureResponse = await azureGptQuery(prompt);
            if (!azureResponse) {
              throw new Error("Empty response from Azure GPT");
            }

            try {
              const batchResults = JSON.parse(azureResponse);
              if (!Array.isArray(batchResults)) {
                throw new Error("Invalid response format from Azure GPT");
              }
              allResponses.push(...batchResults);
            } catch (parseError) {
              console.error("Failed to parse Azure GPT response:", parseError);
              throw new Error(`Response parsing failed: ${parseError.message}`);
            }
          } catch (apiError) {
            console.error("Azure GPT query failed:", apiError);
            throw new Error(`API request failed: ${apiError.message}`);
          }
        } catch (error) {
          console.error(
            `Error processing batch ${Math.floor(i / batchSize) + 1}:`,
            error
          );
          throw error;
        }
      } catch (batchError) {
        console.error(
          `Batch ${Math.floor(i / batchSize) + 1} failed:`,
          batchError
        );
        // Continue with next batch instead of failing completely
        await createTeamsUpdate(
          context,
          `Warning: Batch ${
            Math.floor(i / batchSize) + 1
          } failed, continuing with next batch...`,
          "",
          "⚠️",
          "warning"
        );
      }
    }

    return allResponses;
  } catch (error) {
    console.error("Client response processing failed:", error);
    throw new Error(`Failed to process client responses: ${error.message}`);
  }
};

/**
 * Splits array into fixed-size chunks.
 *
 * @description
 * Divides input array into smaller arrays of specified size.
 * Last chunk may be smaller if array length isn't evenly divisible.
 *
 * Validation includes:
 * - Array input type
 * - Positive integer chunk size
 *
 * @param {Array} array - Input array
 * @param {number} size - Chunk size
 *
 * @throws {Error} When input isn't array
 * @throws {Error} When size isn't positive integer
 * @returns {Array<Array>} Array of chunks
 *
 * @example
 * chunk([1,2,3,4,5], 2)
 * // Returns [[1,2], [3,4], [5]]
 */
export function chunk(array, size) {
  try {
    // Validate inputs
    if (!Array.isArray(array)) {
      throw new Error("Input must be an array");
    }
    if (!Number.isInteger(size) || size <= 0) {
      throw new Error("Chunk size must be a positive integer");
    }

    const chunked = [];
    for (let i = 0; i < array.length; i += size) {
      chunked.push(array.slice(i, i + size));
    }
    return chunked;
  } catch (error) {
    console.error("Array chunking failed:", error);
    throw new Error(`Failed to chunk array: ${error.message}`);
  }
}

export function createConversationReference(
  conversationId,
  channelId,
  serviceUrl,
  tenantId
) {
  return {
    channelId: channelId,
    serviceUrl: serviceUrl,
    conversation: { id: conversationId },
    tenantId: tenantId,
  };
}

export async function continueTeamsConversation(
  adapter,
  conversationReference,
  card
) {
  await adapter.continueConversation(
    conversationReference,
    async (turnContext) => {
      try {
        await turnContext.sendActivity({
          attachments: [
            {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: card,
            },
          ],
        });
      } catch (error) {
        console.error("Failed to send activity:", error);
        throw new Error(
          `Failed to send workflow progress notification: ${error.message}`
        );
      }
    }
  );
}
