import { updateRfiDataWithAzureGptQuery } from "./llmQueries.js";

// Helper function to convert a column index to a letter (e.g., 0 -> A, 4 -> E, 26 -> AA)
function getColumnLetter(colIndex) {
  let letter = "";
  while (colIndex >= 0) {
    letter = String.fromCharCode((colIndex % 26) + 65) + letter;
    colIndex = Math.floor(colIndex / 26) - 1;
  }
  return letter;
}

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

// Function to group data by RFI
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

// Function to update RFI rows
export async function updateRfiCellData(filteredRows) {
  // Groups all common rfi's into an object with the rfi, and projectsAffected array
  const groupedData = groupByRFI(filteredRows);

  // Extract all RFI attributes into an array
  const allRfiAttributes = groupedData.map((group) => group.rfi);

  // Amend RFI data using OpenAI
  const updatedRfiAttributes = await updateRfiDataWithAzureGptQuery(
    allRfiAttributes
  );

  // Parse the updated RFI attributes
  const parsedUpdatedRfiAttributes = JSON.parse(updatedRfiAttributes);

  // Add the updatedRfi attribute to each object in the groupedData array
  groupedData.forEach((group, index) => {
    group.updatedRfi = parsedUpdatedRfiAttributes[index];
  });

  return groupedData;
}

// Function to prepare data for updating Google Sheets
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

// Function to get the ranges for general and specific issues RFI
export const getRfiRanges = (
  startRowGeneral,
  startRowSpecific,
  generalDataLength,
  specificDataLength
) => {
  const rangeForGeneralIssuesRfi = `A${startRowGeneral}:B${
    startRowGeneral + generalDataLength - 1
  }`;
  const rangeForSpecificIssuesRfi = `A${startRowSpecific}:B${
    startRowSpecific + specificDataLength - 1
  }`;

  return { rangeForGeneralIssuesRfi, rangeForSpecificIssuesRfi };
};

// Function to update the Excel spreadsheet for a given range and data
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

// Function to calculate the cell range based on the data array with an offset
export const getCellRange = (data, startCell = "A1") => {
  const numRows = data.length; // Number of rows in the data
  const numCols = data[0] ? data[0].length : 0; // Number of columns in the first row

  // Calculate the end cell based on the number of rows and columns
  const endCell =
    String.fromCharCode(startCell.charCodeAt(0) + numCols - 1) +
    (parseInt(startCell.slice(1)) + numRows - 1);

  return `${startCell}:${endCell}`; // Return the range in A1:B2 format
};
