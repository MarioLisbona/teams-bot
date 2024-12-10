import { filterRowsForRFICells } from "./utils.js";
import {
  updateRfiCellData,
  prepareRfiCellDataForRfiSpreadsheet,
  getRfiRanges,
  updateExcelData,
  getCellRange,
} from "./utils.js";
import { clearWorksheetRange } from "./clearWorksheet.js";
import { copyFileInOneDrive } from "./copyFile.js";

// Retrieve data from the Testing Excel sheet
// Extracts all the RFI cells in the sheet and used OpenAI to create an updateRFI string for the RFI Spreadsheet
export const processTesting = async (client, userId, workbookId, sheetName) => {
  try {
    // Construct the URL for the Excel file's used range
    const range = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/usedRange`;

    // Fetch the data from all non-empty rows in the sheet
    const response = await client.api(range).get();

    // Extract the values from the response
    const data = response.values;

    // Filter out non-empty rows from the data returned by usedRange
    // useRange is returning alot of cells that are empty, so we need to filter out the empty rows
    const nonEmptyRows = data.filter((row) =>
      row.some((cell) => cell !== null && cell !== "")
    );

    console.log({ "data length": data.length });
    console.log({ "nonEmptyRows length": nonEmptyRows.length });

    // Filter the data to only include rows where a non-empty cell contains the substring "RFI"
    // returns an array of objects with the rfi, cellReference and iid attributes
    const rfiCellData = filterRowsForRFICells(nonEmptyRows);

    console.log("Workbook ID:", workbookId);
    console.log("Rows with RFI Data:", rfiCellData.length);

    // rfi value from each object in rfiCellData array is passed to OpenAI to create an updateRFI string
    // the updatedRFI string is added to each object
    const updatedRfiCellData = await updateRfiCellData(rfiCellData);

    // // Write updatedRfiCellData to a json file in the root of the project
    // fs.writeFileSync(
    //   "updatedRfiCellTestData.json",
    //   JSON.stringify(updatedRfiCellData)
    // );

    return updatedRfiCellData;
  } catch (error) {
    // Log the error if data retrieval fails
    console.error("Error retrieving data:", error.message);
    console.error("Full error details:", error);
  }
};

// Function to update an Excel spreadsheet with new data
export const updateRfiSpreadsheet = async (
  client,
  userId,
  workbookId,
  sheetName,
  rfiCellData
) => {
  // The array of ranges to clear - only clear cell data, not headings or images
  // This range is based off the RFI Spreadsheet in the main client workbook
  const ranges = ["C14:I34", "C42:I141"];

  // Clear the ranges before updating the RFI spreadsheet
  await clearWorksheetRange(userId, workbookId, sheetName, ranges);

  // Filter groupedData into two arrays: one where projectsAffected.length >= 4, and one where projectsAffected.length < 4
  // This is done to separate the RFI's into general and specific issues
  const generalIssuesRfi =
    rfiCellData.filter((group) => group.projectsAffected.length >= 4).length > 0
      ? rfiCellData.filter((group) => group.projectsAffected.length >= 4)
      : []; // Assign only if at least one group has projectsAffected.length >= 4

  const specificIssuesRfi =
    rfiCellData.filter((group) => group.projectsAffected.length < 4).length > 0
      ? rfiCellData.filter((group) => group.projectsAffected.length < 4)
      : []; // Assign only if at least one group has projectsAffected.length < 4

  // Prepare data for both sets of groups
  const generalIssuesRfiData =
    generalIssuesRfi.length > 0
      ? prepareRfiCellDataForRfiSpreadsheet(generalIssuesRfi)
      : []; // Call only if generalIssuesRfi is not empty

  const specificIssuesRfiData =
    specificIssuesRfi.length > 0
      ? prepareRfiCellDataForRfiSpreadsheet(specificIssuesRfi)
      : []; // Call only if specificIssuesRfi is not empty

  // Define starting row for both cases
  const startRowGeneralIssuesRfi = 14; // Start from row 7 for general issues
  const startRowSpecificIssuesRfi = 42; // Start from row 18 for specific issues

  // Get the ranges for general and specific issues RFI for the update request
  const { rangeForGeneralIssuesRfi, rangeForSpecificIssuesRfi } = getRfiRanges(
    startRowGeneralIssuesRfi,
    startRowSpecificIssuesRfi,
    generalIssuesRfiData.length > 0 ? generalIssuesRfiData.length : 0, // Use length only if generalIssuesRfiData is not empty
    specificIssuesRfiData.length > 0 ? specificIssuesRfiData.length : 0 // Use length only if specificIssuesRfiData is not empty
  );

  console.log("Range for general issues RFI:", rangeForGeneralIssuesRfi);
  console.log("Range for specific issues RFI:", rangeForSpecificIssuesRfi);

  // Construct the URL for the Excel file's using ranges for general and specific issues RFI
  const urlGeneralIssuesRfi = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='${rangeForGeneralIssuesRfi}')`;
  const urlSpecificIssuesRfi = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='${rangeForSpecificIssuesRfi}')`;

  // Prepare the request body with the data to update only if generalIssuesRfiData is not empty
  if (generalIssuesRfiData.length > 0) {
    const requestBodyGeneralIssuesRfi = {
      values: generalIssuesRfiData,
    };

    // Call the new function for general issues RFI
    await updateExcelData(
      client,
      urlGeneralIssuesRfi,
      requestBodyGeneralIssuesRfi,
      "general issues RFI"
    );
  }

  // Prepare the request body with the data to update for specific issues RFI only if specificIssuesRfiData is not empty
  if (specificIssuesRfiData.length > 0) {
    const requestBodySpecificIssuesRfi = {
      values: specificIssuesRfiData,
    };

    // Call the new function for specific issues RFI
    await updateExcelData(
      client,
      urlSpecificIssuesRfi,
      requestBodySpecificIssuesRfi,
      "specific issues RFI"
    );
  }
};

// Function to copy a worksheet to a new spreadsheet
export const copyWorksheetToClientWorkbook = async (
  client,
  userId,
  sourceWorkbookId,
  sourceWorksheetName,
  clientName,
  targetDirectoryId
) => {
  const newWorksheetName = "RFI Responses";

  // Get the RFI Client Template workbook ID from the .env file
  const templateWorkbookId = process.env.RFI_CLIENT_TEMPLATE_ID;

  console.log("Template workbook ID:", templateWorkbookId);

  // Pass the targetDirectoryId to copyFileInOneDrive
  const { newWorkbookId, newWorkbookName } = await copyFileInOneDrive(
    userId,
    templateWorkbookId,
    `RFI Responses - ${clientName}.xlsx`,
    targetDirectoryId
  );

  console.log("New workbook created with ID:", newWorkbookId);

  // Extract the data from the source worksheet
  console.log("Fetching data from source worksheet...");
  const existingData = await client
    .api(
      `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${sourceWorkbookId}/workbook/worksheets/${sourceWorksheetName}/usedRange`
    )
    .get();

  // Check if existingData is valid
  if (!existingData || !existingData.values) {
    throw new Error("No data found in the existing worksheet.");
  }

  const cellValuesData = existingData.values;
  console.log(`Found ${cellValuesData.length} rows of data to copy`);

  // Calculate the cell range for the data
  const newRangeAddress = getCellRange(cellValuesData, "A1", true);
  console.log(`Writing to range: ${newRangeAddress}`);

  // Try to write to the worksheet named "RFI Spreadsheet Template" first
  try {
    await client
      .api(
        `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${newWorkbookId}/workbook/worksheets('${newWorksheetName}')/range(address='${newRangeAddress}')`
      )
      .patch({
        values: cellValuesData,
      });
    console.log("Data written successfully to RFI Spreadsheet Template");
  } catch (error) {
    console.log("Error writing to RFI Spreadsheet Template:", error.message);
    // Try writing to Sheet1 as fallback
    try {
      await client
        .api(
          `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${newWorkbookId}/workbook/worksheets('Sheet1')/range(address='${newRangeAddress}')`
        )
        .patch({
          values: cellValuesData,
        });
      console.log("Data written successfully to Sheet1");
    } catch (fallbackError) {
      console.log("Error writing to Sheet1:", fallbackError.message);
      throw fallbackError;
    }
  }

  return { newWorkbookId, newWorkbookName };
};
