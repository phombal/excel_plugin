import { Logger } from '../utils/logger';
import { addMessageToChat } from '../ui/chat.js';
import * as XLSX from 'xlsx';

export async function createPivotTable(context, params) {
  Logger.startOperation("Creating Pivot Table");
  try {
    // Get the source worksheet
    const sourceSheet = context.workbook.worksheets.getActiveWorksheet();
    sourceSheet.load(["name", "position"]);
    await context.sync();
    Logger.info("Source worksheet loaded:", sourceSheet.name);

    // Get the exact range with data
    const sourceRange = sourceSheet.getRange(params.sourceRange);
    sourceRange.load(["address", "values", "columnCount", "rowCount"]);
    await context.sync();
    
    // Log detailed source data information
    Logger.info("=== Source Data Details ===");
    Logger.info("Sheet Name:", sourceSheet.name);
    Logger.info("Range Address:", sourceRange.address);
    Logger.info("Row Count:", sourceRange.rowCount);
    Logger.info("Column Count:", sourceRange.columnCount);
    Logger.info("Data Values:", JSON.stringify(sourceRange.values, null, 2));

    // Verify we have data
    if (sourceRange.rowCount < 2 || sourceRange.columnCount < 1) {
      throw new Error("Source range must have at least one header row and one data row");
    }

    // Get header row and clean up header names
    const headers = sourceRange.values[0].map(header => header?.toString().trim() || "");
    Logger.info("=== Header Information ===");
    Logger.info("Raw Headers:", sourceRange.values[0]);
    Logger.info("Cleaned Headers:", headers);

    // Validate that all specified fields exist in headers
    const validateFields = (fields, fieldType) => {
      if (!fields) return;
      Logger.info(`\n=== Validating ${fieldType} Fields ===`);
      fields.forEach(field => {
        const fieldIndex = headers.indexOf(field);
        if (fieldIndex === -1) {
          throw new Error(`${fieldType} field "${field}" not found in headers. Available headers: ${headers.join(", ")}`);
        }
        Logger.info(`✓ Found ${fieldType} field "${field}" at column ${fieldIndex + 1}`);
      });
    };

    validateFields(params.rowFields, "Row");
    validateFields(params.columnFields, "Column");
    validateFields(params.dataFields, "Data");

    // Create a new worksheet for the pivot table
    Logger.info("\n=== Creating Pivot Table Worksheet ===");
    const pivotSheet = context.workbook.worksheets.add("PivotTable Analysis");
    pivotSheet.activate();
    await context.sync();
    Logger.info("Created new worksheet:", pivotSheet.name);

    // Add title to the sheet first
    Logger.info("\n=== Adding Title ===");
    const titleCell = pivotSheet.getRange("A1");
    titleCell.values = [["Pivot Table Analysis"]];
    titleCell.format.font.bold = true;
    titleCell.format.font.size = 14;
    await context.sync();
    Logger.info("Title added successfully");

    // Create the pivot table
    Logger.info("\n=== Creating Pivot Table ===");
    
    // Create the pivot table at A3
    const pivotTable = pivotSheet.pivotTables.add(
      "A3",                                    // Use string location instead of Range object
      sourceRange.address,                     // Use string address instead of Range object
      "PivotTable1"
    );
    await context.sync();
    Logger.info("Base pivot table created");

    // Load the hierarchies
    Logger.info("\n=== Loading Hierarchies ===");
    pivotTable.load("hierarchies");
    await context.sync();
    Logger.info("Hierarchies loaded");

    // Add fields in sequence
    Logger.info("\n=== Adding Fields to Pivot Table ===");
    
    // Add row fields
    if (params.rowFields) {
      for (const field of params.rowFields) {
        Logger.info(`Adding row field: ${field}`);
        pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(field));
        await context.sync();
      }
    }

    // Add column fields
    if (params.columnFields) {
      for (const field of params.columnFields) {
        Logger.info(`Adding column field: ${field}`);
        pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(field));
        await context.sync();
      }
    }

    // Add data fields
    if (params.dataFields) {
      for (const field of params.dataFields) {
        Logger.info(`Adding data field: ${field}`);
        const dataHierarchy = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(field));
        if (params.summarizeBy) {
          dataHierarchy.summarizeBy = params.summarizeBy;
        }
        await context.sync();
      }
    }

    // Set layout options
    Logger.info("\n=== Setting Layout Options ===");
    pivotTable.layout.layoutType = "Tabular";
    pivotTable.layout.showRowHeaders = true;
    pivotTable.layout.showColumnHeaders = true;
    pivotTable.layout.enableFieldList = true;
    await context.sync();

    // Refresh the pivot table
    Logger.info("\n=== Refreshing Pivot Table ===");
    pivotTable.refresh();
    await context.sync();

    // Auto-fit columns
    Logger.info("Auto-fitting columns");
    pivotSheet.getUsedRange().format.autofitColumns();
    await context.sync();

    Logger.info("\n=== Pivot Table Creation Complete ===");
    return {
      sheetName: pivotSheet.name,
      success: true
    };
  } catch (error) {
    Logger.error("Create Pivot Table", error, { 
      params,
      errorDetails: {
        message: error.message,
        stack: error.stack,
        code: error.code,
        data: {
          sourceRange: params.sourceRange,
          rowFields: params.rowFields,
          columnFields: params.columnFields,
          dataFields: params.dataFields
        }
      }
    });
    throw error;
  }
}

export async function createChart(context, params) {
  Logger.startOperation("Creating Chart");
  try {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const sourceRange = sheet.getRange(params.sourceRange);
    sourceRange.load("values");
    await context.sync();

    const chart = sheet.charts.add(
      params.chartType || "ColumnClustered",
      sourceRange,
      "Auto"
    );

    chart.title.text = params.title || "Data Analysis";

    if (params.axisLabels) {
      if (params.axisLabels.xAxis) {
        chart.axes.categoryAxis.title.text = params.axisLabels.xAxis;
      }
      if (params.axisLabels.yAxis) {
        chart.axes.valueAxis.title.text = params.axisLabels.yAxis;
      }
    }

    chart.setPosition(params.position?.top || "A1", params.position?.left || null);

    if (params.size) {
      chart.height = params.size.height || 300;
      chart.width = params.size.width || 500;
    }

    await context.sync();
    Logger.info("Chart created successfully with parameters:", params);
  } catch (error) {
    Logger.error("Create Chart", error, { params });
    throw error;
  }
}

export async function formatRange(context, params) {
  Logger.startOperation("Formatting Range");
  try {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(params.range);
    
    if (params.format) {
      applyBasicFormatting(range, params.format);
    }

    if (params.conditionalFormat) {
      applyConditionalFormatting(range, params.conditionalFormat);
    }

    await context.sync();
    Logger.info("Range formatted successfully with parameters:", params);
  } catch (error) {
    Logger.error("Format Range", error, { params });
    throw error;
  }
}

function applyBasicFormatting(range, format) {
  if (format.font) {
    const { bold, italic, size, color, name } = format.font;
    if (bold) range.format.font.bold = bold;
    if (italic) range.format.font.italic = italic;
    if (size) range.format.font.size = size;
    if (color) range.format.font.color = color;
    if (name) range.format.font.name = name;
  }

  if (format.fill?.color) {
    range.format.fill.color = format.fill.color;
  }

  if (format.borders) {
    const edges = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'];
    edges.forEach(edge => {
      if (format.borders[edge.toLowerCase().replace('edge', '')]) {
        range.format.borders.getItem(edge).style = format.borders[edge.toLowerCase().replace('edge', '')];
      }
    });
  }

  if (format.numberFormat) {
    range.numberFormat = format.numberFormat;
  }

  if (format.alignment) {
    const { horizontal, vertical, wrapText } = format.alignment;
    if (horizontal) range.format.horizontalAlignment = horizontal;
    if (vertical) range.format.verticalAlignment = vertical;
    if (wrapText !== undefined) range.format.wrapText = wrapText;
  }
}

function applyConditionalFormatting(range, conditionalFormat) {
  const format = range.conditionalFormats.add(conditionalFormat.type);
  
  switch (conditionalFormat.type) {
    case "ColorScale":
      format.colorScale.criteria = conditionalFormat.criteria;
      break;
    case "DataBar":
      format.dataBar.barColor = conditionalFormat.barColor;
      break;
    case "IconSet":
      format.iconSet.style = conditionalFormat.style;
      break;
  }
}

async function processUploadedFiles(uploadedFiles, processButton) {
  try {
    processButton.disabled = true;
    const modelStatus = document.querySelector(".model-status");
    modelStatus.textContent = "Processing files...";

    // Add a message to the chat indicating file processing
    addMessageToChat('user', 'Process uploaded financial documents');
    const assistantMessage = addMessageToChat('assistant', '<div class="loading">Processing financial documents...</div>');

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Process each file
      for (const filename of uploadedFiles) {
        try {
          const fileInput = document.getElementById('fileInput');
          const file = Array.from(fileInput.files).find(f => f.name === filename);
          
          if (!file) {
            throw new Error(`File ${filename} not found in input`);
          }

          // Read the file using FileReader
          addMessageToChat('assistant', `Reading file ${filename}`);
          const arrayBuffer = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = (e) => reject(new Error(`FileReader error: ${e.target.error}`));
            reader.readAsArrayBuffer(file);
          });

          // Parse the file using SheetJS
          addMessageToChat('assistant', `Parsing file ${filename}`);
          let workbook;
          try {
            workbook = XLSX.read(new Uint8Array(arrayBuffer), { 
              type: 'array',
              cellDates: true,
              cellNF: true,
              cellStyles: true
            });
          } catch (parseError) {
            throw new Error(`Failed to parse file: ${parseError.message}`);
          }

          if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
            throw new Error('No sheets found in workbook');
          }

          // Get the first sheet
          const firstSheetName = workbook.SheetNames[0];
          const firstSheet = workbook.Sheets[firstSheetName];
          
          if (!firstSheet) {
            throw new Error(`Sheet "${firstSheetName}" not found in workbook`);
          }

          // Convert to JSON with error handling
          addMessageToChat('assistant', `Converting sheet data to JSON`);
          let jsonData;
          try {
            jsonData = XLSX.utils.sheet_to_json(firstSheet, { 
              header: 1,
              raw: false,
              dateNF: 'yyyy-mm-dd'
            });
          } catch (jsonError) {
            throw new Error(`Failed to convert sheet to JSON: ${jsonError.message}`);
          }

          if (!jsonData || !jsonData.length) {
            throw new Error('No data found in sheet');
          }

          // Calculate dimensions properly
          const rowCount = jsonData.length;
          const colCount = Math.max(...jsonData.map(row => Array.isArray(row) ? row.length : 0));

          if (rowCount === 0 || colCount === 0) {
            throw new Error('Invalid data dimensions detected');
          }

          // Normalize the data array to ensure consistent dimensions
          const normalizedData = jsonData.map(row => {
            const arrayRow = Array.isArray(row) ? row : [row];
            return arrayRow.concat(Array(colCount - arrayRow.length).fill(""));
          });

          // Create a new worksheet for each file
          addMessageToChat('assistant', `Creating new worksheet for ${filename}`);
          const newSheet = context.workbook.worksheets.add(file.name.split('.')[0]);
          
          // Write the data to the worksheet with validated dimensions
          addMessageToChat('assistant', `Writing ${rowCount} rows and ${colCount} columns of data to worksheet`);
          
          try {
            // Set the values in chunks to handle large datasets better
            const CHUNK_SIZE = 1000;
            for (let startRow = 0; startRow < rowCount; startRow += CHUNK_SIZE) {
              const chunkRows = Math.min(CHUNK_SIZE, rowCount - startRow);
              const range = newSheet.getRangeByIndexes(
                startRow,    // Starting row
                0,          // Starting column
                chunkRows,  // Number of rows
                colCount    // Number of columns
              );
              range.values = normalizedData.slice(startRow, startRow + chunkRows);
              await context.sync();
            }

            // Format the worksheet after all data is written
            const fullRange = newSheet.getRangeByIndexes(0, 0, rowCount, colCount);
            fullRange.format.autofitColumns();
            fullRange.format.autofitRows();

            // Add headers if present (first row)
            if (rowCount > 0) {
              const headerRange = newSheet.getRangeByIndexes(0, 0, 1, colCount);
              headerRange.format.fill.color = "#D3D3D3";
              headerRange.format.font.bold = true;
            }

            await context.sync();
          } catch (writeError) {
            console.error('Error writing data:', writeError);
            throw new Error(`Failed to write data to worksheet: ${writeError.message}\nDimensions: ${rowCount}x${colCount}`);
          }

          addMessageToChat('assistant', `Successfully processed ${filename}`);
        } catch (error) {
          console.error(`Error processing file ${filename}:`, error);
          addMessageToChat('assistant', `Error processing ${filename}: ${error.message}`);
          throw error;
        }
      }

      // Update the assistant message with success
      assistantMessage.querySelector('.message-content').innerHTML = 
        `<div class="status-message success">Successfully processed ${uploadedFiles.size} file(s)!</div>`;

      modelStatus.textContent = "Ready";
    });
  } catch (error) {
    console.error("Error processing files:", error);
    assistantMessage.querySelector('.message-content').innerHTML = 
      `<div class="status-message error">Error processing files: ${error.message}</div>`;
  } finally {
    processButton.disabled = false;
  }
}

async function tryImplementation(implementationCode, statusArea) {
  // Basic security validation
  const forbiddenPatterns = [
    "eval\\(",
    "Function\\(",
    "setTimeout\\(",
    "setInterval\\(",
    "new\\s+Function",
    "document\\.write",
    "<script",
    "window\\.",
    "localStorage",
    "sessionStorage",
    "indexedDB",
    "fetch\\("
  ];

  const securityRegex = new RegExp(forbiddenPatterns.join("|"), "i");
  if (securityRegex.test(implementationCode)) {
    throw new Error("Implementation contains potentially unsafe code");
  }

  // Create the function from the code string
  let executeFunction;
  try {
    executeFunction = new Function('return ' + implementationCode)();
  } catch (error) {
    console.error("Error creating function:", error);
    throw new Error("Failed to parse implementation code: " + error.message);
  }

  // Validate the function is actually async
  if (!(executeFunction.constructor.name === "AsyncFunction")) {
    throw new Error("Implementation must be an async function");
  }

  try {
    await Excel.run(async (context) => {
      console.log("Executing implementation code");
      statusArea.textContent = "Executing changes...";
      
      // Start undo batch to make changes atomic
      context.application.suspendScreenUpdatingUntilNextSync();
      
      await executeFunction(context);
      await context.sync();
      statusArea.textContent = "Changes implemented successfully!";
      statusArea.style.color = "green";
    });
    return { success: true };
  } catch (error) {
    console.error("Error during execution:", error);
    statusArea.textContent = "Attempt failed: " + error.message;
    statusArea.style.color = "orange";
    return { success: false, error };
  }
}

export { processUploadedFiles, tryImplementation }; 