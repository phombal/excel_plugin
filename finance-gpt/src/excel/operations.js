import { Logger } from '../utils/logger';

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