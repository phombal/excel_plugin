/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows the task pane when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function showTaskPane(event) {
  Office.addin.showAsTaskpane();
  event.completed();
}

/**
 * Gets the current worksheet data and sends it to OpenAI for analysis
 * @param {string} userQuery - The user's question about the spreadsheet
 * @returns {Promise<string>} - The AI response
 */
async function analyzeSpreadsheet(userQuery) {
  try {
    // Get the current worksheet
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load("values, address");
      
      await context.sync();
      
      // Prepare the data for OpenAI
      const spreadsheetData = {
        data: usedRange.values,
        range: usedRange.address,
        query: userQuery
      };

      // Call OpenAI API (implementation needed)
      const response = await callOpenAI(spreadsheetData);
      return response;
    });
  } catch (error) {
    console.error("Error analyzing spreadsheet:", error);
    throw error;
  }
}

/**
 * Implements financial formulas or modifications suggested by AI
 * @param {string} aiSuggestion - The AI-generated formula or modification
 */
async function implementFinancialModel(aiSuggestion) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Parse and implement the AI suggestion
      // This will need to be implemented based on the specific format
      // of AI responses and desired functionality
      
      await context.sync();
    });
  } catch (error) {
    console.error("Error implementing financial model:", error);
    throw error;
  }
}

/**
 * Calls the OpenAI API with spreadsheet data
 * @param {Object} data - The spreadsheet data and query
 * @returns {Promise<string>} - The AI response
 */
async function callOpenAI(data) {
  // Get API key from environment variable or secure configuration
  const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
  
  if (!OPENAI_API_KEY) {
    throw new Error("OpenAI API key not configured. Please set the OPENAI_API_KEY environment variable.");
  }
  
  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${OPENAI_API_KEY}`
      },
      body: JSON.stringify({
        model: "gpt-4o",
        messages: [{
          role: "system",
          content: "You are a financial analysis assistant. Analyze the provided Excel data and respond to queries."
        }, {
          role: "user",
          content: JSON.stringify(data)
        }],
        temperature: 0.7
      })
    });

    const result = await response.json();
    return result.choices[0].message.content;
  } catch (error) {
    console.error("Error calling OpenAI:", error);
    throw error;
  }
}

// Register the functions with Office
Office.actions.associate("showTaskPane", showTaskPane);
