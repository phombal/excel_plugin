/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  console.log("Office.onReady called", { host: info.host });
  if (info.host === Office.HostType.Excel) {
    console.log("Excel detected, setting up event handlers");
    
    // Get references to elements
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    
    if (sideloadMsg && appBody) {
      // Hide loading message and show app
      sideloadMsg.style.display = "none";
      appBody.style.display = "flex";
      
      // Set up event handlers
      document.getElementById("submitQuery").onclick = handleQuery;
      document.getElementById("implementSuggestion").onclick = handleImplementation;
      
      console.log("UI initialized successfully");
    } else {
      console.error("Required elements not found:", {
        sideloadMsg: !!sideloadMsg,
        appBody: !!appBody
      });
    }
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

async function handleQuery() {
  console.log("handleQuery started");
  try {
    const queryInput = document.getElementById("queryInput").value;
    if (!queryInput.trim()) {
      throw new Error("Please enter a query first");
    }
    
    console.log("Query input:", queryInput);
    const responseArea = document.getElementById("responseArea");
    const implementButton = document.getElementById("implementSuggestion");
    
    responseArea.innerHTML = "Analyzing...";
    implementButton.style.display = "none";
    
    await Excel.run(async (context) => {
      console.log("Excel.run started");
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load("values, address, rowCount, columnCount");
      
      await context.sync();
      
      if (!usedRange || !usedRange.values || usedRange.values.length === 0) {
        throw new Error("No data found in the active worksheet");
      }
      
      console.log("Got spreadsheet data:", {
        address: usedRange.address,
        rowCount: usedRange.values.length,
        colCount: usedRange.values[0].length
      });
      
      // Prepare the data for OpenAI
      const spreadsheetData = {
        data: usedRange.values,
        range: usedRange.address,
        query: queryInput,
        sheetMetadata: {
          rowCount: usedRange.values.length,
          columnCount: usedRange.values[0].length,
          hasHeaders: true // Assuming first row contains headers
        }
      };

      console.log("Calling OpenAI API...");
      const response = await callOpenAI(spreadsheetData);
      console.log("OpenAI API response received");
      
      responseArea.innerHTML = response;
      
      // Show implement button if the response contains executable code
      if (response.includes("IMPLEMENT:") && response.includes("```javascript")) {
        console.log("Implementation code detected, showing button");
        implementButton.style.display = "block";
      }
    });
  } catch (error) {
    console.error("Error in handleQuery:", error);
    document.getElementById("responseArea").innerHTML = "Error: " + error.message;
  }
}

async function handleImplementation() {
  console.log("handleImplementation started");
  try {
    const response = document.getElementById("responseArea").innerHTML;
    const implementationMatch = response.match(/IMPLEMENT:\s*```javascript\s*([\s\S]*?)\s*```/);
    
    if (!implementationMatch) {
      throw new Error("No valid implementation code found in the response");
    }

    let implementationCode = implementationMatch[1].trim();
    console.log("Found implementation code");

    // Validate the code contains a valid async function
    if (!implementationCode.includes("async function") || !implementationCode.includes("context")) {
      throw new Error("Invalid implementation code format");
    }

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

    const statusArea = document.createElement("div");
    statusArea.className = "status-message";
    document.getElementById("responseArea").appendChild(statusArea);

    try {
      await Excel.run(async (context) => {
        console.log("Executing implementation code");
        statusArea.textContent = "Executing changes...";
        
        // Start undo batch to make changes atomic
        context.application.suspendScreenUpdatingUntilNextSync();
        
        try {
          await executeFunction(context);
          await context.sync();
          statusArea.textContent = "Changes implemented successfully!";
          statusArea.style.color = "green";
        } catch (error) {
          console.error("Error during execution:", error);
          throw error;
        }
      });
    } catch (error) {
      statusArea.textContent = "Error executing changes: " + error.message;
      statusArea.style.color = "red";
      throw error;
    }

  } catch (error) {
    console.error("Error in handleImplementation:", error);
    const errorMessage = "Error: " + error.message;
    document.getElementById("responseArea").innerHTML += "\n\n" + errorMessage;
  }
}

async function callOpenAI(data) {
  console.log("callOpenAI started");
  const OPENAI_API_KEY = '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';
  
  try {
    console.log("Making OpenAI API request...");
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${OPENAI_API_KEY}`
      },
      body: JSON.stringify({
        model: "gpt-4",
        messages: [{
          role: "system",
          content: `You are a financial analysis assistant. Analyze the provided Excel data and respond to queries. 
          If the user's query involves any edits to the excel sheet, generate Office.js code that solves their request.
          
          Rules for generating Office.js code:
          1. Always wrap the code in an async function that takes a 'context' parameter
          2. Use proper error handling with try/catch blocks
          3. Always include context.sync() calls where necessary
          4. Use proper Office.js API patterns and best practices
          6. Return meaningful error messages if operations fail
          7. Always include error handling
          8. Validate inputs and ranges before operations
          9. MOST IMPORTANTLY: ALWAYS ENSURE THAT THE CODE IS EXECUTABLE AND FREE OF ANY SYNTAX AND RUNTIME ERRORS
          
          Format your response as follows for modifications:
          IMPLEMENT:
          \`\`\`javascript
          async function executeChanges(context) {
            try {
              // Your Office.js code here
              await context.sync();
            } catch (error) {
              throw new Error("Failed to execute changes: " + error.message);
            }
          }
          \`\`\`
          
          For analysis questions without modifications, provide a direct answer.`
        }, {
          role: "user",
          content: JSON.stringify(data)
        }],
        temperature: 0.7
      })
    });

    if (!response.ok) {
      throw new Error(`API request failed with status ${response.status}`);
    }

    const result = await response.json();
    console.log("OpenAI API response received");
    
    if (!result.choices || !result.choices[0]?.message?.content) {
      throw new Error("Invalid response format from OpenAI API");
    }
    
    return result.choices[0].message.content;
  } catch (error) {
    console.error("Error in callOpenAI:", error);
    throw new Error("Failed to get AI response: " + error.message);
  }
}
