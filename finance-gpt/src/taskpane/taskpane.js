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
    console.log("Query input:", queryInput);
    const responseArea = document.getElementById("responseArea");
    
    responseArea.innerHTML = "Analyzing...";
    
    await Excel.run(async (context) => {
      console.log("Excel.run started");
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load("values, address");
      
      await context.sync();
      console.log("Got spreadsheet data:", {
        address: usedRange.address,
        rowCount: usedRange.values?.length,
        colCount: usedRange.values?.[0]?.length
      });
      
      // Prepare the data for OpenAI
      const spreadsheetData = {
        data: usedRange.values,
        range: usedRange.address,
        query: queryInput
      };

      console.log("Calling OpenAI API...");
      // Call OpenAI API
      const response = await callOpenAI(spreadsheetData);
      console.log("OpenAI API response received:", response);
      responseArea.innerHTML = response;
      
      // Show implement button if the response contains actionable suggestions
      if (response.includes("IMPLEMENT:")) {
        console.log("Implementation suggestion detected, showing button");
        document.getElementById("implementSuggestion").style.display = "block";
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
    const implementationPart = response.split("IMPLEMENT:")[1].trim();
    console.log("Implementation part:", implementationPart);
    
    await Excel.run(async (context) => {
      console.log("Excel.run started for implementation");
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Parse and implement the AI suggestion
      // This will need to be implemented based on the specific format
      // of AI responses and desired functionality
      
      await context.sync();
      document.getElementById("responseArea").innerHTML += "\n\nImplementation complete!";
    });
  } catch (error) {
    console.error("Error in handleImplementation:", error);
    document.getElementById("responseArea").innerHTML += "\n\nError implementing suggestion: " + error.message;
  }
}

async function callOpenAI(data) {
  console.log("callOpenAI started");
  const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
  
  if (!OPENAI_API_KEY) {
    console.error("OpenAI API key not found in environment");
    throw new Error("OpenAI API key not configured. Please set the OPENAI_API_KEY environment variable.");
  }
  
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
          content: "You are a financial analysis assistant. Analyze the provided Excel data and respond to queries."
        }, {
          role: "user",
          content: JSON.stringify(data)
        }],
        temperature: 0.7
      })
    });

    console.log("OpenAI API response status:", response.status);
    const result = await response.json();
    console.log("OpenAI API result:", result);
    
    if (!result.choices || !result.choices[0]?.message?.content) {
      console.error("Unexpected API response format:", result);
      throw new Error("Unexpected response from OpenAI API");
    }
    
    return result.choices[0].message.content;
  } catch (error) {
    console.error("Error in callOpenAI:", error);
    throw error;
  }
}
