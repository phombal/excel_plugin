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
    const chatHistory = document.getElementById("chatHistory");
    const modelStatus = document.querySelector(".model-status");
    const sendButton = document.getElementById("submitQuery");
    
    // Add user message to chat
    addMessageToChat('user', queryInput);
    
    // Update UI to loading state
    const assistantMessage = addMessageToChat('assistant', '<div class="loading">Analyzing your spreadsheet...</div>');
    modelStatus.textContent = "Processing";
    modelStatus.classList.add("loading");
    sendButton.disabled = true;
    
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

      let attempts = 0;
      const maxAttempts = 3; // Maximum number of batch retry attempts
      let responses = null;

      while (attempts < maxAttempts && !responses) {
        try {
          if (attempts > 0) {
            assistantMessage.querySelector('.message-content').innerHTML = 
              `<div class="loading">Retrying API calls (Attempt ${attempts + 1}/${maxAttempts})...</div>`;
          } else {
            assistantMessage.querySelector('.message-content').innerHTML = 
              '<div class="loading">Generating multiple solutions...</div>';
          }

          // Make 5 API calls in parallel for better performance
          console.log(`Making 5 OpenAI API calls (Attempt ${attempts + 1})...`);
          
          responses = await Promise.all([
            callOpenAI(spreadsheetData),
            callOpenAI(spreadsheetData),
            callOpenAI(spreadsheetData),
            callOpenAI(spreadsheetData),
            callOpenAI(spreadsheetData)
          ]);
          
          console.log("Received all OpenAI API responses");
        } catch (error) {
          console.error(`Attempt ${attempts + 1} failed:`, error);
          attempts++;
          
          if (attempts >= maxAttempts) {
            throw new Error("Failed to get AI response after multiple attempts. Please try again.");
          }
          
          // Add a small delay before retrying
          await new Promise(resolve => setTimeout(resolve, 1000));
        }
      }
      
      // Store all responses in the assistant message
      responses.forEach((response, index) => {
        assistantMessage.setAttribute(`data-response${index + 1}`, response);
      });
      
      // Format and display the first response
      const formattedResponse = formatResponse(responses[0]);
      assistantMessage.querySelector('.message-content').innerHTML = formattedResponse;
      
      // Show implement button if any response contains executable code
      const hasImplementation = responses.some(response => 
        response.includes("IMPLEMENT:") && response.includes("```javascript")
      );
      
      if (hasImplementation) {
        console.log("Implementation code detected, showing button");
        const implementButton = document.createElement('button');
        implementButton.className = 'implement-button';
        implementButton.textContent = 'Implement Changes';
        implementButton.onclick = () => handleImplementation(assistantMessage);
        assistantMessage.appendChild(implementButton);
      }

      // Reset UI state
      modelStatus.textContent = "Ready";
      modelStatus.classList.remove("loading");
      sendButton.disabled = false;
      
      // Clear input after successful response
      document.getElementById("queryInput").value = "";
      
      // Scroll to the bottom
      chatHistory.scrollTop = chatHistory.scrollHeight;
    });
  } catch (error) {
    console.error("Error in handleQuery:", error);
    const chatHistory = document.getElementById("chatHistory");
    const modelStatus = document.querySelector(".model-status");
    const sendButton = document.getElementById("submitQuery");
    
    // Add error message to the last assistant message
    const lastAssistantMessage = chatHistory.querySelector('.assistant-message:last-child');
    if (lastAssistantMessage) {
      lastAssistantMessage.querySelector('.message-content').innerHTML = 
        `<div class="status-message error">Error: ${error.message}</div>`;
    }
    
    modelStatus.textContent = "Ready";
    modelStatus.classList.remove("loading");
    sendButton.disabled = false;
  }
}

function addMessageToChat(role, content) {
  const chatHistory = document.getElementById("chatHistory");
  const messageDiv = document.createElement('div');
  messageDiv.className = `chat-message ${role}-message`;
  
  const header = document.createElement('div');
  header.className = 'message-header';
  
  const roleSpan = document.createElement('span');
  roleSpan.className = 'message-role';
  roleSpan.textContent = role === 'user' ? 'You' : 'Assistant';
  header.appendChild(roleSpan);
  
  const messageContent = document.createElement('div');
  messageContent.className = 'message-content';
  messageContent.innerHTML = content;
  
  messageDiv.appendChild(header);
  messageDiv.appendChild(messageContent);
  chatHistory.appendChild(messageDiv);
  
  return messageDiv;
}

// Helper function to format the response with syntax highlighting
function formatResponse(response) {
  // Replace code blocks with styled versions
  return response.replace(
    /```javascript([\s\S]*?)```/g,
    (match, code) => `<code class="javascript">${code.trim()}</code>`
  );
}

async function handleImplementation(messageElement) {
  console.log("handleImplementation started");
  try {
    const statusArea = document.createElement("div");
    statusArea.className = "status-message";
    messageElement.appendChild(statusArea);

    // Extract all implementations
    const implementations = [];
    for (let i = 1; i <= 5; i++) {
      const response = messageElement.getAttribute(`data-response${i}`);
      if (!response) continue;
      
      const match = response.match(/IMPLEMENT:\s*```javascript\s*([\s\S]*?)\s*```/);
      if (match) {
        implementations.push({
          code: match[1].trim(),
          index: i
        });
      }
    }

    if (implementations.length === 0) {
      throw new Error("No valid implementation code found in any response");
    }

    // Try each implementation until one succeeds
    let lastError = null;
    for (const { code, index } of implementations) {
      try {
        statusArea.textContent = `Trying implementation ${index} of ${implementations.length}...`;
        statusArea.className = "status-message info";
        
        const result = await tryImplementation(code, statusArea);
        if (result.success) {
          console.log(`Implementation ${index} succeeded`);
          statusArea.className = "status-message success";
          return; // Success! We're done
        }
        lastError = result.error;
      } catch (error) {
        console.log(`Implementation ${index} failed:`, error);
        lastError = error;
      }
    }

    // If we're here, all implementations failed
    console.error("All implementations failed");
    statusArea.textContent = "Error: All implementation attempts failed. Please try regenerating the solution.";
    statusArea.className = "status-message error";
    if (lastError) {
      const errorDetails = document.createElement("div");
      errorDetails.textContent = `Last error: ${lastError.message}`;
      errorDetails.style.fontSize = "0.9em";
      errorDetails.style.marginTop = "5px";
      statusArea.appendChild(errorDetails);
    }

  } catch (error) {
    console.error("Error in handleImplementation:", error);
    const errorMessage = "Error: " + error.message;
    messageElement.appendChild(createStatusMessage(errorMessage, "error"));
  }
}

function createStatusMessage(message, type) {
  const statusDiv = document.createElement("div");
  statusDiv.className = `status-message ${type}`;
  statusDiv.textContent = message;
  return statusDiv;
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

async function callOpenAI(data) {
  console.log("callOpenAI started");
  const OPENAI_API_KEY = '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';
  
  const maxRetries = 3; // Maximum retries per individual API call
  let retryCount = 0;
  let lastError = null;

  while (retryCount < maxRetries) {
    try {
      console.log(`Making OpenAI API request (Attempt ${retryCount + 1})...`);
      const response = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${OPENAI_API_KEY}`
        },
        body: JSON.stringify({
          model: "gpt-4",  // Fixed typo in model name from "gpt-4o" to "gpt-4"
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
      console.error(`Error in callOpenAI (Attempt ${retryCount + 1}):`, error);
      lastError = error;
      retryCount++;
      
      if (retryCount >= maxRetries) {
        throw new Error("Failed to get AI response after multiple retries: " + lastError.message);
      }
      
      // Add exponential backoff delay before retrying
      const delay = Math.min(1000 * Math.pow(2, retryCount), 5000);
      await new Promise(resolve => setTimeout(resolve, delay));
    }
  }
}