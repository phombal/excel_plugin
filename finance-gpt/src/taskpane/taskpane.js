/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import { setupFileUpload } from '../ui/fileUpload.js';
import { getAIResponse } from '../services/ai-service.js';
import { addMessageToChat, formatResponse, createStatusMessage } from '../ui/chat.js';
import { tryImplementation } from '../excel/operations.js';

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
      document.getElementById("clearChat").onclick = clearChat;
      setupFileUpload();
      
      // Auto-resize textarea
      const textarea = document.getElementById("queryInput");
      textarea.addEventListener("input", function() {
        this.style.height = "auto";
        this.style.height = (this.scrollHeight) + "px";
      });
      
      console.log("UI initialized successfully");
    } else {
      console.error("Required elements not found:", {
        sideloadMsg: !!sideloadMsg,
        appBody: !!appBody
      });
    }
  }
});

function clearChat() {
  const chatHistory = document.getElementById("chatHistory");
  chatHistory.innerHTML = "";
  addMessageToChat('assistant', 'Chat history cleared. How can I help you?');
}

async function callClaude(data) {
  try {
    const CLAUDE_API_KEY = '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';
    
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': CLAUDE_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-3-sonnet-20240229',
        max_tokens: 4096,
        messages: [
          {
            role: 'user',
            content: `You are an Excel expert assistant. Analyze this spreadsheet data and answer the following query:
            
            Spreadsheet Data:
            ${JSON.stringify(data.data, null, 2)}
            
            Range: ${data.range}
            Row Count: ${data.sheetMetadata.rowCount}
            Column Count: ${data.sheetMetadata.columnCount}
            Has Headers: ${data.sheetMetadata.hasHeaders}
            
            User Query: ${data.query}
            
            Please provide a detailed response with explanations and any relevant Excel formulas or operations.`
          }
        ]
      })
    });

    if (!response.ok) {
      throw new Error(`Claude API error: ${response.status} ${response.statusText}`);
    }

    const result = await response.json();
    return result.content[0].text;
  } catch (error) {
    console.error('Error calling Claude API:', error);
    throw new Error('Failed to get response from Claude. Please try again or switch to GPT-4.');
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
      
      // Prepare the data for AI service
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
          console.log(`Making 5 AI API calls (Attempt ${attempts + 1})...`);
          
          responses = await Promise.all([
            getAIResponse(spreadsheetData),
            getAIResponse(spreadsheetData),
            getAIResponse(spreadsheetData),
            getAIResponse(spreadsheetData),
            getAIResponse(spreadsheetData)
          ]);
          
          console.log("Received all AI API responses");
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
      const chatHistory = document.getElementById("chatHistory");
      chatHistory.scrollTop = chatHistory.scrollHeight;
    });
  } catch (error) {
    console.error("Error in handleQuery:", error);
    const modelStatus = document.querySelector(".model-status");
    const sendButton = document.getElementById("submitQuery");
    
    // Add error message to the last assistant message
    const lastAssistantMessage = document.querySelector('.assistant-message:last-child');
    if (lastAssistantMessage) {
      lastAssistantMessage.querySelector('.message-content').innerHTML = 
        `<div class="status-message error">Error: ${error.message}</div>`;
    }
    
    modelStatus.textContent = "Ready";
    modelStatus.classList.remove("loading");
    sendButton.disabled = false;
  }
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