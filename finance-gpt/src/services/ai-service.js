// System prompt for both models
const SYSTEM_PROMPT = `You are a financial analysis assistant. Analyze the provided Excel data and respond to queries. 
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

For analysis questions without modifications, provide a direct answer.`;

async function callOpenAI(data) {
  console.log("callOpenAI started");
  const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
  
  const maxRetries = 3;
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
          model: "gpt-4",
          messages: [{
            role: "system",
            content: SYSTEM_PROMPT
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
      
      const delay = Math.min(1000 * Math.pow(2, retryCount), 5000);
      await new Promise(resolve => setTimeout(resolve, delay));
    }
  }
}

async function callClaude(data) {
  console.log("callClaude started");
  const CLAUDE_API_KEY = process.env.CLAUDE_API_KEY;
  
  const maxRetries = 3;
  let retryCount = 0;
  let lastError = null;

  while (retryCount < maxRetries) {
    try {
      console.log(`Making Claude API request (Attempt ${retryCount + 1})...`);
      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': CLAUDE_API_KEY,
          'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify({
          model: "claude-3-5-sonnet-20241022",
          max_tokens: 4096,
          temperature: 0,
          system: SYSTEM_PROMPT,
          messages: [{
            role: "user",
            content: [
              {
                type: "text",
                text: `Here is my query and spreadsheet data to analyze:\nQuery: ${data.query}\n\nSpreadsheet Data:\n${JSON.stringify(data.data, null, 2)}\n\nRange: ${data.range}\n\nSheet Metadata:\n${JSON.stringify(data.sheetMetadata, null, 2)}`
              }
            ]
          }]
        })
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => null);
        throw new Error(
          `API request failed with status ${response.status}: ${
            errorData ? JSON.stringify(errorData) : 'No error details available'
          }`
        );
      }

      const result = await response.json();
      console.log("Claude API response received");
      
      if (!result.content || !result.content[0]?.text) {
        throw new Error("Invalid response format from Claude API");
      }
      
      return result.content[0].text;
    } catch (error) {
      console.error(`Error in callClaude (Attempt ${retryCount + 1}):`, error);
      lastError = error;
      retryCount++;
      
      if (retryCount >= maxRetries) {
        throw new Error("Failed to get AI response after multiple retries: " + lastError.message);
      }
      
      const delay = Math.min(1000 * Math.pow(2, retryCount), 5000);
      await new Promise(resolve => setTimeout(resolve, delay));
    }
  }
}

async function getAIResponse(data) {
  const modelSelect = document.getElementById("modelSelect");
  const selectedModel = modelSelect.value;
  
  if (selectedModel === "gpt4") {
    return callOpenAI(data);
  } else {
    return callClaude(data);
  }
}

export { getAIResponse }; 