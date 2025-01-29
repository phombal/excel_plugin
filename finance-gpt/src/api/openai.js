import { Logger } from '../utils/logger';

export async function callOpenAI(data) {
  Logger.startOperation("OpenAI API Call");
  const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
  
  if (!OPENAI_API_KEY) {
    Logger.error("OpenAI API", new Error("API key not configured"));
    throw new Error("OpenAI API key not configured. Please set the OPENAI_API_KEY environment variable.");
  }
  
  try {
    Logger.info("Request Details:", {
      model: "gpt-4o-mini",
      dataShape: {
        rowCount: data.data.length,
        columnCount: data.data[0].length,
        headers: data.headers
      },
      query: data.query
    });

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
          content: getSystemPrompt()
        }, {
          role: "user",
          content: JSON.stringify(data)
        }],
        temperature: 0.7
      })
    });

    Logger.info("API Response Status:", response.status);
    const result = await response.json();
    
    if (!result.choices || !result.choices[0]?.message?.content) {
      throw new Error("Unexpected response from OpenAI API");
    }
    
    Logger.endOperation("OpenAI API Call");
    return result.choices[0].message.content;
  } catch (error) {
    Logger.error("OpenAI API Call", error);
    throw error;
  }
}

function getSystemPrompt() {
  return `You are an Excel expert assistant that helps users analyze and manipulate spreadsheet data.
  When users ask you to perform Excel operations like creating pivot tables, charts, or formatting:
  1. First explain what you're going to do
  2. Then provide the commands in the following JSON format:
  EXCEL_COMMAND:
  [
    {
      "type": "CREATE_PIVOT_TABLE",
      "params": {
        "sourceRange": "A1:D10",  // Must be the exact range containing the data including headers
        "rowFields": ["Category"],  // Must EXACTLY match column headers - no extra spaces!
        "columnFields": ["Year"],   // Must EXACTLY match column headers - no extra spaces!
        "dataFields": ["Sales"],    // Must EXACTLY match column headers - no extra spaces!
        "summarizeBy": "Sum"        // Optional: Sum, Count, Average, etc.
      }
    }
  ]
  END_COMMAND
  
  CRITICAL REQUIREMENTS:
  - Field names must EXACTLY match the column headers in the spreadsheet - no trailing or leading spaces
  - Source range must include the header row
  - Always analyze the data structure first and use the exact column names as shown in the headers
  - Double check that field names have no extra spaces
  - For pivot tables, prefer using 1-2 row fields, 0-1 column fields, and 1 data field
  
  Available command types:
  - CREATE_PIVOT_TABLE
  - CREATE_CHART
  - FORMAT_RANGE
  
  For non-command queries, provide clear and concise analysis of the data.`;
} 