// --- Global Variables ---
const ADMIN_EMAIL_FOR_ERRORS = "abela.j@gardenschool.edu.my"; // Set to the admin's email for error notifications, or "" if not needed.

// --- Column Names (MUST match your Sheet headers EXACTLY) ---
const STATUS_COLUMN_NAME = "Status";
const ASK_AI_COLUMN_NAME = "Ask AI"; // MUST match the Form question/Sheet Header
const EMAIL_COLUMN_NAME = "Email address"; // MUST match the Form question/Sheet Header
const AI_RESPONSE_COLUMN_NAME = "AI Response";
const ERROR_DETAILS_COLUMN_NAME = "Error Details"; // Optional
const STUDENT_NAME_COLUMN = 2; //  "Your name" column - Double-check correct index

// --- Settings Sheet and Master Prompt ---
const SETTINGS_SHEET_NAME = "Settings"; // Name of the sheet where the Master Prompt is stored
const MASTER_PROMPT_CELL = "B1"; // Cell where the master prompt is stored (e.g., B1)
const MASTER_PROMPT_LABEL_CELL = "A1"; // Cell with the "Master Prompt" label (e.g., A1)

// Function to get API key securely
function getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    Logger.log("FATAL ERROR: GEMINI_API_KEY script property not set.");
    if (ADMIN_EMAIL_FOR_ERRORS) {
      MailApp.sendEmail(ADMIN_EMAIL_FOR_ERRORS, "CRITICAL New Generator Script Error", "Gemini API Key is missing in Script Properties.");
    }
  }
  return apiKey;
}

// --- Triggered on Form Submission ---
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const newRow = e.range.getRow(); // Row of the new submission
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // --- Get Column Indices (Error checking for missing headers) ---
  const statusColIndex = headers.indexOf(STATUS_COLUMN_NAME) + 1;
  const askAiColIndex = headers.indexOf(ASK_AI_COLUMN_NAME) + 1;
  const emailColIndex = headers.indexOf(EMAIL_COLUMN_NAME) + 1;
  const aiResponseColIndex = headers.indexOf(AI_RESPONSE_COLUMN_NAME) + 1;
  const errorDetailsColIndex = headers.indexOf(ERROR_DETAILS_COLUMN_NAME) + 1; // Optional

  if (statusColIndex <= 0 || askAiColIndex <= 0 || emailColIndex <= 0 || aiResponseColIndex <= 0) {
    Logger.log("Error: One or more required columns ('Status', 'Ask AI', 'Email address', 'AI Response') not found.");
    if (ADMIN_EMAIL_FOR_ERRORS) MailApp.sendEmail(ADMIN_EMAIL_FOR_ERRORS, "New Generator Config Error", "Check column names in the Sheet and Script.");
    return; // Exit if required columns are missing
  }

  // --- 1. Set Initial Status ---
  sheet.getRange(newRow, statusColIndex).setValue("Generating"); // Indicate processing has started

  // --- 2. Get Data from Submission ---
  const askAiPrompt = sheet.getRange(newRow, askAiColIndex).getValue();
  const studentEmail = sheet.getRange(newRow, emailColIndex).getValue();

  // --- Validate Email ---
  if (!studentEmail || !studentEmail.includes('@')) {
    Logger.log(`Invalid student email found in row ${newRow}: ${studentEmail}`);
    sheet.getRange(newRow, statusColIndex).setValue("Error - Invalid Email");
    if (errorDetailsColIndex > 0) sheet.getRange(newRow, errorDetailsColIndex).setValue("Invalid student email format.");
    if (ADMIN_EMAIL_FOR_ERRORS) MailApp.sendEmail(ADMIN_EMAIL_FOR_ERRORS, "New Generator - Invalid Email", `Invalid email address "${studentEmail}" for prompt: "${askAiPrompt}" in row ${newRow}.`);
    return;
  }

  // --- 3. Call Gemini API ---
  const apiKey = getApiKey();
  if (!apiKey) {
    sheet.getRange(newRow, statusColIndex).setValue("Error - API Key Missing");
    if (errorDetailsColIndex > 0) sheet.getRange(newRow, errorDetailsColIndex).setValue("API Key configuration missing.");
    // The getApiKey function should already notify the admin.
    return;
  }

  try {
    const aiResponse = callGeminiApi(askAiPrompt, apiKey);

    // --- 4. Process AI Response ---
    if (aiResponse && aiResponse.text) {
      const aiStory = aiResponse.text;
      sheet.getRange(newRow, aiResponseColIndex).setValue(aiStory); // Log the AI's response
      sheet.getRange(newRow, statusColIndex).setValue("Generated");  // Update the status to generated (not sent yet)
      Logger.log(`AI response generated successfully for row ${newRow}`);
    } else {
      // Handle AI API errors
      const errorMessage = (aiResponse && aiResponse.error) ? aiResponse.error : "Unknown error generating response.";
      Logger.log(`AI Error for row ${newRow}: ${errorMessage}`);
      sheet.getRange(newRow, statusColIndex).setValue("Error - AI Failed");
      if (aiResponseColIndex > 0) sheet.getRange(newRow, aiResponseColIndex).setValue("Error - See details");
      if (errorDetailsColIndex > 0) sheet.getRange(newRow, errorDetailsColIndex).setValue(errorMessage);
      if (ADMIN_EMAIL_FOR_ERRORS) MailApp.sendEmail(ADMIN_EMAIL_FOR_ERRORS, "New Generator - AI Error", `Failed to generate response for prompt: "${askAiPrompt}" (Row ${newRow}). Error: ${errorMessage}`);
    }

  } catch (error) {
    // Handle script errors
    Logger.log(`Script Error processing row ${newRow}: ${error}`);
    sheet.getRange(newRow, statusColIndex).setValue("Error - Script Failed");
    if (errorDetailsColIndex > 0) sheet.getRange(newRow, errorDetailsColIndex).setValue(`Script Error: ${error.message}`);
    if (ADMIN_EMAIL_FOR_ERRORS) MailApp.sendEmail(ADMIN_EMAIL_FOR_ERRORS, "New Generator - Script Error", `Script failed while processing prompt: "${askAiPrompt}" (Row ${newRow}). Error: ${error}`);
  }
}

// --- Function to call the Gemini API ---
function callGeminiApi(promptText, apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`; // Use appropriate model

  // --- Get the Master Prompt from the "Settings" Sheet ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME); // Replace "Settings" if you used a different sheet name
  if (!settingsSheet) {
    Logger.log("Error: 'Settings' sheet not found. Create a sheet named 'Settings' and add the Master Prompt.");
    return { error: "Settings sheet not found." };  // Handle the error
  }
  const masterPromptCell = settingsSheet.getRange(MASTER_PROMPT_CELL); // Assuming Master Prompt is in B1. Change as needed.
  const masterPrompt = masterPromptCell.getValue();

  if (!masterPrompt) {
    Logger.log("Error: Master prompt not found in Settings sheet.");
    return { error: "Master prompt not found." };  // Handle the error
  }

  // --- Customize the Prompt for Gemini, Incorporating the Master Prompt ---
  const fullPrompt = `${masterPrompt} Student Prompt: "${promptText}"`;  // Combine Master Prompt and student's prompt

  const requestBody = {
    "contents": [{
      "parts": [{
        "text": fullPrompt
      }]
    }],
    // --- Recommended Safety Settings --- (Adjust as needed for your use case)
    "safetySettings": [
      { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_LOW_AND_ABOVE" },
      { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_LOW_AND_ABOVE" },
      { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_LOW_AND_ABOVE" },
      { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_LOW_AND_ABOVE" }
    ],
    // --- Optional: Generation Configuration ---
    "generationConfig": {
      "temperature": 0.7, // Adjust temperature for creativity
      "maxOutputTokens": 512 // Adjust token limit
      // "topK": ...,
      // "topP": ...
    }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(requestBody),
    'muteHttpExceptions': true // Important for error handling
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      if (jsonResponse.candidates && jsonResponse.candidates.length > 0 &&
          jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts &&
          jsonResponse.candidates[0].content.parts.length > 0 && jsonResponse.candidates[0].content.parts[0].text) {
        return { text: jsonResponse.candidates[0].content.parts[0].text.trim() }; // Extract the text from the response
      } else if (jsonResponse.promptFeedback && jsonResponse.promptFeedback.blockReason) {
        const blockReason = jsonResponse.promptFeedback.blockReason;
        Logger.log(`Prompt blocked. Reason: ${blockReason}`);
        return { error: `Prompt blocked due to safety settings (${blockReason})` };
      }
      else {
        Logger.log("Unexpected AI response structure: " + responseBody);
        return { error: "Unexpected response structure from AI." };
      }
    } else {
      Logger.log(`Gemini API Error - Code: ${responseCode}, Body: ${responseBody}`);
      return { error: `API request failed with status code ${responseCode}. Details: ${responseBody}` };
    }
  } catch (error) {
    Logger.log("Error during UrlFetchApp call: " + error);
    return { error: `Network or script error during API call: ${error.message}` };
  }
}

function sendSelectedStories() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // --- Get Column Indices (Error checking for missing headers) ---
  const statusColIndex = headers.indexOf(STATUS_COLUMN_NAME) + 1;
  const emailColIndex = headers.indexOf(EMAIL_COLUMN_NAME) + 1;
  const aiResponseColIndex = headers.indexOf(AI_RESPONSE_COLUMN_NAME) + 1;
  const askAiColIndex = headers.indexOf(ASK_AI_COLUMN_NAME) + 1; // Get prompt

  // Get Send Email? column index
  const sendEmailColIndex = headers.indexOf("Send Email?") + 1;  // Assuming the column is labeled "Send Email?"

  // --- Get the data from all rows ---
  const data = sheet.getDataRange().getValues();  // Gets all data

  // Iterate through rows, skipping the header row
  for (let i = 1; i < data.length; i++) {
    const row = i + 1; // Sheet row (starting from 1), data array index is 0-based
    const sendEmail = data[i][sendEmailColIndex - 1]; // Checkbox value (TRUE/FALSE)

    // If checkbox is TRUE, send the email
    if (sendEmail === true) {
      const studentEmail = data[i][emailColIndex - 1];
      const aiResponse = data[i][aiResponseColIndex - 1];
      const askAiPrompt = data[i][askAiColIndex - 1];  // Get prompt
      const studentName = data[i][STUDENT_NAME_COLUMN]; // "Your name" column

      if (!studentEmail || !aiResponse || !askAiPrompt) {
        Logger.log(`Skipping row ${row} - Missing email, AI response, or prompt.`);
        continue; // Skip to the next row
      }

      // --- Send the Email ---
      let greeting = 'Hi!';
      if (studentName) {
        greeting = `Hi ${studentName}!`; // Personalized greeting
      }

      const subject = "AI Response to your prompt!";
      const body = `${greeting}\n\nYou asked:\n"${askAiPrompt}"\n\nHere's your AI response:\n\n---\n${aiResponse}\n---\n\nEnjoy!`; // Customize as needed

      try {
        MailApp.sendEmail(studentEmail, subject, body);
        sheet.getRange(row, statusColIndex).setValue("Sent");
        Logger.log(`AI response sent to ${studentEmail} for row ${row}`);
      } catch (error) {
        Logger.log("Error sending email: " + error);
        sheet.getRange(row, statusColIndex).setValue("Error - Send Failed");
        if (ADMIN_EMAIL_FOR_ERRORS) MailApp.sendEmail(ADMIN_EMAIL_FOR_ERRORS, "New Generator - Send Failed", `Failed to send response to ${studentEmail} for row ${row}. Error: ${error}`);
      } finally { // Make sure to run this block, even if there is an error.
          // *** UNTICK THE CHECKBOX HERE ***
          sheet.getRange(row, sendEmailColIndex).setValue(false);
      }
    }
  }
  SpreadsheetApp.getUi().alert("Emails sent (or attempted) successfully!"); // Notify the teacher.
}

// --- Function to set the master prompt (Menu item) ---
function setMasterPrompt() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
      'Set Master Prompt',
      'Enter the new master prompt:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response
  if (result.getSelectedButton() == ui.Button.OK) {
    const newPrompt = result.getResponseText();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (settingsSheet) {
      const masterPromptCell = settingsSheet.getRange(MASTER_PROMPT_CELL);
      masterPromptCell.setValue(newPrompt);
      SpreadsheetApp.getUi().alert('Master Prompt updated!');
    } else {
      SpreadsheetApp.getUi().alert('Error: Settings sheet not found.');
    }
  } else {
    // User canceled the prompt
    SpreadsheetApp.getUi().alert('Master prompt update canceled.');
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Q&A Actions')
      .addItem('Send Selected Answers', 'sendSelectedStories') // New menu item
      .addItem('Set Master Prompt', 'setMasterPrompt') // Add the new menu item
      .addToUi();
}
