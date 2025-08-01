# Excel AI Function Assistant User Guide

## Overview
This tool provides core AI interaction capabilities for Excel, enabling AI processing of cell content through functions. It supports batch data operations, significantly improving data processing efficiency.

## Key Features
1.  **One-click Configuration**: Quickly configure API Key, model, and endpoint using the `SetupMyAI` macro. Supports various large models (requires applying for API Keys independently).
2.  **Dynamic Invocation**: Processes both single cells and range data.
3.  **Error Handling**: Automatically captures and displays error messages during API calls.

## Usage Instructions
1.  Open the Excel workbook and press `Alt+F11` to open the VBA editor.
2.  In the VBA editor, select "File" → "Import File..." and import the `EXCEL AI.bas` module file.
3.  Close the VBA editor, return to the Excel interface, press `Alt+F8`, select and run the `SetupMyAI` macro.
4.  Follow the prompts to configure the model name, API endpoint (URL), and API Key.
5.  Use the `AI_PROCESS(cellData As Range, prompt As String)` function to process single cells or ranges.
6.  (Optional) To reset configuration, run the `ResetAIConfig` macro to change models or API Keys.

## Important Notes
1.  Model configurations are automatically stored for repeated use, avoiding redundant input.
2.  **Save Workbooks**: To retain configurations long-term, save workbooks as `.xlsm` (macro-enabled workbook) format. When reopening, simply run `SetupMyAI` to load configurations—no re-entry needed.
3.  **Security Risk**: API Keys and model information are stored in a **hidden worksheet** named `_AIConfig_`. **Sharing or publishing such workbooks poses severe API Key leakage risks!**
4.  **Temporary Storage Risk**: Even if not saved as `.xlsm`, running `SetupMyAI` in a regular `.xlsx` file stores API Keys/model info in the hidden worksheet, **still creating leakage risks**.
5.  **Safe Sharing Recommendations**: To share results securely:
    *   **Copy** cells containing `AI_PROCESS` function results.
    *   **Paste Special → Values** (convert function results to static values).
    *   Copy **only the result values** to a new workbook for sharing.

## Operational Tips
1.  **Save Configuration**: After successfully running `SetupMyAI`, configurations are auto-saved to the hidden worksheet `_AIConfig_`.
2.  **Processing Scale**: Due to model context window limitations, **control batch cell counts and text length per cell**.
3.  **Use Cases**: **Avoid** using large models for simple calculations (e.g., SUM, AVERAGE, COUNT). Excel's built-in functions are faster, more accurate, and hallucination-free. AI functions excel at complex tasks like text processing, comprehension, generation, translation, summarization, and analysis.
4.  **Optimize Prompts**: Ensure concise outputs by ending prompts with directives like "**Output the result directly without explanations or processes**". Example: "Analyze the sentiment of the following text. Output only the result (Positive/Negative/Neutral) without explanations".

## Practical Examples
### Direct Cell AI Function Calls
1.  **Summarization**:
    *   Enter text in cell A1.
    *   Enter formula in B1:  
        `=AI_PROCESS(A1, "Summarize the following text in 100 words. Output the result directly without explanations")`  
    *   **Output**: B1 displays the summary.

2.  **Translation**:
    *   Enter Chinese text in A2.
    *   Enter formula in B2:  
        `=AI_PROCESS(A2, "Translate the following text to English. Output the result directly without explanations")`  
    *   **Output**: B2 displays the English translation.

3.  **Sentiment Analysis**:
    *   Enter a review in A3.
    *   Enter formula in B3:  
        `=AI_PROCESS(A3, "Analyze the sentiment of the following text. Output only the result (Positive/Negative/Neutral) without explanations")`  
    *   **Output**: B3 displays the sentiment (e.g., "Positive").

### Dynamic Cross-Cell Processing
*   **Example**: Batch summarization for Column A into Column B:
    1.  Enter formula in B1:  
        `=AI_PROCESS(A1, "Summarize the following content. Output the result directly without explanations")`  
    2.  Drag B1's fill handle (small square at cell's bottom-right corner) downward to auto-fill formulas for corresponding Column A data.

### Range Data Processing
*   **Example**: Analyze trend of numeric sequence in range A1:A10:
    1.  Enter formula in B1:  
        `=AI_PROCESS(A1:A10, "Analyze the trend of this data sequence (e.g., increasing, decreasing, fluctuating)")`  
    2.  **Output**: B1 displays trend analysis (e.g., "This sequence shows a stable linear upward trend").

## FAQs
1.  **API Call Failure**:
    *   Verify the API Key is correct and active.
    *   Ensure model name and endpoint URL are correctly configured and match the selected model.
    *   Confirm network connectivity to the API endpoint.
    *   Check Excel's status bar or pop-up error messages.
2.  **Function Returns Error (e.g., #VALUE!)**:
    *   Check if the `cellData` parameter references valid cells/ranges.
    *   Verify `prompt` is a valid text string.
    *   Ensure `SetupMyAI` was successfully executed.
3.  **Unexpected Results**:
    *   Optimize prompts: make them clearer, more specific, and include constraints like "output directly".
    *   Check if input data is within the model's comprehension scope.
    *   Simplify tasks or reduce batch size.
4.  **Missing or Deleted Hidden Worksheet `_AIConfig_`**:
    *   Re-run `SetupMyAI` to recreate the worksheet during configuration.
