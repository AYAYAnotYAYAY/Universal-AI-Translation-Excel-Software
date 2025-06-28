# Universal AI Translation Excel Software

This is a universal AI translation tool for Excel files, supporting various AI models and providing a user-friendly interface for batch translation tasks.

## Features

-   **Excel Integration**: Preview Excel files directly within the application and select columns for translation.
-   **Multi-AI Support**:
    -   **Gemini**: Supports Google's Gemini models via their OpenAI-compatible endpoint.
    -   **DeepSeek**: Supports DeepSeek models.
    -   **Custom**: Supports any other AI service that has an OpenAI-compatible API.
-   **Dynamic Model Lists**: Automatically fetches and displays available models from Gemini and DeepSeek after you provide an API key.
-   **Robust Batch Translation**: Translates text in batches of 100 rows to improve efficiency and reduce API calls.
-   **Intelligent Error Handling**:
    -   Automatically retries on API rate limit errors (`429`), parsing the recommended wait time.
    -   Strictly validates that the number of translated lines matches the number of source lines to prevent data misalignment.
-   **Proxy Support**: Configure and use HTTP or SOCKS5 proxies for network requests.
-   **Persistent Configuration**: Saves your AI model and proxy settings locally in a `config.json` file.

## How to Use

1.  **Configure AI Model**:
    -   Go to "Model Management".
    -   Add a new configuration, select the provider (e.g., Gemini, DeepSeek), and enter your API key.
    -   Click "Get Model List" to choose a specific model.
    -   Save the configuration.
2.  **Load Excel File**:
    -   Click "Browse..." to load your `.xlsx` or `.xls` file.
    -   The content will be displayed in the preview panel.
3.  **Select Columns**:
    -   **Left-click** on a cell in the column you want to translate from. This sets the "Source Column" and the starting row.
    -   **Right-click** on a column to set it as the "Target Column" where translations will be placed.
4.  **Translate**:
    -   Ensure the correct AI model and languages are selected.
    -   A dialog will ask you to confirm that the Excel file is closed.
    -   Click "Start Translation". The progress will be shown in the log panel.

---

### Version History

#### v2.0

**Major Features & Enhancements**

-   **DeepSeek API Integration**: Added native support for DeepSeek as an API provider, including `deepseek-chat` and `deepseek-reasoner` models.
-   **Smart Model Fetching**: The app now automatically fetches available models for Gemini and DeepSeek once an API key is entered.
-   **Robust Batching Engine**: Re-architected the translation logic to process 100 rows per batch, significantly reducing API calls and improving stability.
-   **Advanced Prompt Engineering**: Implemented a new default prompt with unique delimiters and explicit instructions (including examples) to ensure the AI returns data in the correct format, resolving line-mismatch errors.

**Bug Fixes & Stability**

-   **Rate Limit Auto-Retry**: The application now intelligently handles `429` errors by waiting for the API-suggested duration before automatically retrying.
-   **File Lock Prevention**: Changed the file saving logic to perform a single save at the end of the entire translation process and added a pre-flight check to ensure the target file is closed, eliminating `Permission Denied` errors.
-   **UI & Startup Fixes**: Corrected bugs related to UI updates from background threads and fixed a critical f-string `NameError` that caused the application to crash on startup.