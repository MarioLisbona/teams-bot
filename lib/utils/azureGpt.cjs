const { AzureOpenAI } = require("openai");

// Use dynamic import for ES Module
(async () => {
  try {
    const { loadEnvironmentVariables } = await import(
      "../environment/setupEnvironment.js"
    );
    await loadEnvironmentVariables();
  } catch (error) {
    console.error("Failed to load environment variables:", error);
    throw new Error(`Environment setup failed: ${error.message}`);
  }
})();

/**
 * Azure OpenAI integration module for GPT-4 queries.
 *
 * @module azureGpt
 * @requires openai
 * @requires ../environment/setupEnvironment
 *
 * Environment variables required:
 * - AZURE_OPENAI_API_KEY: Authentication key for Azure OpenAI service
 * - AZURE_OPENAI_DEPLOYMENT_NAME: Name of the deployed model
 * - AZURE_OPENAI_ENDPOINT: Azure OpenAI service endpoint
 */

/**
 * Sends a query to Azure OpenAI and retrieves the GPT-4 response.
 *
 * @description
 * Handles the complete Azure OpenAI interaction:
 * 1. Initializes Azure OpenAI client with environment credentials
 * 2. Sends chat completion request with user prompt
 * 3. Extracts and processes response text
 * 4. Provides error handling for API interactions
 *
 * @param {string} userPrompt - Text prompt to send to GPT-4
 *
 * @throws {Error} When environment setup fails
 * @throws {Error} When Azure OpenAI client initialization fails
 * @throws {Error} When API request fails
 * @returns {Promise<string>} Trimmed response text from GPT-4
 *
 * @example
 * try {
 *   const response = await azureGptQuery("Analyze this audit data");
 *   console.log(response);
 * } catch (error) {
 *   console.error("GPT query failed:", error);
 * }
 *
 * @version 2024-10-21
 * @since 1.0.0
 */
async function azureGptQuery(userPrompt) {
  try {
    // Define the scope, deployment, and API version
    // TODO: Investigate why scope isnt used and isnt needed
    const scope = process.env.AZURE_OPENAI_ENDPOINT;
    const deployment = process.env.AZURE_OPENAI_DEPLOYMENT_NAME;
    const apiVersion = "2024-10-21";

    try {
      // Create a client with the Azure OpenAI API key
      const client = new AzureOpenAI({
        apiKey: process.env.AZURE_OPENAI_API_KEY,
        deployment,
        apiVersion,
      });

      // Create a chat completion request to the Azure OpenAI API
      const response = await client.chat.completions.create({
        messages: [{ role: "user", content: userPrompt }],
        model: "gpt-4o",
      });

      // Extract and return the response text
      const responseText = response.choices[0].message.content.trim();
      return responseText;
    } catch (error) {
      console.error("Azure OpenAI API request failed:", error);
      throw new Error(`Failed to get AI response: ${error.message}`);
    }
  } catch (error) {
    console.error("Azure GPT query failed:", error);
    throw new Error(`GPT query failed: ${error.message}`);
  }
}

module.exports = { azureGptQuery };
