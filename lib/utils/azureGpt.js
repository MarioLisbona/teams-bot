import { AzureOpenAI } from "openai";
import { loadEnvironmentVariables } from "../environment/setupEnvironment.js";

loadEnvironmentVariables();

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
 * @throws {Error} When Azure OpenAI client initialization fails
 * @throws {Error} When API request fails
 * @returns {Promise<string>} Trimmed response text from GPT-4
 */
async function azureGptQuery(userPrompt) {
  console.log("azureGptQuery called");
  try {
    const deployment = process.env.AZURE_OPENAI_DEPLOYMENT_NAME;
    const apiVersion = "2024-10-21";

    try {
      const client = new AzureOpenAI({
        apiKey: process.env.AZURE_OPENAI_API_KEY,
        deployment,
        apiVersion,
      });

      const response = await client.chat.completions.create({
        messages: [{ role: "user", content: userPrompt }],
        model: "gpt-4o",
      });

      return response.choices[0].message.content.trim();
    } catch (error) {
      console.error("Azure OpenAI API request failed:", error);
      throw new Error(`Failed to get AI response: ${error.message}`);
    }
  } catch (error) {
    console.error("Azure GPT query failed:", error);
    throw new Error(`GPT query failed: ${error.message}`);
  }
}

export { azureGptQuery };
