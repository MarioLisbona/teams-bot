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
 * This function sends a query to the Azure OpenAI API and returns the response.
 * @param {string} userPrompt - The prompt to send to the Azure OpenAI API.
 * @returns {string} - The response from the Azure OpenAI API.
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
