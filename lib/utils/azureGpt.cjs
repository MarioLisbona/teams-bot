const { AzureOpenAI } = require("openai");

// Use dynamic import for ES Module
// This needs to be async - not sure why!??
(async () => {
  const { loadEnvironmentVariables } = await import(
    "../environment/setupEnvironment.js"
  );
  await loadEnvironmentVariables();
})();

async function azureGptQuery(userPrompt) {
  // Define the scope, deployment, and API version
  // TODO: Investigate why scope isnt used and isnt needed
  const scope = process.env.AZURE_OPENAI_ENDPOINT;
  const deployment = process.env.AZURE_OPENAI_DEPLOYMENT_NAME;
  const apiVersion = "2024-10-21";

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
}

module.exports = { azureGptQuery };
