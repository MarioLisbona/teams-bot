import { azureGptQuery } from "./azureGpt.cjs";
import { updateRfiDataWithAzureGptPrompt } from "./prompts.js";

export async function updateRfiDataWithAzureGptQuery(data) {
  // Create the prompt for the OpenAI API
  const prompt = await updateRfiDataWithAzureGptPrompt(data);

  // Send the prompt to the OpenAI API and return the response
  const amendedRdiData = await azureGptQuery(prompt);

  return amendedRdiData;
}
