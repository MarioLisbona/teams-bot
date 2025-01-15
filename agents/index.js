import { ChatOpenAI } from "@langchain/openai";

/**
 * Azure OpenAI chat model instance configured with zero temperature for deterministic responses.
 * Uses environment variables for Azure-specific configuration.
 * @constant {ChatOpenAI}
 * @property {number} temperature - Set to 0 for consistent, deterministic outputs
 * @property {string} azureOpenAIApiKey - Azure OpenAI API key from environment variables
 * @property {string} azureOpenAIApiInstanceName - Azure instance name from environment variables
 * @property {string} azureOpenAIApiDeploymentName - Azure deployment name from environment variables
 * @property {string} azureOpenAIApiVersion - Azure API version from environment variables
 */
export const llm = new ChatOpenAI({
  temperature: 0,
  azureOpenAIApiKey: process.env.AZURE_OPENAI_API_KEY,
  azureOpenAIApiInstanceName: process.env.AZURE_OPENAI_INSTANCE_NAME,
  azureOpenAIApiDeploymentName: process.env.AZURE_OPENAI_DEPLOYMENT_NAME,
  azureOpenAIApiVersion: process.env.AZURE_OPENAI_API_VERSION,
});

export const formatLLMResponse = (result) => {
  return result
    .split("\n")
    .filter((line) => line.trim())
    .join("\n");
};
