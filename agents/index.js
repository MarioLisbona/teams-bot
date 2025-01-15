import { ChatOpenAI } from "@langchain/openai";

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
