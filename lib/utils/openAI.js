import dotenv from "dotenv";
import axios from "axios";
import { loadEnvironmentVariables } from "../environment/setupEnvironment.js";

loadEnvironmentVariables();

const OPENAI_API_KEY = process.env.OPENAI_API_KEY;

export async function openAiQuery(prompt) {
  try {
    const openaiResponse = await axios.post(
      "https://api.openai.com/v1/chat/completions",
      {
        model: "gpt-4o-mini",
        messages: [{ role: "user", content: prompt }],
        temperature: 0, // Setting temperature to 0 for deterministic output
      },
      {
        headers: {
          Authorization: `Bearer ${OPENAI_API_KEY}`,
          "Content-Type": "application/json",
        },
      }
    );

    const responseText = openaiResponse.data.choices[0].message.content.trim();

    return responseText;
  } catch (error) {
    console.error("Error generating summary:", error);
    return "Oops, I couldn't generate a response. Please try again.";
  }
}
