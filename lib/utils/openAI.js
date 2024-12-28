import dotenv from "dotenv";
import axios from "axios";
import { loadEnvironmentVariables } from "../environment/setupEnvironment.js";

loadEnvironmentVariables();

const OPENAI_API_KEY = process.env.OPENAI_API_KEY;

/**
 * Sends a query to the OpenAI API and returns the response.
 * @param {string} prompt - The prompt to send to the OpenAI API.
 * @returns {Promise<string>} The generated response text from the API.
 * @throws {Error} If the API request fails.
 * @throws {Error} If the response parsing fails.
 * @throws {Error} If the API key is missing or invalid.
 */
export async function openAiQuery(prompt) {
  try {
    // Validate API key
    if (!OPENAI_API_KEY) {
      throw new Error("OpenAI API key is missing");
    }

    try {
      // Make API request
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

      try {
        // Parse and validate response
        if (!openaiResponse.data?.choices?.[0]?.message?.content) {
          throw new Error("Invalid response format from OpenAI API");
        }

        const responseText =
          openaiResponse.data.choices[0].message.content.trim();
        if (!responseText) {
          throw new Error("Empty response from OpenAI API");
        }

        return responseText;
      } catch (error) {
        console.error("Error parsing API response:", error);
        throw new Error(`Failed to parse OpenAI response: ${error.message}`);
      }
    } catch (error) {
      console.error("API request failed:", {
        status: error.response?.status,
        statusText: error.response?.statusText,
        data: error.response?.data,
        message: error.message,
      });
      throw new Error(`OpenAI API request failed: ${error.message}`);
    }
  } catch (error) {
    console.error("Error generating response:", error);
    // Return a user-friendly error message while still logging the actual error
    return "Oops, I couldn't generate a response. Please try again.";
  }
}
