import dotenv from "dotenv";
import axios from "axios";
import { loadEnvironmentVariables } from "../environment/setupEnvironment.js";

loadEnvironmentVariables();

const OPENAI_API_KEY = process.env.OPENAI_API_KEY;

/**
 * Sends prompts to OpenAI's GPT-4 API and processes responses.
 *
 * @description
 * Handles complete API interaction workflow:
 * 1. Validates API key availability
 * 2. Sends request to chat completions endpoint
 * 3. Uses deterministic response generation (temperature = 0)
 * 4. Validates and processes API response
 * 5. Provides comprehensive error handling
 *
 * Error handling includes:
 * - API key validation
 * - Request failure detection
 * - Response format validation
 * - Empty response checking
 * - Detailed error logging
 *
 * @param {string} prompt - Text prompt for GPT-4 model
 *
 * @throws {Error} When API key is missing or invalid
 * @throws {Error} When API request fails
 * @throws {Error} When response format is invalid
 * @throws {Error} When response is empty
 * @returns {Promise<string>} Generated response text or error message
 *   - Success: Trimmed response from GPT-4
 *   - Failure: User-friendly error message
 *
 * @example
 * try {
 *   const response = await openAiQuery(
 *     "Analyze this audit finding:"
 *   );
 *   console.log(response);
 * } catch (error) {
 *   console.error("Query failed:", error);
 * }
 *
 * @requires axios
 * @requires OPENAI_API_KEY environment variable
 *
 * Configuration:
 * - Model: gpt-4o-mini
 * - Temperature: 0 (deterministic)
 * - Request timeout: default
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
