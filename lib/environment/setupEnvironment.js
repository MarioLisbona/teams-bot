//====================================================================================
/**
 * Loads environment variables from the appropriate .env file based on NODE_ENV.
 *
 * @description
 * In development mode, attempts to load from .env.local.btrmnt first.
 * Falls back to .env.production if .env.local.btrmnt doesn't exist or if not in development mode.
 *
 * @throws {Error} If environment variables cannot be loaded from the selected file
 * @returns {void}
 */
//====================================================================================
import dotenv from "dotenv";
import fs from "fs";

export function loadEnvironmentVariables() {
  try {
    const envFile =
      process.env.NODE_ENV === "development" && fs.existsSync(".env.local")
        ? ".env.local"
        : ".env.production";

    const result = dotenv.config({ path: envFile });

    if (result.error) {
      throw new Error(`Error loading environment from ${envFile}`);
    }
  } catch (error) {
    console.error("Failed to load environment variables:", error);
    throw error; // Re-throw as this is a critical setup function
  }
}
