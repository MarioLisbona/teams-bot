//====================================================================================
/**
 * Function to load environment variables from the appropriate .env file
 * Checks for .env.local in development, falls back to .env.production
 */
//====================================================================================
import dotenv from "dotenv";
import fs from "fs";

export function loadEnvironmentVariables() {
  const envFile =
    process.env.NODE_ENV === "development" && fs.existsSync(".env.local.btrmnt")
      ? ".env.local.btrmnt"
      : ".env.production.btrmnt";
  dotenv.config({ path: envFile });
}
