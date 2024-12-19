import { createTeamsUpdate } from "../utils/utils.js";

/**
 * This function handles the call to the Evidence Pack Processing Agent.
 * It sends the selected client directory to the agent.
 * @param {Object} context - The context object containing the activity from the Teams bot.
 */
export const handleCallProcessJobAgent = async (context) => {
  // parse the selected client directory from the context
  const selectedClientDirectory = JSON.parse(
    context.activity.value.directoryChoice
  );
  // Get the name of the selected client directory
  const selectedDirectoryName = selectedClientDirectory.name;
  const selectedDirectoryID = selectedClientDirectory.id;
  console.log(
    `Sending ${selectedDirectoryName} ID: ${selectedDirectoryID} to Evidence Pack Processing Agent`
  );

  await createTeamsUpdate(
    context,
    `Sending ${selectedDirectoryName} ID: ${selectedDirectoryID} to Evidence Pack Processing Agent`
  );
};
