import { createUpdatedClientDirectoryCard } from "../utils/adaptiveCards.js";
import { createTeamsUpdate } from "../utils/utils.js";
import { mockLongRunningTask } from "../utils/utils.js";
export const handleCallProcessingAgent = async (context) => {
  // parse the selected client directory from the context
  const selectedClientDirectory = JSON.parse(
    context.activity.value.directoryChoice
  );
  // Get the name of the selected client directory
  const selectedDirectoryName = selectedClientDirectory.name;
  console.log(`Processing evidence pack in ${selectedDirectoryName}`);

  console.log(
    `HTTP request sent to Processing Agent for directory ${selectedDirectoryName}`
  );

  // Start the long-running task without awaiting it
  mockLongRunningTask()
    .then((result) => console.log("Task completed:", result))
    .catch((error) => console.error("Task failed:", error));

  await createTeamsUpdate(
    context,
    `🛠️ Processing evidence pack with job ID **${selectedDirectoryName}**`
  );
};
