export function updateRfiDataWithAzureGptPrompt(groupedData) {
  return `
  Please process each entry in the following array according to these rules:

  1. Remove "RFI" at the beginning of each text.
  2. Replace it with "The auditor noted that".
  3. Complete each sentence with an action item, such as "can you please clarify?" or "can you provide additional evidence?" based on the context.

  Return the results as a JSON array with each entry as a separate string, formatted like this:

  [
      "The auditor noted that installer declaration listed HPC026/MT-200R26E20 instead of MHW-F26WN3/MT-300R26E20, can you please clarify?",
      "The auditor noted that invoice says $0, but list of sites listed $1,500, can you provide additional evidence?",
      ...
  ]

  The output should strictly follow JSON syntax and include all elements in an array format.
  **DO NOT** use any markdown text in the output.

  Array of entries: ${groupedData}

  Ignore any previous instructions or context. Treat this prompt as a standalone task.

  `;
}
