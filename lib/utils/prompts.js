/**
 * Creates an AI prompt for analyzing client audit responses against historical data.
 *
 * @description
 * Generates a structured prompt that:
 * 1. Sets AI assistant context for audit analysis
 * 2. Defines knowledge base structure and fields:
 *    - issuesIdentified: Original audit findings
 *    - acpResponse: Client's response to RFI
 *    - auditorNotes: Evaluation status and notes
 * 3. Explains audit workflow process
 * 4. Specifies valid status codes:
 *    - "Finding and recommendation"
 *    - "Observation"
 *    - "No issues"
 *    - "RFI"
 *
 * @param {Object} knowledgeBase - Historical audit response data
 * @param {Array<Object>} knowledgeBase.entries - Previous audit entries
 * @param {string} knowledgeBase.entries[].issuesIdentified - Original issue
 * @param {string} knowledgeBase.entries[].acpResponse - Client response
 * @param {string} knowledgeBase.entries[].auditorNotes - Previous evaluation
 *
 * @param {Object} clientResponses - Current responses requiring analysis
 * @param {string} clientResponses.issuesIdentified - Current issue
 * @param {string} clientResponses.acpResponse - Current client response
 *
 * @returns {string} Formatted prompt for OpenAI analysis including:
 *   - Assistant role definition
 *   - Data structure explanation
 *   - Workflow description
 *   - Analysis instructions
 *   - Historical and current data
 */
export function analyseAcpResponsePrompt(knowledgeBase, clientResponses) {
  return `
You are a helpful assistant that can help me analyze responses from clients that are being audited.

The knowledge base contains responses from previous requests for information (RFI) from auditors. The format of the knowledge base is JSON with the following fields:
- issuesIdentified: The issue that the auditor has identified.
- acpResponse: The response from the client to the RFI.
- auditorNotes: The notes from the auditor about the response.

The workflow is as follows:
1. An auditor will identify an issue in a job that is being audited.
  - Attribute used: issuesIdentified
2. The auditor will then send an RFI to the client asking for more information regarding the issue.
3. The client will then respond to the RFI with a response.
  - Attribute used: acpResponse
4. Your job is to analyze the response and see if it is satisfactory and fill in the auditorNotes field:
  - Attribute used: auditorNotes

The auditorNotes field should **only** contain one of the following statuses:
- "Finding and recommendation"
- "Observation"
- "No issues"
- "RFI"

Provide additional information with the status code in the auditorNotes field if the response does not satisfactorily address the issue.

Your task is to analyze the knowledge base and identify correlations and patterns between the issuesIdentified and acpResponse attributes and their relationships to the auditorNotes and closedOut attributes.

Use these correlations to infer and generate the values for auditorNotes when only issuesIdentified and acpResponse are provided. Ensure consistency with the patterns observed in the knowledge base.

You will be provided with an incomplete json object where only issuesIdentified and acpResponse are given, generate detailed auditorNotes using the specified status codes (F, M, or RFI) and relevant context.

Here is the knowledge base:
${JSON.stringify(knowledgeBase)}

Here are the responses from the client:
${JSON.stringify(clientResponses)}

**Return the completed json object only. DO NOT use markdown formatting.**

`;
}
