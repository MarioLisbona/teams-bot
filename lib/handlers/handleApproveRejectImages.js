export async function handleApproveRejectImages(context) {
  console.log("handleApproveRejectImages function called");
  const action = context.activity.value.action;
  const images = context.activity.value.images;

  // Delete the original review card
  await context.deleteActivity(context.activity.replyToId);

  // Send processing card
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: {
          type: "AdaptiveCard",
          version: "1.0",
          body: [
            {
              type: "Container",
              style: "emphasis",
              items: [
                {
                  type: "ColumnSet",
                  columns: [
                    {
                      type: "Column",
                      width: "auto",
                      verticalContentAlignment: "center",
                      items: [
                        {
                          type: "TextBlock",
                          text: "🔄",
                          size: "extraLarge",
                          spacing: "none",
                        },
                      ],
                    },
                    {
                      type: "Column",
                      width: "stretch",
                      items: [
                        {
                          type: "TextBlock",
                          text: `Processing ${action}`,
                          weight: "bolder",
                          size: "medium",
                          color: "accent",
                        },
                        {
                          type: "TextBlock",
                          text: "Processing in Progress",
                          spacing: "none",
                          isSubtle: true,
                        },
                      ],
                    },
                  ],
                },
              ],
              padding: "default",
            },
          ],
        },
      },
    ],
  });

  // Simulate processing
  await new Promise((resolve) => setTimeout(resolve, 5000));

  // Send completion card
  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: {
          type: "AdaptiveCard",
          version: "1.0",
          body: [
            {
              type: "Container",
              style: action === "approve" ? "good" : "attention",
              items: [
                {
                  type: "ColumnSet",
                  columns: [
                    {
                      type: "Column",
                      width: "auto",
                      verticalContentAlignment: "center",
                      items: [
                        {
                          type: "TextBlock",
                          text: action === "approve" ? "✅" : "❌",
                          size: "extraLarge",
                          spacing: "none",
                        },
                      ],
                    },
                    {
                      type: "Column",
                      width: "stretch",
                      items: [
                        {
                          type: "TextBlock",
                          text: `Images ${
                            action === "approve" ? "Approved" : "Rejected"
                          }`,
                          weight: "bolder",
                          size: "medium",
                        },
                        {
                          type: "TextBlock",
                          text: `${
                            action === "approve" ? "Approval" : "Rejection"
                          } Complete`,
                          spacing: "none",
                          color: action === "approve" ? "good" : "attention",
                        },
                      ],
                    },
                  ],
                },
              ],
              padding: "default",
            },
          ],
        },
      },
    ],
  });
}
