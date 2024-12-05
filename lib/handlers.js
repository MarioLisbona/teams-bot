// ... existing code ...
} else if (context.activity.value?.action === "selectDirectory") {
  // ... other existing code ...

  // For Testing files only
  await handleDirectorySelection(context, selectedDirectoryId, { filterPattern: 'Testing' });
}

// For generic use (no filtering)
// await handleDirectorySelection(context, selectedDirectoryId);