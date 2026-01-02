/*
 * Office Bridge Add-in Commands - Outlook
 */

/* global Office */

Office.onReady(() => {
  // Commands ready
});

function action(_event: Office.AddinCommands.Event): void {
  // Placeholder for future command actions
  _event.completed();
}

// Register the command
Office.actions?.associate?.("action", action);
