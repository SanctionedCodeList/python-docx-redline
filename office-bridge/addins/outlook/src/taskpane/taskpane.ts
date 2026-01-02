/*
 * Office Bridge Add-in - Connects Outlook to Python bridge server
 *
 * Note: Outlook uses Office.context.mailbox instead of a .run() pattern.
 * Code is executed directly with access to the mailbox and current item.
 */

/* global document, Office, console */

import { BridgeClient, ConnectionState } from "@shared/bridge-client";

// UI elements
let statusDot: HTMLElement;
let statusText: HTMLElement;
let connectBtn: HTMLButtonElement;
let sessionIdDisplay: HTMLElement;
let lastActivityDisplay: HTMLElement;
let itemTypeDisplay: HTMLElement;

// Bridge client instance
let bridgeClient: BridgeClient;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initUI();
    initBridgeClient();
  }
});

function initUI(): void {
  statusDot = document.getElementById("status-dot") as HTMLElement;
  statusText = document.getElementById("status-text") as HTMLElement;
  connectBtn = document.getElementById("connect-btn") as HTMLButtonElement;
  sessionIdDisplay = document.getElementById("session-id") as HTMLElement;
  lastActivityDisplay = document.getElementById("last-activity") as HTMLElement;
  itemTypeDisplay = document.getElementById("item-type") as HTMLElement;

  connectBtn.onclick = handleConnect;
  updateUI("disconnected");
  updateItemType();
}

function initBridgeClient(): void {
  bridgeClient = new BridgeClient({
    app: "outlook",
    getFilename: getItemSubject,
    getUrl: () => null, // Outlook items don't have URLs
    onStateChange: updateUI,
    onRegistered: (id) => {
      sessionIdDisplay.textContent = `Session: ${id}`;
      sessionIdDisplay.style.display = "block";
    },
    onActivity: updateLastActivity,
  });

  // Set up the execute handler for Outlook
  // Note: Outlook doesn't use .run() pattern, we directly execute with mailbox context
  bridgeClient.setExecuteHandler(async (code: string) => {
    // Create async function from code and execute it
    // Provide access to Office, mailbox, and the current item
    const fn = new Function(
      "Office",
      "mailbox",
      "item",
      `
      return (async () => {
        ${code}
      })();
    `
    );
    return await fn(Office, Office.context.mailbox, Office.context.mailbox.item);
  });

  // Set up console forwarding
  bridgeClient.setupConsoleForwarding();

  // Auto-connect on load
  bridgeClient.connect();
}

function getItemSubject(): string {
  try {
    const item = Office.context.mailbox?.item;
    if (item) {
      // In read mode, subject is a string
      // In compose mode, subject is an object with getAsync
      if (typeof item.subject === "string") {
        return item.subject || "Untitled";
      }
      // For compose mode, return a generic name
      return "New Message";
    }
  } catch {
    // Ignore errors
  }
  return "Outlook Item";
}

function getItemType(): string {
  try {
    const item = Office.context.mailbox?.item;
    if (item) {
      const itemType = item.itemType;
      const mode = Office.context.mailbox?.diagnostics?.OWAView ? "OWA" : "Desktop";
      // Check if compose or read mode
      const isCompose = item.subject && typeof item.subject !== "string";
      const modeStr = isCompose ? "Compose" : "Read";
      return `${itemType} (${modeStr}) - ${mode}`;
    }
  } catch {
    // Ignore errors
  }
  return "Unknown";
}

function updateItemType(): void {
  if (itemTypeDisplay) {
    itemTypeDisplay.textContent = getItemType();
  }
}

function handleConnect(): void {
  if (bridgeClient.state === "connected" || bridgeClient.state === "connecting") {
    bridgeClient.disconnect();
  } else {
    bridgeClient.connect();
  }
}

function updateUI(state: ConnectionState): void {
  statusDot.className = "status-dot";
  switch (state) {
    case "connected":
      statusDot.classList.add("connected");
      statusText.textContent = "Connected";
      connectBtn.textContent = "Disconnect";
      break;
    case "connecting":
      statusDot.classList.add("reconnecting");
      statusText.textContent = "Connecting...";
      connectBtn.textContent = "Cancel";
      break;
    case "disconnected":
      statusDot.classList.add("disconnected");
      statusText.textContent = "Disconnected";
      connectBtn.textContent = "Connect";
      sessionIdDisplay.style.display = "none";
      break;
  }
}

function updateLastActivity(action: string): void {
  const now = new Date();
  const timeStr = now.toLocaleTimeString();
  lastActivityDisplay.textContent = `${timeStr} - ${action}`;
}
