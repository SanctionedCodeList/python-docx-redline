/*
 * Office Bridge Add-in - Connects PowerPoint to Python bridge server
 */

/* global document, Office, PowerPoint, console */

import { BridgeClient, ConnectionState } from "@shared/bridge-client";

// UI elements
let statusDot: HTMLElement;
let statusText: HTMLElement;
let connectBtn: HTMLButtonElement;
let sessionIdDisplay: HTMLElement;
let lastActivityDisplay: HTMLElement;

// Bridge client instance
let bridgeClient: BridgeClient;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
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

  connectBtn.onclick = handleConnect;
  updateUI("disconnected");
}

function initBridgeClient(): void {
  bridgeClient = new BridgeClient({
    app: "powerpoint",
    getFilename: getPresentationName,
    getUrl: () => Office.context.document?.url || null,
    onStateChange: updateUI,
    onRegistered: (id) => {
      sessionIdDisplay.textContent = `Session: ${id}`;
      sessionIdDisplay.style.display = "block";
    },
    onActivity: updateLastActivity,
  });

  // Set up the execute handler for PowerPoint
  bridgeClient.setExecuteHandler(async (code: string) => {
    return await PowerPoint.run(async (context) => {
      // Create async function from code and execute it
      const fn = new Function(
        "context",
        "PowerPoint",
        "Office",
        `
        return (async () => {
          ${code}
        })();
      `
      );
      return await fn(context, PowerPoint, Office);
    });
  });

  // Set up console forwarding
  bridgeClient.setupConsoleForwarding();

  // Auto-connect on load
  bridgeClient.connect();
}

function getPresentationName(): string {
  try {
    const url = Office.context.document?.url;
    if (url) {
      const parts = url.replace(/\\/g, "/").split("/");
      return parts[parts.length - 1] || "Untitled";
    }
  } catch {
    // Ignore errors
  }
  return "Untitled Presentation";
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
