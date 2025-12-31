/*
 * Office Bridge Add-in - Connects Word to Python bridge server
 */

/* global document, Office, Word, console */

// Connection state
type ConnectionState = "disconnected" | "connecting" | "connected";

let ws: WebSocket | null = null;
let connectionState: ConnectionState = "disconnected";
let reconnectAttempts = 0;
let reconnectTimeout: number | null = null;
let documentId: string | null = null;

// Original console methods for forwarding
const originalConsole = {
  log: console.log.bind(console),
  warn: console.warn.bind(console),
  error: console.error.bind(console),
};

// UI elements
let statusDot: HTMLElement;
let statusText: HTMLElement;
let tokenInput: HTMLInputElement;
let connectBtn: HTMLButtonElement;
let documentIdDisplay: HTMLElement;
let lastActivityDisplay: HTMLElement;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Get UI elements
    statusDot = document.getElementById("status-dot") as HTMLElement;
    statusText = document.getElementById("status-text") as HTMLElement;
    tokenInput = document.getElementById("token-input") as HTMLInputElement;
    connectBtn = document.getElementById("connect-btn") as HTMLButtonElement;
    documentIdDisplay = document.getElementById("document-id") as HTMLElement;
    lastActivityDisplay = document.getElementById("last-activity") as HTMLElement;

    // Set up event handlers
    connectBtn.onclick = handleConnect;
    tokenInput.onkeypress = (e) => {
      if (e.key === "Enter") handleConnect();
    };

    // Update initial UI
    updateUI();

    // Override console methods to forward to bridge
    setupConsoleForwarding();
  }
});

function setupConsoleForwarding() {
  console.log = (...args: unknown[]) => {
    originalConsole.log(...args);
    sendConsoleMessage("log", args);
  };

  console.warn = (...args: unknown[]) => {
    originalConsole.warn(...args);
    sendConsoleMessage("warn", args);
  };

  console.error = (...args: unknown[]) => {
    originalConsole.error(...args);
    sendConsoleMessage("error", args);
  };
}

function sendConsoleMessage(level: string, args: unknown[]) {
  if (ws && ws.readyState === WebSocket.OPEN) {
    try {
      const message = args
        .map((arg) => (typeof arg === "object" ? JSON.stringify(arg) : String(arg)))
        .join(" ");
      ws.send(
        JSON.stringify({
          type: "console",
          level,
          message,
        })
      );
    } catch {
      // Ignore serialization errors
    }
  }
}

function updateUI() {
  // Update status indicator
  statusDot.className = "status-dot";
  switch (connectionState) {
    case "connected":
      statusDot.classList.add("connected");
      statusText.textContent = "Connected";
      connectBtn.textContent = "Disconnect";
      tokenInput.disabled = true;
      break;
    case "connecting":
      statusDot.classList.add("reconnecting");
      statusText.textContent = "Connecting...";
      connectBtn.textContent = "Cancel";
      tokenInput.disabled = true;
      break;
    case "disconnected":
      statusDot.classList.add("disconnected");
      statusText.textContent = "Disconnected";
      connectBtn.textContent = "Connect";
      tokenInput.disabled = false;
      break;
  }

  // Update document ID display
  if (documentId) {
    documentIdDisplay.textContent = `Document: ${documentId}`;
    documentIdDisplay.style.display = "block";
  } else {
    documentIdDisplay.style.display = "none";
  }
}

function updateLastActivity(action: string) {
  const now = new Date();
  const timeStr = now.toLocaleTimeString();
  lastActivityDisplay.textContent = `${timeStr} - ${action}`;
}

function getDocumentFilename(): string {
  try {
    const url = Office.context.document.url;
    if (url) {
      // Extract filename from URL or path
      const parts = url.replace(/\\/g, "/").split("/");
      return parts[parts.length - 1] || "Untitled";
    }
  } catch {
    // Ignore errors
  }
  return "Untitled";
}

function handleConnect() {
  if (connectionState === "connected" || connectionState === "connecting") {
    disconnect();
  } else {
    connect();
  }
}

function connect() {
  const token = tokenInput.value.trim();
  if (!token) {
    statusText.textContent = "Please enter a token";
    return;
  }

  connectionState = "connecting";
  updateUI();

  try {
    ws = new WebSocket("wss://localhost:3847");

    ws.onopen = () => {
      connectionState = "connected";
      reconnectAttempts = 0;
      updateUI();
      updateLastActivity("Connected");

      // Send registration message
      const filename = getDocumentFilename();
      ws!.send(
        JSON.stringify({
          type: "register",
          token: token,
          document: {
            filename: filename,
            url: Office.context.document.url || null,
          },
        })
      );
    };

    ws.onmessage = async (event) => {
      try {
        const msg = JSON.parse(event.data);
        await handleMessage(msg);
      } catch (err) {
        originalConsole.error("Failed to handle message:", err);
      }
    };

    ws.onclose = () => {
      ws = null;
      if (connectionState === "connected") {
        // Unexpected close, try to reconnect
        scheduleReconnect();
      } else {
        connectionState = "disconnected";
        updateUI();
      }
    };

    ws.onerror = (err) => {
      originalConsole.error("WebSocket error:", err);
      updateLastActivity("Connection error");
    };
  } catch (err) {
    originalConsole.error("Failed to connect:", err);
    connectionState = "disconnected";
    updateUI();
    updateLastActivity("Connection failed");
  }
}

function disconnect() {
  if (reconnectTimeout) {
    clearTimeout(reconnectTimeout);
    reconnectTimeout = null;
  }
  reconnectAttempts = 0;

  if (ws) {
    ws.close();
    ws = null;
  }

  connectionState = "disconnected";
  documentId = null;
  updateUI();
  updateLastActivity("Disconnected");
}

function scheduleReconnect() {
  if (reconnectTimeout) return;

  // Exponential backoff: 1s, 2s, 4s, 8s, 16s, 30s max
  const delay = Math.min(1000 * Math.pow(2, reconnectAttempts), 30000);
  reconnectAttempts++;

  connectionState = "connecting";
  statusText.textContent = `Reconnecting in ${delay / 1000}s...`;
  updateUI();

  reconnectTimeout = window.setTimeout(() => {
    reconnectTimeout = null;
    connect();
  }, delay);
}

async function handleMessage(msg: { type: string; id?: string; payload?: { id?: string; code?: string } }) {
  switch (msg.type) {
    case "registered":
      documentId = msg.payload?.id || null;
      updateUI();
      updateLastActivity("Registered");
      break;

    case "execute":
      if (msg.id && msg.payload?.code) {
        updateLastActivity("Executing code...");
        await executeCode(msg.id, msg.payload.code);
      }
      break;

    case "ping":
      if (ws && ws.readyState === WebSocket.OPEN) {
        ws.send(JSON.stringify({ type: "pong" }));
      }
      break;

    default:
      originalConsole.log("Unknown message type:", msg.type);
  }
}

async function executeCode(requestId: string, code: string) {
  try {
    const result = await Word.run(async (context) => {
      // Create async function from code and execute it
      // The code has access to context, Word, and Office globals
      const fn = new Function(
        "context",
        "Word",
        "Office",
        `
        return (async () => {
          ${code}
        })();
      `
      );
      return await fn(context, Word, Office);
    });

    // Send success response
    if (ws && ws.readyState === WebSocket.OPEN) {
      ws.send(
        JSON.stringify({
          type: "result",
          id: requestId,
          success: true,
          result: serializeResult(result),
        })
      );
    }
    updateLastActivity("Code executed successfully");
  } catch (err) {
    // Send error response
    const errorMessage = err instanceof Error ? err.message : String(err);
    const errorStack = err instanceof Error ? err.stack : undefined;

    if (ws && ws.readyState === WebSocket.OPEN) {
      ws.send(
        JSON.stringify({
          type: "result",
          id: requestId,
          success: false,
          error: {
            message: errorMessage,
            stack: errorStack,
          },
        })
      );
    }
    updateLastActivity("Code execution failed");
    originalConsole.error("Execution error:", err);
  }
}

function serializeResult(result: unknown): unknown {
  if (result === undefined) return null;
  if (result === null) return null;
  if (typeof result === "string" || typeof result === "number" || typeof result === "boolean") {
    return result;
  }
  if (Array.isArray(result)) {
    return result.map(serializeResult);
  }
  if (typeof result === "object") {
    try {
      // Try to convert to plain object
      const obj: Record<string, unknown> = {};
      for (const key of Object.keys(result)) {
        obj[key] = serializeResult((result as Record<string, unknown>)[key]);
      }
      return obj;
    } catch {
      return String(result);
    }
  }
  return String(result);
}
