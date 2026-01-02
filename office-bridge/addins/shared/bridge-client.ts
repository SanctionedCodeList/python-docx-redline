/**
 * Shared Bridge Client for Office Add-ins
 *
 * This module provides a reusable WebSocket client that connects
 * Office add-ins to the bridge server.
 */

/* global WebSocket, console */

export type AppType = 'word' | 'excel' | 'powerpoint' | 'outlook';
export type ConnectionState = 'disconnected' | 'connecting' | 'connected';

export interface BridgeClientOptions {
  app: AppType;
  getFilename: () => string;
  getUrl: () => string | null;
  onStateChange?: (state: ConnectionState) => void;
  onRegistered?: (id: string) => void;
  onActivity?: (action: string) => void;
}

export interface ExecuteHandler {
  (code: string): Promise<unknown>;
}

// Original console methods for forwarding
const originalConsole = {
  log: console.log.bind(console),
  warn: console.warn.bind(console),
  error: console.error.bind(console),
};

export class BridgeClient {
  private ws: WebSocket | null = null;
  private connectionState: ConnectionState = 'disconnected';
  private reconnectAttempts = 0;
  private reconnectTimeout: ReturnType<typeof setTimeout> | null = null;
  private sessionId: string | null = null;
  private executeHandler: ExecuteHandler | null = null;

  private readonly app: AppType;
  private readonly getFilename: () => string;
  private readonly getUrl: () => string | null;
  private readonly onStateChange?: (state: ConnectionState) => void;
  private readonly onRegistered?: (id: string) => void;
  private readonly onActivity?: (action: string) => void;

  constructor(options: BridgeClientOptions) {
    this.app = options.app;
    this.getFilename = options.getFilename;
    this.getUrl = options.getUrl;
    this.onStateChange = options.onStateChange;
    this.onRegistered = options.onRegistered;
    this.onActivity = options.onActivity;
  }

  get state(): ConnectionState {
    return this.connectionState;
  }

  get id(): string | null {
    return this.sessionId;
  }

  /**
   * Set the handler for executing code from the bridge.
   */
  setExecuteHandler(handler: ExecuteHandler): void {
    this.executeHandler = handler;
  }

  /**
   * Connect to the bridge server.
   */
  connect(): void {
    this.updateState('connecting');

    try {
      this.ws = new WebSocket('wss://localhost:3847');

      this.ws.onopen = () => {
        this.connectionState = 'connected';
        this.reconnectAttempts = 0;
        this.updateState('connected');
        this.onActivity?.('Connected');

        // Send registration message with app type
        this.ws!.send(
          JSON.stringify({
            type: 'register',
            app: this.app,
            filename: this.getFilename(),
            url: this.getUrl(),
          })
        );
      };

      this.ws.onmessage = async (event) => {
        try {
          const msg = JSON.parse(event.data as string);
          await this.handleMessage(msg);
        } catch (err) {
          originalConsole.error('Failed to handle message:', err);
        }
      };

      this.ws.onclose = () => {
        this.ws = null;
        if (this.connectionState === 'connected') {
          // Unexpected close, try to reconnect
          this.scheduleReconnect();
        } else {
          this.updateState('disconnected');
        }
      };

      this.ws.onerror = (err) => {
        originalConsole.error('WebSocket error:', err);
        this.onActivity?.('Connection error');
      };
    } catch (err) {
      originalConsole.error('Failed to connect:', err);
      this.updateState('disconnected');
      this.onActivity?.('Connection failed');
    }
  }

  /**
   * Disconnect from the bridge server.
   */
  disconnect(): void {
    if (this.reconnectTimeout) {
      clearTimeout(this.reconnectTimeout);
      this.reconnectTimeout = null;
    }
    this.reconnectAttempts = 0;

    if (this.ws) {
      this.ws.close();
      this.ws = null;
    }

    this.sessionId = null;
    this.updateState('disconnected');
    this.onActivity?.('Disconnected');
  }

  /**
   * Send a console message to the bridge.
   */
  sendConsoleMessage(level: string, args: unknown[]): void {
    if (this.ws && this.ws.readyState === WebSocket.OPEN) {
      try {
        const message = args
          .map((arg) => (typeof arg === 'object' ? JSON.stringify(arg) : String(arg)))
          .join(' ');
        this.ws.send(
          JSON.stringify({
            type: 'console',
            level,
            message,
          })
        );
      } catch {
        // Ignore serialization errors
      }
    }
  }

  /**
   * Set up console forwarding to the bridge.
   */
  setupConsoleForwarding(): void {
    console.log = (...args: unknown[]) => {
      originalConsole.log(...args);
      this.sendConsoleMessage('log', args);
    };

    console.warn = (...args: unknown[]) => {
      originalConsole.warn(...args);
      this.sendConsoleMessage('warn', args);
    };

    console.error = (...args: unknown[]) => {
      originalConsole.error(...args);
      this.sendConsoleMessage('error', args);
    };
  }

  private updateState(state: ConnectionState): void {
    this.connectionState = state;
    this.onStateChange?.(state);
  }

  private scheduleReconnect(): void {
    if (this.reconnectTimeout) return;

    // Exponential backoff: 1s, 2s, 4s, 8s, 16s, 30s max
    const delay = Math.min(1000 * Math.pow(2, this.reconnectAttempts), 30000);
    this.reconnectAttempts++;

    this.updateState('connecting');
    this.onActivity?.(`Reconnecting in ${delay / 1000}s...`);

    this.reconnectTimeout = setTimeout(() => {
      this.reconnectTimeout = null;
      this.connect();
    }, delay);
  }

  private async handleMessage(msg: {
    type: string;
    id?: string;
    payload?: { id?: string; code?: string };
  }): Promise<void> {
    switch (msg.type) {
      case 'registered':
        this.sessionId = msg.payload?.id || null;
        if (this.sessionId) {
          this.onRegistered?.(this.sessionId);
        }
        this.onActivity?.('Registered');
        break;

      case 'execute':
        if (msg.id && msg.payload?.code && this.executeHandler) {
          this.onActivity?.('Executing code...');
          await this.executeCode(msg.id, msg.payload.code);
        }
        break;

      case 'ping':
        if (this.ws && this.ws.readyState === WebSocket.OPEN) {
          this.ws.send(JSON.stringify({ type: 'pong' }));
        }
        break;

      default:
        originalConsole.log('Unknown message type:', msg.type);
    }
  }

  private async executeCode(requestId: string, code: string): Promise<void> {
    if (!this.executeHandler) {
      this.sendError(requestId, 'No execute handler registered');
      return;
    }

    try {
      const result = await this.executeHandler(code);

      if (this.ws && this.ws.readyState === WebSocket.OPEN) {
        this.ws.send(
          JSON.stringify({
            type: 'result',
            id: requestId,
            success: true,
            result: this.serializeResult(result),
          })
        );
      }
      this.onActivity?.('Code executed successfully');
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : String(err);
      const errorStack = err instanceof Error ? err.stack : undefined;

      this.sendError(requestId, errorMessage, errorStack);
      this.onActivity?.('Code execution failed');
      originalConsole.error('Execution error:', err);
    }
  }

  private sendError(requestId: string, message: string, stack?: string): void {
    if (this.ws && this.ws.readyState === WebSocket.OPEN) {
      this.ws.send(
        JSON.stringify({
          type: 'result',
          id: requestId,
          success: false,
          error: { message, stack },
        })
      );
    }
  }

  private serializeResult(result: unknown): unknown {
    if (result === undefined) return null;
    if (result === null) return null;
    if (typeof result === 'string' || typeof result === 'number' || typeof result === 'boolean') {
      return result;
    }
    if (Array.isArray(result)) {
      return result.map((item) => this.serializeResult(item));
    }
    if (typeof result === 'object') {
      try {
        const obj: Record<string, unknown> = {};
        for (const key of Object.keys(result)) {
          obj[key] = this.serializeResult((result as Record<string, unknown>)[key]);
        }
        return obj;
      } catch {
        return String(result);
      }
    }
    return String(result);
  }
}
