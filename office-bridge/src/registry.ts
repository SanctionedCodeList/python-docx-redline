import type { WebSocket } from 'ws';
import type { AppType, ConnectedSession, SessionInfo } from './types.js';

// Pending execution in the queue
interface PendingExecution {
  requestId: string;
  code: string;
  timeout: number;
  resolve: (result: unknown) => void;
  reject: (error: Error) => void;
}

// Pending PDF export request
interface PendingPdfExport {
  requestId: string;
  timeout: number;
  resolve: (pdfBase64: string) => void;
  reject: (error: Error) => void;
}

export class SessionRegistry {
  private sessions = new Map<string, ConnectedSession>();
  private socketToId = new Map<WebSocket, string>();
  private idToSocket = new Map<string, WebSocket>();
  private executionQueues = new Map<string, PendingExecution[]>();
  private activeExecutions = new Map<string, PendingExecution>();
  private pendingPdfExports = new Map<string, PendingPdfExport>();

  private readonly tombstoneTimeout = 5 * 60 * 1000; // 5 minutes

  register(id: string, ws: WebSocket, app: AppType, filename: string, path: string): void {
    const session: ConnectedSession = {
      id,
      app,
      filename,
      path,
      connectedAt: new Date(),
      lastActivity: new Date(),
      status: 'connected',
    };

    this.sessions.set(id, session);
    this.socketToId.set(ws, id);
    this.idToSocket.set(id, ws);
    this.executionQueues.set(id, []);
  }

  unregister(ws: WebSocket): ConnectedSession | undefined {
    const id = this.socketToId.get(ws);
    if (!id) return undefined;

    const session = this.sessions.get(id);
    if (session) {
      session.status = 'disconnected';
      session.disconnectedAt = new Date();

      // Schedule tombstone cleanup
      setTimeout(() => {
        const current = this.sessions.get(id);
        if (current?.status === 'disconnected') {
          this.sessions.delete(id);
          this.executionQueues.delete(id);
        }
      }, this.tombstoneTimeout);
    }

    this.socketToId.delete(ws);
    this.idToSocket.delete(id);

    // Reject any pending executions
    const queue = this.executionQueues.get(id) || [];
    const active = this.activeExecutions.get(id);

    if (active) {
      active.reject(new Error('Session disconnected'));
      this.activeExecutions.delete(id);
    }

    for (const pending of queue) {
      pending.reject(new Error('Session disconnected'));
    }
    this.executionQueues.set(id, []);

    return session;
  }

  get(id: string): ConnectedSession | undefined {
    return this.sessions.get(id);
  }

  getSocket(id: string): WebSocket | undefined {
    return this.idToSocket.get(id);
  }

  getIdBySocket(ws: WebSocket): string | undefined {
    return this.socketToId.get(ws);
  }

  list(appFilter?: AppType): SessionInfo[] {
    return Array.from(this.sessions.values())
      .filter((session) => !appFilter || session.app === appFilter)
      .map((session) => ({
        id: session.id,
        app: session.app,
        filename: session.filename,
        path: session.path,
        connectedAt: session.connectedAt.toISOString(),
        lastActivity: session.lastActivity.toISOString(),
        status: session.status,
      }));
  }

  listConnected(appFilter?: AppType): SessionInfo[] {
    return this.list(appFilter).filter((session) => session.status === 'connected');
  }

  updateActivity(id: string): void {
    const session = this.sessions.get(id);
    if (session) {
      session.lastActivity = new Date();
    }
  }

  // Queue an execution for a session
  async queueExecution(
    id: string,
    requestId: string,
    code: string,
    timeout: number
  ): Promise<unknown> {
    const session = this.sessions.get(id);
    if (!session || session.status !== 'connected') {
      throw new Error(`Session ${id} not connected`);
    }

    return new Promise((resolve, reject) => {
      const pending: PendingExecution = { requestId, code, timeout, resolve, reject };
      const queue = this.executionQueues.get(id) || [];
      queue.push(pending);
      this.executionQueues.set(id, queue);

      // Try to process the queue
      this.processQueue(id);
    });
  }

  // Process the next item in a document's queue
  private processQueue(id: string): void {
    // If already executing, wait
    if (this.activeExecutions.has(id)) return;

    const queue = this.executionQueues.get(id);
    if (!queue || queue.length === 0) return;

    const pending = queue.shift()!;
    this.activeExecutions.set(id, pending);

    // Send execute message to add-in
    const ws = this.idToSocket.get(id);
    if (!ws) {
      pending.reject(new Error('WebSocket not found'));
      this.activeExecutions.delete(id);
      this.processQueue(id);
      return;
    }

    ws.send(
      JSON.stringify({
        type: 'execute',
        id: pending.requestId,
        payload: {
          code: pending.code,
          timeout: pending.timeout,
        },
      })
    );

    // Set up timeout
    setTimeout(() => {
      const active = this.activeExecutions.get(id);
      if (active?.requestId === pending.requestId) {
        active.reject(new Error('Execution timeout'));
        this.activeExecutions.delete(id);
        this.processQueue(id);
      }
    }, pending.timeout);
  }

  // Handle result from add-in
  handleResult(
    id: string,
    requestId: string,
    success: boolean,
    result?: unknown,
    error?: Error
  ): void {
    const active = this.activeExecutions.get(id);
    if (!active || active.requestId !== requestId) return;

    if (success) {
      active.resolve(result);
    } else {
      active.reject(error || new Error('Execution failed'));
    }

    this.activeExecutions.delete(id);
    this.updateActivity(id);
    this.processQueue(id);
  }

  // Queue a PDF export request
  async queuePdfExport(
    id: string,
    requestId: string,
    timeout: number
  ): Promise<string> {
    const session = this.sessions.get(id);
    if (!session || session.status !== 'connected') {
      throw new Error(`Session ${id} not connected`);
    }

    if (session.app !== 'word') {
      throw new Error('PDF export is only supported for Word documents');
    }

    const ws = this.idToSocket.get(id);
    if (!ws) {
      throw new Error('WebSocket not found');
    }

    return new Promise((resolve, reject) => {
      const pending: PendingPdfExport = { requestId, timeout, resolve, reject };
      this.pendingPdfExports.set(requestId, pending);

      // Send getPdf message to add-in
      ws.send(
        JSON.stringify({
          type: 'getPdf',
          id: requestId,
        })
      );

      // Set up timeout
      setTimeout(() => {
        const pendingExport = this.pendingPdfExports.get(requestId);
        if (pendingExport) {
          pendingExport.reject(new Error('PDF export timeout'));
          this.pendingPdfExports.delete(requestId);
        }
      }, timeout);
    });
  }

  // Handle PDF result from add-in
  handlePdfResult(
    id: string,
    requestId: string,
    success: boolean,
    pdfBase64?: string,
    error?: Error
  ): void {
    const pending = this.pendingPdfExports.get(requestId);
    if (!pending) return;

    if (success && pdfBase64) {
      pending.resolve(pdfBase64);
    } else {
      pending.reject(error || new Error('PDF export failed'));
    }

    this.pendingPdfExports.delete(requestId);
    this.updateActivity(id);
  }
}

// Backward compatibility alias
export const DocumentRegistry = SessionRegistry;
