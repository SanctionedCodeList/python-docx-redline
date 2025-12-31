import type { WebSocket } from 'ws';
import type { ConnectedDocument, DocumentInfo } from './types.js';

// Pending execution in the queue
interface PendingExecution {
  requestId: string;
  code: string;
  timeout: number;
  resolve: (result: unknown) => void;
  reject: (error: Error) => void;
}

export class DocumentRegistry {
  private documents = new Map<string, ConnectedDocument>();
  private socketToId = new Map<WebSocket, string>();
  private idToSocket = new Map<string, WebSocket>();
  private executionQueues = new Map<string, PendingExecution[]>();
  private activeExecutions = new Map<string, PendingExecution>();

  private readonly tombstoneTimeout = 5 * 60 * 1000; // 5 minutes

  register(id: string, ws: WebSocket, filename: string, path: string): void {
    const doc: ConnectedDocument = {
      id,
      filename,
      path,
      connectedAt: new Date(),
      lastActivity: new Date(),
      status: 'connected',
    };

    this.documents.set(id, doc);
    this.socketToId.set(ws, id);
    this.idToSocket.set(id, ws);
    this.executionQueues.set(id, []);
  }

  unregister(ws: WebSocket): ConnectedDocument | undefined {
    const id = this.socketToId.get(ws);
    if (!id) return undefined;

    const doc = this.documents.get(id);
    if (doc) {
      doc.status = 'disconnected';
      doc.disconnectedAt = new Date();

      // Schedule tombstone cleanup
      setTimeout(() => {
        const current = this.documents.get(id);
        if (current?.status === 'disconnected') {
          this.documents.delete(id);
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
      active.reject(new Error('Document disconnected'));
      this.activeExecutions.delete(id);
    }

    for (const pending of queue) {
      pending.reject(new Error('Document disconnected'));
    }
    this.executionQueues.set(id, []);

    return doc;
  }

  get(id: string): ConnectedDocument | undefined {
    return this.documents.get(id);
  }

  getSocket(id: string): WebSocket | undefined {
    return this.idToSocket.get(id);
  }

  getIdBySocket(ws: WebSocket): string | undefined {
    return this.socketToId.get(ws);
  }

  list(): DocumentInfo[] {
    return Array.from(this.documents.values()).map((doc) => ({
      id: doc.id,
      filename: doc.filename,
      path: doc.path,
      connectedAt: doc.connectedAt.toISOString(),
      lastActivity: doc.lastActivity.toISOString(),
      status: doc.status,
    }));
  }

  listConnected(): DocumentInfo[] {
    return this.list().filter((doc) => doc.status === 'connected');
  }

  updateActivity(id: string): void {
    const doc = this.documents.get(id);
    if (doc) {
      doc.lastActivity = new Date();
    }
  }

  // Queue an execution for a document
  async queueExecution(
    id: string,
    requestId: string,
    code: string,
    timeout: number
  ): Promise<unknown> {
    const doc = this.documents.get(id);
    if (!doc || doc.status !== 'connected') {
      throw new Error(`Document ${id} not connected`);
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
}
