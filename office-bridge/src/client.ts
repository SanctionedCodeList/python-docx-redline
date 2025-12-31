import type { DocumentInfo, ServerInfo } from './types.js';

export interface ConnectOptions {
  port?: number;  // Default: 3847
  host?: string;  // Default: 'localhost'
}

export interface ExecuteOptions {
  timeout?: number;  // Default: 30000 (30 seconds)
}

export interface ConsoleEntry {
  level: 'log' | 'warn' | 'error' | 'info';
  message: string;
  timestamp: Date;
  documentId: string;
}

export interface WordDocument {
  readonly id: string;
  readonly filename: string;
  readonly path: string;
  readonly connectedAt: Date;
  readonly status: 'connected' | 'disconnected';

  executeJs<T = unknown>(code: string, options?: ExecuteOptions): Promise<T>;
}

export interface BridgeClient {
  documents(): Promise<WordDocument[]>;
  close(): Promise<void>;
}

export async function connect(options: ConnectOptions = {}): Promise<BridgeClient> {
  const host = options.host ?? 'localhost';
  const port = options.port ?? 3847;
  const baseUrl = `https://${host}:${port}`;

  // Fetch server info to verify connection
  const infoRes = await fetch(baseUrl);
  if (!infoRes.ok) {
    throw new Error(`Failed to connect to bridge server: ${infoRes.status}`);
  }
  const serverInfo = await infoRes.json() as ServerInfo;
  console.log(`Connected to Office Bridge (${serverInfo.documents} documents)`);

  return {
    async documents(): Promise<WordDocument[]> {
      const res = await fetch(`${baseUrl}/documents`);
      if (!res.ok) {
        throw new Error(`Failed to list documents: ${res.status}`);
      }
      const data = await res.json() as { documents: DocumentInfo[] };

      return data.documents.map(doc => createDocumentHandle(baseUrl, doc));
    },

    async close(): Promise<void> {
      // Nothing to clean up for HTTP client
      console.log('Disconnected from Office Bridge');
    },
  };
}

function createDocumentHandle(baseUrl: string, info: DocumentInfo): WordDocument {
  return {
    id: info.id,
    filename: info.filename,
    path: info.path,
    connectedAt: new Date(info.connectedAt),
    status: info.status,

    async executeJs<T = unknown>(code: string, options: ExecuteOptions = {}): Promise<T> {
      const timeout = options.timeout ?? 30000;

      const res = await fetch(`${baseUrl}/execute`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          documentId: info.id,
          code,
          timeout,
        }),
      });

      if (!res.ok) {
        const error = await res.text();
        throw new Error(`Execution failed: ${error}`);
      }

      const result = await res.json() as { success: boolean; result?: T; error?: { message: string } };

      if (!result.success) {
        throw new Error(result.error?.message ?? 'Execution failed');
      }

      return result.result as T;
    },
  };
}
