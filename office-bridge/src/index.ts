import express from 'express';
import { createServer } from 'https';
import { readFileSync } from 'fs';
import { homedir } from 'os';
import { join } from 'path';
import { WebSocketServer, WebSocket } from 'ws';
import { v4 as uuidv4 } from 'uuid';
import { DocumentRegistry } from './registry.js';
import type { AddInMessage, ServerInfo, ResultMessage, ConsoleMessage } from './types.js';

export interface ServerOptions {
  port?: number;
}

export interface BridgeServer {
  port: number;
  registry: DocumentRegistry;
  close: () => Promise<void>;
}

export async function serve(options: ServerOptions = {}): Promise<BridgeServer> {
  const port = options.port ?? 3847;
  const registry = new DocumentRegistry();

  const app = express();
  app.use(express.json());

  // Server info endpoint
  app.get('/', (_req, res) => {
    const info: ServerInfo = {
      port,
      documents: registry.listConnected().length,
    };
    res.json(info);
  });

  // List documents endpoint
  app.get('/documents', (_req, res) => {
    res.json({ documents: registry.listConnected() });
  });

  // Execute code in a document
  app.post('/execute', async (req, res) => {
    const { documentId, code, timeout = 30000 } = req.body as {
      documentId: string;
      code: string;
      timeout?: number;
    };

    try {
      const result = await registry.queueExecution(documentId, uuidv4(), code, timeout);
      res.json({ success: true, result });
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Unknown error';
      res.status(500).json({ success: false, error: { message } });
    }
  });

  // Use the same dev certificates as the add-in
  const certDir = join(homedir(), '.office-addin-dev-certs');
  const server = createServer({
    key: readFileSync(join(certDir, 'localhost.key')),
    cert: readFileSync(join(certDir, 'localhost.crt')),
  }, app);
  const wss = new WebSocketServer({ server });

  // Console output handlers (for forwarding to clients)
  const consoleHandlers = new Map<string, Set<(entry: ConsoleMessage['payload']) => void>>();

  wss.on('connection', (ws: WebSocket) => {
    console.log('Add-in connected');

    ws.on('message', (data: Buffer) => {
      try {
        const message = JSON.parse(data.toString()) as AddInMessage;

        switch (message.type) {
          case 'register': {
            // No token validation - localhost only
            const registerMsg = message as { type: 'register'; document: { filename: string; url?: string } };

            const id = uuidv4();
            registry.register(id, ws, registerMsg.document.filename, registerMsg.document.url);

            ws.send(JSON.stringify({
              type: 'registered',
              payload: { id },
            }));

            console.log(`Document registered: ${registerMsg.document.filename} (${id})`);
            break;
          }

          case 'result': {
            // Add-in sends { type, id, success, result?, error? } without payload wrapper
            const resultMsg = message as { type: 'result'; id: string; success: boolean; result?: unknown; error?: { message: string; stack?: string } };
            const docId = registry.getIdBySocket(ws);
            if (docId) {
              const error = resultMsg.error
                ? new Error(resultMsg.error.message)
                : undefined;
              registry.handleResult(
                docId,
                resultMsg.id,
                resultMsg.success,
                resultMsg.result,
                error
              );
            }
            break;
          }

          case 'console': {
            // Add-in sends { type, level, message } without payload wrapper
            const consoleMsg = message as { type: 'console'; level: string; message: string };
            const docId = registry.getIdBySocket(ws);
            if (docId) {
              const handlers = consoleHandlers.get(docId);
              if (handlers) {
                for (const handler of handlers) {
                  handler({ level: consoleMsg.level, message: consoleMsg.message });
                }
              }
              // Also log to server console
              console.log(`[${docId}] ${consoleMsg.level}: ${consoleMsg.message}`);
            }
            break;
          }
        }
      } catch (err) {
        console.error('Failed to parse message:', err);
      }
    });

    ws.on('close', () => {
      const doc = registry.unregister(ws);
      if (doc) {
        console.log(`Document disconnected: ${doc.filename} (${doc.id})`);
      }
    });

    ws.on('error', (err) => {
      console.error('WebSocket error:', err);
    });
  });

  // Start server
  await new Promise<void>((resolve) => {
    server.listen(port, () => {
      resolve();
    });
  });

  console.log('');
  console.log('=================================');
  console.log('  Office Bridge Server Started');
  console.log('=================================');
  console.log(`  Port: ${port}`);
  console.log('  Auth: localhost only (no token)');
  console.log('=================================');
  console.log('');

  return {
    port,
    registry,
    close: async () => {
      return new Promise((resolve, reject) => {
        wss.close();
        server.close((err) => {
          if (err) reject(err);
          else resolve();
        });
      });
    },
  };
}
