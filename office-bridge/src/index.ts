import express from 'express';
import { createServer } from 'https';
import { readFileSync } from 'fs';
import { homedir } from 'os';
import { join } from 'path';
import { WebSocketServer, WebSocket } from 'ws';
import { v4 as uuidv4 } from 'uuid';
import { pdf } from 'pdf-to-img';
import { SessionRegistry } from './registry.js';
import type { AddInMessage, AppType, ServerInfo, ConsoleMessage } from './types.js';

export interface ServerOptions {
  port?: number;
}

export interface BridgeServer {
  port: number;
  registry: SessionRegistry;
  close: () => Promise<void>;
}

export async function serve(options: ServerOptions = {}): Promise<BridgeServer> {
  const port = options.port ?? 3847;
  const registry = new SessionRegistry();

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

  // List sessions endpoint (with optional app filter)
  app.get('/sessions', (req, res) => {
    const app = req.query.app as AppType | undefined;
    res.json({ sessions: registry.listConnected(app) });
  });

  // Backward compatibility: /documents endpoint
  app.get('/documents', (_req, res) => {
    res.json({ documents: registry.listConnected('word') });
  });

  // Execute code in a session
  app.post('/execute', async (req, res) => {
    // Support both sessionId and documentId for backward compatibility
    const { sessionId, documentId, code, timeout = 30000 } = req.body as {
      sessionId?: string;
      documentId?: string;
      code: string;
      timeout?: number;
    };
    const id = sessionId || documentId;

    if (!id) {
      res.status(400).json({ success: false, error: { message: 'sessionId is required' } });
      return;
    }

    try {
      const result = await registry.queueExecution(id, uuidv4(), code, timeout);
      res.json({ success: true, result });
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Unknown error';
      res.status(500).json({ success: false, error: { message } });
    }
  });

  // Get page images from a Word document
  app.post('/page-images', async (req, res) => {
    const { sessionId, scale = 1.5, pages } = req.body as {
      sessionId: string;
      scale?: number;
      pages?: number[];
    };

    if (!sessionId) {
      res.status(400).json({ success: false, error: { message: 'sessionId is required' } });
      return;
    }

    try {
      // Request PDF from the add-in
      const pdfBase64 = await registry.queuePdfExport(sessionId, uuidv4(), 60000);

      // Convert base64 to buffer
      const pdfBuffer = Buffer.from(pdfBase64, 'base64');

      // Convert PDF to images
      const images: { page: number; width: number; height: number; data: string }[] = [];
      let pageNum = 0;

      // pdf() returns a Promise that resolves to an async iterable
      const pdfDoc = await pdf(pdfBuffer, { scale });
      for await (const image of pdfDoc) {
        pageNum++;
        // Filter by requested pages if specified
        if (pages && !pages.includes(pageNum)) {
          continue;
        }

        // image is a Buffer of PNG data
        const base64 = image.toString('base64');
        images.push({
          page: pageNum,
          width: 0, // pdf-to-img doesn't provide dimensions directly
          height: 0,
          data: `data:image/png;base64,${base64}`,
        });
      }

      res.json({ success: true, images, totalPages: pageNum });
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
            // Support both new format (with app) and legacy format (document object)
            const registerMsg = message as {
              type: 'register';
              app?: AppType;
              document?: { filename: string; url?: string };
              filename?: string;
              url?: string;
            };

            // Extract app type (default to 'word' for backward compatibility)
            const app: AppType = registerMsg.app || 'word';
            const filename = registerMsg.document?.filename || registerMsg.filename || 'Untitled';
            const url = registerMsg.document?.url || registerMsg.url;

            const id = uuidv4();
            registry.register(id, ws, app, filename, url || '');

            ws.send(JSON.stringify({
              type: 'registered',
              payload: { id },
            }));

            console.log(`[${app}] Session registered: ${filename} (${id})`);
            break;
          }

          case 'result': {
            // Add-in sends { type, id, success, result?, error? } without payload wrapper
            const resultMsg = message as unknown as { type: 'result'; id: string; success: boolean; result?: unknown; error?: { message: string; stack?: string } };
            const sessionId = registry.getIdBySocket(ws);
            if (sessionId) {
              const error = resultMsg.error
                ? new Error(resultMsg.error.message)
                : undefined;
              registry.handleResult(
                sessionId,
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
            const consoleMsg = message as unknown as { type: 'console'; level: string; message: string };
            const sessionId = registry.getIdBySocket(ws);
            if (sessionId) {
              const handlers = consoleHandlers.get(sessionId);
              if (handlers) {
                const level = consoleMsg.level as 'log' | 'warn' | 'error' | 'info';
                for (const handler of handlers) {
                  handler({ level, message: consoleMsg.message, timestamp: new Date().toISOString() });
                }
              }
              // Also log to server console
              console.log(`[${sessionId}] ${consoleMsg.level}: ${consoleMsg.message}`);
            }
            break;
          }

          case 'pdfResult': {
            // Add-in sends PDF export result
            const pdfMsg = message as unknown as { type: 'pdfResult'; id: string; success: boolean; pdfBase64?: string; error?: { message: string } };
            const sessionId = registry.getIdBySocket(ws);
            if (sessionId) {
              const error = pdfMsg.error
                ? new Error(pdfMsg.error.message)
                : undefined;
              registry.handlePdfResult(
                sessionId,
                pdfMsg.id,
                pdfMsg.success,
                pdfMsg.pdfBase64,
                error
              );
            }
            break;
          }
        }
      } catch (err) {
        console.error('Failed to parse message:', err);
      }
    });

    ws.on('close', () => {
      const session = registry.unregister(ws);
      if (session) {
        console.log(`[${session.app}] Session disconnected: ${session.filename} (${session.id})`);
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
  console.log('=========================================');
  console.log('  Office Bridge Server Started');
  console.log('=========================================');
  console.log(`  Port: ${port}`);
  console.log('  Apps: Word, Excel, PowerPoint, Outlook');
  console.log('  Auth: localhost only (no token)');
  console.log('=========================================');
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
