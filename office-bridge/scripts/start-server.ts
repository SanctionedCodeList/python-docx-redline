import { execSync } from 'child_process';
import { mkdirSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';
import { serve } from '../src/index.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const rootDir = join(__dirname, '..');

// Create tmp directory
mkdirSync(join(rootDir, 'tmp'), { recursive: true });

// Kill existing process on port 3847
const port = 3847;
try {
  const pid = execSync(`lsof -ti:${port}`, { encoding: 'utf8' }).trim();
  if (pid) {
    console.log(`Killing existing process on port ${port} (PID: ${pid})`);
    execSync(`kill -9 ${pid}`);
  }
} catch {
  // No process on port, that's fine
}

// Start server
const server = await serve({ port });

// Handle graceful shutdown
const shutdown = async () => {
  console.log('\nShutting down...');
  await server.close();
  process.exit(0);
};

process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);
process.on('SIGHUP', shutdown);

console.log('Press Ctrl+C to stop\n');
