const { spawn } = require('node:child_process');
const path = require('node:path');
const waitOn = require('wait-on');

const rootDir = path.resolve(__dirname, '..');
const devServerUrl = 'http://127.0.0.1:5173';

let viteProcess = null;
let electronProcess = null;

function stopChild(child) {
  if (!child || child.killed) {
    return;
  }

  child.kill();
}

async function main() {
  const viteEntry = path.join(rootDir, 'node_modules', 'vite', 'bin', 'vite.js');

  console.log(`[dev] starting Vite on ${devServerUrl}`);
  viteProcess = spawn(process.execPath, [viteEntry, '--configLoader', 'native', '--host', '127.0.0.1', '--port', '5173', '--strictPort'], {
    cwd: rootDir,
    stdio: 'inherit',
    shell: false,
    env: process.env
  });

  viteProcess.on('error', (error) => {
    console.error('[dev] failed to start Vite');
    console.error(error);
    process.exit(1);
  });

  await waitOn({
    resources: [devServerUrl],
    timeout: 60_000,
    validateStatus: (status) => status >= 200 && status < 500
  });

  console.log('[dev] launching Electron');
  const electronBinary = require('electron');
  electronProcess = spawn(electronBinary, ['.'], {
    cwd: rootDir,
    stdio: 'inherit',
    shell: false,
    env: {
      ...process.env,
      VITE_DEV_SERVER_URL: devServerUrl
    }
  });

  electronProcess.on('error', (error) => {
    console.error('[dev] failed to launch Electron');
    console.error(error);
    stopChild(viteProcess);
    process.exit(1);
  });

  electronProcess.on('exit', (code) => {
    stopChild(viteProcess);
    process.exit(code ?? 0);
  });
}

process.on('SIGINT', () => {
  stopChild(electronProcess);
  stopChild(viteProcess);
  process.exit(0);
});

process.on('SIGTERM', () => {
  stopChild(electronProcess);
  stopChild(viteProcess);
  process.exit(0);
});

main().catch((error) => {
  console.error('[dev] startup failed');
  console.error(error);
  stopChild(electronProcess);
  stopChild(viteProcess);
  process.exit(1);
});
