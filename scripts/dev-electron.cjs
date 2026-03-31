const { spawn } = require('node:child_process');
const path = require('node:path');
const waitOn = require('wait-on');

async function main() {
  const devServerUrl = 'http://127.0.0.1:5173';

  console.log(`[dev-electron] waiting for ${devServerUrl}`);
  await waitOn({
    resources: [devServerUrl],
    timeout: 60_000,
    validateStatus: (status) => status >= 200 && status < 500
  });

  console.log('[dev-electron] launching Electron');
  const electronBinary = require('electron');
  const child = spawn(electronBinary, ['.'], {
    cwd: path.resolve(__dirname, '..'),
    stdio: 'inherit',
    env: {
      ...process.env,
      VITE_DEV_SERVER_URL: devServerUrl
    }
  });

  child.on('exit', (code) => {
    process.exit(code ?? 0);
  });

  child.on('error', (error) => {
    console.error('[dev-electron] failed to launch Electron');
    console.error(error);
    process.exit(1);
  });
}

main().catch((error) => {
  console.error('[dev-electron] startup failed');
  console.error(error);
  process.exit(1);
});
