const { spawn } = require('node:child_process');
const path = require('node:path');

function main() {
  const electronBinary = require('electron');
  const child = spawn(electronBinary, ['.'], {
    cwd: path.resolve(__dirname, '..'),
    stdio: 'inherit',
    env: process.env
  });

  child.on('exit', (code) => {
    process.exit(code ?? 0);
  });

  child.on('error', (error) => {
    console.error('[start-electron] failed to launch Electron');
    console.error(error);
    process.exit(1);
  });
}

main();
