const path = require('node:path');
const { spawnSync } = require('node:child_process');

const script = path.join(__dirname, 'generate-ai-industry-report-full-native.js');
const result = spawnSync(process.execPath, [script], { stdio: 'inherit' });

if (result.error) {
  console.error(result.error);
  process.exitCode = 1;
} else {
  process.exitCode = result.status === null ? 1 : result.status;
}
