const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const DEFAULT_TIMEOUT_MS = 120000;
const POWERPOINT_COM_MODES = Object.freeze(['auto', 'native', 'powerpoint-com']);
const SCRIPT_ROOT = path.resolve(__dirname);

class PowerPointComError extends Error {
  constructor(message, options = {}) {
    super(message);
    this.name = 'PowerPointComError';
    if (options.cause) this.cause = options.cause;
    if (options.code) this.code = options.code;
  }
}

function normalizePowerPointComConfig(config = {}) {
  const mode = String(config.mode || 'auto');
  if (!POWERPOINT_COM_MODES.includes(mode)) {
    throw new PowerPointComError(`Unsupported PowerPoint backend mode: ${mode}`);
  }
  const timeoutMs = config.timeout_ms ?? config.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  if (!Number.isInteger(timeoutMs) || timeoutMs < 1000) {
    throw new PowerPointComError('PowerPoint COM timeout_ms must be an integer of at least 1000 milliseconds.');
  }
  return {
    mode,
    visible: config.visible === true,
    timeout_ms: timeoutMs,
    display_alerts: config.display_alerts === true || config.displayAlerts === true
  };
}

function selectPresentationBackend({ mode = 'auto', platform = process.platform, comAvailable = false } = {}) {
  if (!POWERPOINT_COM_MODES.includes(mode)) {
    throw new PowerPointComError(`Unsupported PowerPoint backend mode: ${mode}`);
  }
  if (mode === 'native') return 'native';
  if (mode === 'powerpoint-com') {
    if (platform !== 'win32') {
      throw new PowerPointComError('PowerPoint COM requires Windows and an installed desktop version of Microsoft PowerPoint.');
    }
    if (!comAvailable) {
      throw new PowerPointComError('PowerPoint COM is unavailable. Install desktop Microsoft PowerPoint and ensure the PowerPoint.Application COM class is registered.');
    }
    return 'powerpoint-com';
  }
  return platform === 'win32' && comAvailable ? 'powerpoint-com' : 'native';
}

function defaultRunner(executable, args, options) {
  return spawnSync(executable, args, options);
}

function probePowerPointCom({
  platform = process.platform,
  timeoutMs = DEFAULT_TIMEOUT_MS,
  runner = defaultRunner
} = {}) {
  if (platform !== 'win32') {
    return { available: false, reason: 'PowerPoint COM requires Windows.' };
  }
  const command = [
    "$ErrorActionPreference = 'Stop'",
    '$app = $null',
    'try {',
    "  $app = New-Object -ComObject PowerPoint.Application",
    "  $app.Visible = -1",
    "  $app.WindowState = 2",
    "  $app.DisplayAlerts = 1",
    "  Write-Output 'PowerPoint COM available'",
    '}',
    'finally {',
    '  if ($null -ne $app) { try { $app.Quit() } catch {} }',
    '}'
  ].join('\n');
  try {
    const result = runner('powershell.exe', [
      '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', command
    ], { encoding: 'utf8', timeout: timeoutMs, windowsHide: true });
    if (result?.status === 0) return { available: true, reason: null };
    return {
      available: false,
      reason: String(result?.stderr || result?.stdout || 'PowerPoint COM probe failed.').trim()
    };
  } catch (error) {
    return { available: false, reason: error.message || 'PowerPoint COM probe failed.' };
  }
}

function validateScriptName(scriptName) {
  if (
    typeof scriptName !== 'string' ||
    !scriptName ||
    scriptName.includes('/') ||
    scriptName.includes('\\') ||
    path.basename(scriptName) !== scriptName ||
    !scriptName.toLowerCase().endsWith('.ps1')
  ) {
    throw new PowerPointComError('PowerPoint COM script name must be a single .ps1 filename.');
  }
  return scriptName;
}

function runPowerPointComScript(scriptName, args = [], {
  platform = process.platform,
  timeoutMs = DEFAULT_TIMEOUT_MS,
  visible = false,
  displayAlerts = false,
  runner = defaultRunner
} = {}) {
  const safeScriptName = validateScriptName(scriptName);
  if (platform !== 'win32') {
    throw new PowerPointComError('PowerPoint COM requires Windows and an installed desktop version of Microsoft PowerPoint.');
  }
  if (!Array.isArray(args)) throw new PowerPointComError('PowerPoint COM script arguments must be an array.');
  const scriptPath = path.resolve(SCRIPT_ROOT, safeScriptName);
  if (!fs.existsSync(scriptPath)) {
    throw new PowerPointComError(`PowerPoint COM script is not bundled: ${safeScriptName}`);
  }
  let result;
  try {
    result = runner('powershell.exe', [
      '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath,
      ...args.map((arg) => String(arg))
    ], {
      encoding: 'utf8',
      timeout: timeoutMs,
      windowsHide: true,
      env: {
        ...process.env,
        PPT_CREATER_POWERPOINT_VISIBLE: visible ? '1' : '0',
        PPT_CREATER_POWERPOINT_DISPLAY_ALERTS: displayAlerts ? '1' : '0'
      }
    });
  } catch (error) {
    throw new PowerPointComError(`PowerPoint COM script could not start: ${error.message || error}`, { cause: error });
  }
  if (result?.status === null) {
    throw new PowerPointComError(`PowerPoint COM script timed out after ${timeoutMs} ms.`);
  }
  if (result?.error) throw new PowerPointComError(`PowerPoint COM script failed: ${result.error.message}`, { cause: result.error });
  if (result?.status !== 0) {
    const detail = String(result?.stderr || result?.stdout || 'unknown error').trim();
    throw new PowerPointComError(`PowerPoint COM script failed: ${detail}`);
  }
  return result;
}

function resolvePowerPointComConfig(config = {}) {
  const source = config.powerpoint_com || config.powerpointCom || config;
  return normalizePowerPointComConfig({
    ...source,
    mode: source.mode || config.backend || 'auto'
  });
}

function getPowerPointComStatus({
  config = {},
  platform = process.platform,
  probe = true,
  runner = defaultRunner
} = {}) {
  const normalized = resolvePowerPointComConfig(config);
  if (probe === false) {
    return {
      configured_mode: normalized.mode,
      selected_backend: normalized.mode === 'native' ? 'native' : null,
      available: null,
      reason: 'PowerPoint COM probe not requested.',
      probe_performed: false,
      platform,
      visible: normalized.visible,
      timeout_ms: normalized.timeout_ms,
      display_alerts: normalized.display_alerts
    };
  }
  const probeResult = probePowerPointCom({
    platform,
    timeoutMs: normalized.timeout_ms,
    runner
  });
  let selectedBackend = null;
  let selectionError = null;
  try {
    selectedBackend = selectPresentationBackend({
      mode: normalized.mode,
      platform,
      comAvailable: probeResult.available
    });
  } catch (error) {
    selectionError = error.message;
  }
  return {
    configured_mode: normalized.mode,
    selected_backend: selectedBackend,
    available: probeResult.available,
    reason: selectionError || probeResult.reason,
    probe_performed: true,
    platform,
    visible: normalized.visible,
    timeout_ms: normalized.timeout_ms,
    display_alerts: normalized.display_alerts
  };
}

class PowerPointComBackend {
  constructor(config = {}, dependencies = {}) {
    this.config = resolvePowerPointComConfig(config);
    this.platform = dependencies.platform || process.platform;
    this.runner = dependencies.runner || defaultRunner;
  }

  probe() {
    return probePowerPointCom({
      platform: this.platform,
      timeoutMs: this.config.timeout_ms,
      runner: this.runner
    });
  }

  status() {
    return getPowerPointComStatus({
      config: this.config,
      platform: this.platform,
      runner: this.runner
    });
  }

  runScript(scriptName, args = []) {
    return runPowerPointComScript(scriptName, args, {
      platform: this.platform,
      timeoutMs: this.config.timeout_ms,
      visible: this.config.visible,
      displayAlerts: this.config.display_alerts,
      runner: this.runner
    });
  }
}

module.exports = {
  DEFAULT_TIMEOUT_MS,
  POWERPOINT_COM_MODES,
  PowerPointComError,
  normalizePowerPointComConfig,
  resolvePowerPointComConfig,
  selectPresentationBackend,
  probePowerPointCom,
  getPowerPointComStatus,
  PowerPointComBackend,
  runPowerPointComScript
};
