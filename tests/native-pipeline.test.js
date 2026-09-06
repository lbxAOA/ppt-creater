const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const {
  NATIVE_PIPELINE_VERSION,
  assertNativePipeline,
  buildNativePipelineContract,
  optimizeTextForSlot,
  buildOptimizedPages,
  normalizeObjectFill,
  buildPreviewContract
} = require('../src/native-pipeline');
const {
  PowerPointComError,
  normalizePowerPointComConfig,
  selectPresentationBackend,
  probePowerPointCom,
  getPowerPointComStatus,
  PowerPointComBackend,
  runPowerPointComScript
} = require('../src/powerpoint-com');

test('PowerPoint COM config defaults to auto without loading COM', () => {
  assert.deepEqual(normalizePowerPointComConfig(), {
    mode: 'auto', visible: false, timeout_ms: 120000, display_alerts: false
  });
  assert.deepEqual(normalizePowerPointComConfig({
    mode: 'powerpoint-com', visible: true, timeout_ms: 5000, display_alerts: true
  }), {
    mode: 'powerpoint-com', visible: true, timeout_ms: 5000, display_alerts: true
  });
});

test('backend selection keeps auto portable and rejects unavailable forced COM', () => {
  assert.equal(selectPresentationBackend({ mode: 'auto', platform: 'linux', comAvailable: false }), 'native');
  assert.equal(selectPresentationBackend({ mode: 'auto', platform: 'win32', comAvailable: true }), 'powerpoint-com');
  assert.throws(
    () => selectPresentationBackend({ mode: 'powerpoint-com', platform: 'linux', comAvailable: false }),
    (error) => error instanceof PowerPointComError && /Windows/.test(error.message)
  );
  assert.throws(
    () => selectPresentationBackend({ mode: 'powerpoint-com', platform: 'win32', comAvailable: false }),
    (error) => error instanceof PowerPointComError && /desktop Microsoft PowerPoint/.test(error.message)
  );
});

test('COM probe is platform-gated and closes its probe process command', () => {
  let invocation;
  const unavailable = probePowerPointCom({ platform: 'linux', runner: (...args) => { invocation = args; } });
  assert.equal(unavailable.available, false);
  assert.match(unavailable.reason, /Windows/);
  assert.equal(invocation, undefined);

  const runner = (executable, args, options) => {
    invocation = { executable, args, options };
    return { status: 0, stdout: 'PowerPoint COM available\n', stderr: '' };
  };
  const available = probePowerPointCom({ platform: 'win32', runner });
  assert.equal(available.available, true);
  assert.equal(invocation.executable, 'powershell.exe');
  assert.ok(invocation.args.includes('-NonInteractive'));
  assert.ok(invocation.options.timeout > 0);
});

test('COM script runner rejects traversal and reports process failures', () => {
  assert.throws(() => runPowerPointComScript('../secret.ps1', [], { platform: 'win32', runner: () => ({ status: 0 }) }), /script name/);
  assert.throws(
    () => runPowerPointComScript('preview-all-pages.ps1', [], {
      platform: 'win32', runner: () => ({ status: 1, stdout: '', stderr: 'PowerPoint failed' })
    }),
    (error) => error instanceof PowerPointComError && /PowerPoint failed/.test(error.message)
  );
});

test('COM script runner normalizes runner exceptions and timeouts', () => {
  assert.throws(
    () => runPowerPointComScript('preview-all-pages.ps1', [], {
      platform: 'win32', runner: () => { throw new Error('spawn unavailable'); }
    }),
    (error) => error instanceof PowerPointComError && /could not start.*spawn unavailable/.test(error.message)
  );
  assert.throws(
    () => runPowerPointComScript('preview-all-pages.ps1', [], {
      platform: 'win32', timeoutMs: 1000, runner: () => ({ status: null, error: { message: 'ETIMEDOUT' } })
    }),
    (error) => error instanceof PowerPointComError && /timed out after 1000 ms/.test(error.message)
  );
});

test('status reports auto fallback and forced COM errors without exposing secrets', () => {
  const fallback = getPowerPointComStatus({ config: { backend: 'auto' }, platform: 'linux' });
  assert.deepEqual({
    configured_mode: fallback.configured_mode,
    selected_backend: fallback.selected_backend,
    available: fallback.available
  }, { configured_mode: 'auto', selected_backend: 'native', available: false });

  const forced = getPowerPointComStatus({ config: { backend: 'powerpoint-com' }, platform: 'linux' });
  assert.equal(forced.selected_backend, null);
  assert.match(forced.reason, /Windows/);
});

test('status can explicitly report that COM probing was skipped', () => {
  const status = getPowerPointComStatus({
    config: { backend: 'auto' },
    platform: 'win32',
    probe: false
  });
  assert.equal(status.probe_performed, false);
  assert.equal(status.available, null);
  assert.equal(status.selected_backend, null);
});

test('COM backend carries normalized config into its probe and script runner', () => {
  const calls = [];
  const backend = new PowerPointComBackend({
    backend: 'powerpoint-com', powerpoint_com: {
      visible: true, display_alerts: true, timeout_ms: 4321
    }
  }, {
    platform: 'win32',
    runner: (executable, args, options) => {
      calls.push({ executable, args, options });
      return { status: 0, stdout: 'ok', stderr: '' };
    }
  });
  assert.equal(backend.status().selected_backend, 'powerpoint-com');
  backend.runScript('verify-powerpoint-open.ps1', ['-InputPath', 'deck.pptx']);
  const scriptCall = calls.at(-1);
  assert.equal(scriptCall.options.timeout, 4321);
  assert.equal(scriptCall.options.env.PPT_CREATER_POWERPOINT_VISIBLE, '1');
  assert.equal(scriptCall.options.env.PPT_CREATER_POWERPOINT_DISPLAY_ALERTS, '1');
});

test('COM script runner passes absolute script and user arguments unchanged', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-creater-com-'));
  const calls = [];
  const result = runPowerPointComScript('verify-powerpoint-open.ps1', ['-InputPath', path.join(root, 'deck.pptx')], {
    platform: 'win32', timeoutMs: 3210,
    runner: (executable, args, options) => {
      calls.push({ executable, args, options });
      return { status: 0, stdout: 'ok', stderr: '' };
    }
  });
  assert.equal(result.stdout, 'ok');
  assert.equal(calls.length, 1);
  assert.equal(calls[0].executable, 'powershell.exe');
  assert.equal(calls[0].args.at(-1), path.join(root, 'deck.pptx'));
  assert.equal(calls[0].options.timeout, 3210);
});

test('native pipeline contract rejects non-native generation', () => {
  assert.throws(() => assertNativePipeline({ pipeline: { native_only: false } }), /native-template pipeline/);
  assert.deepEqual(assertNativePipeline({
    pipeline: { native_only: true, require_preview_all_pages: true, require_text_optimization: true, require_powerpoint_qa: true },
    source_pptx: 'source.pptx',
    master_pptx: 'master.pptx',
    output_pptx: 'output.pptx',
    composition_mode: 'native'
  }), { version: NATIVE_PIPELINE_VERSION, valid: true });
  const contract = buildNativePipelineContract({ sourcePptx: 'source.pptx', masterPptx: 'master.pptx', outputPptx: 'output.pptx' });
  assert.equal(contract.pipeline.native_only, true);
  assert.equal(contract.pipeline.require_preview_all_pages, true);
  assert.throws(() => buildNativePipelineContract({ sourcePptx: 's', masterPptx: 'm', outputPptx: 'o', compositionMode: 'harmonized' }), /composition_mode/);
});

test('text optimizer shortens semantically without substring truncation', () => {
  const result = optimizeTextForSlot(
    '建立可观测闭环：将质量、时延、成本、人工介入与业务结果纳入同一仪表盘。',
    { maxChars: 16, role: 'body' }
  );
  assert.equal(result.unresolved, false);
  assert.ok(result.text.length <= 16);
  assert.doesNotMatch(result.text, /建立可观测闭环：将质量/);
  assert.match(result.text, /闭环|质量|成本|业务/);
});

test('native pipeline optimizes every page content before composition', () => {
  const result = buildOptimizedPages([
    { page_index: 1, content: { slot_title: '建立可观测闭环：将质量、时延、成本、人工介入与业务结果纳入仪表盘。' } },
    { page_index: 2, content: { slot_body: '短文案' } }
  ], (_page, slotName) => ({ max_chars_cn: slotName === 'slot_title' ? 12 : 20, role: 'body' }));
  assert.equal(result.pages.length, 2);
  assert.equal(result.unresolved.length, 0);
  assert.equal(result.pages[0].content.slot_title.length <= 12, true);
  assert.equal(result.pages[1].content.slot_body, '短文案');
  assert.ok(result.changes.length >= 1);
});

test('object fill normalization supports registered native kinds', () => {
  assert.deepEqual(normalizeObjectFill({ name: 'slot_table', kind: 'table_cell', value: '42', locator: { row: 2, column: 3 } }), {
    name: 'slot_table', kind: 'table_cell', value: '42', locator: { row: 2, column: 3 }
  });
  assert.throws(() => normalizeObjectFill({ name: 'unknown', kind: 'arbitrary', value: 'x' }), /Unsupported native object kind/);
});

test('preview contract requires every slide to have a nonempty image', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-creater-preview-'));
  const first = path.join(dir, 'slide-001.png');
  const second = path.join(dir, 'slide-002.png');
  fs.writeFileSync(first, Buffer.from('png'));
  fs.writeFileSync(second, Buffer.from('png'));
  assert.deepEqual(buildPreviewContract({ slideCount: 2, files: [first, second] }), {
    slide_count: 2, preview_count: 2, all_pages_previewed: true
  });
  assert.throws(() => buildPreviewContract({ slideCount: 2, files: [first] }), /every slide/);
});
