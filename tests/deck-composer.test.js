const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const test = require('node:test');
const { buildCompositionPlan, composeDeck } = require('../src/deck-composer');

function makeCatalog(root) {
  const create = path.join(root, 'make_templates.py');
  const alpha = path.join(root, 'alpha.pptx');
  fs.writeFileSync(create, [
    'from pptx import Presentation',
    'from pptx.util import Inches',
    'from pptx.dml.color import RGBColor',
    'def make(path):',
    ' p=Presentation(); s=p.slides.add_slide(p.slide_layouts[6])',
    " r=s.shapes.add_shape(1, Inches(1), Inches(1), Inches(2), Inches(1)); r.fill.solid(); r.fill.fore_color.rgb=RGBColor(17,34,51)",
    " t=s.shapes.add_textbox(Inches(1), Inches(2.3), Inches(6), Inches(1)); t.name='slot_alpha_title'; t.text='placeholder'",
    ' p.save(path)',
    `make(r'''${alpha.replace(/\\/g, '\\\\')}''')`
  ].join('\n'));
  execFileSync('python', [create]);
  return {
    targets: {
      alpha: {
        status: 'production_ready',
        source_pptx: alpha,
        master_pptx: alpha,
        pages: [{
          page_index: 1,
          animation_effects: 0,
          slots: [{ name: 'slot_alpha_title', max_chars_cn: 12, animation_locked: false }],
          objects: [{ name: 'object_chart', kind: 'chart_data' }]
        }]
      }
    }
  };
}

test('composition rejects cross-target pages and non-native modes', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-creater-compose-'));
  const catalog = makeCatalog(root);
  assert.throws(() => buildCompositionPlan(catalog, {
    base_target: 'alpha', composition_mode: 'harmonized', pages: [{ target: 'alpha', page_index: 1, content: {} }]
  }), /composition_mode must be "native"/);
});

test('composition optimizes overlong native content before validating slot capacity', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-creater-compose-'));
  const catalog = makeCatalog(root);
  const plan = buildCompositionPlan(catalog, {
    base_target: 'alpha',
    composition_mode: 'native',
    pages: [{ page_index: 1, content: { slot_alpha_title: '建立可观测闭环：将质量、成本和业务结果纳入统一流程。' } }]
  });
  assert.equal(plan.optimization.attempted, true);
  assert.equal(plan.optimization.unresolved_count, 0);
  assert.equal(plan.pages[0].content_original.slot_alpha_title.length > 12, true);
  assert.equal(plan.pages[0].content_final.slot_alpha_title.length <= 12, true);
  assert.equal(plan.pages[0].optimization.slot_alpha_title.strategy, 'semantic_rewrite');
  const outputPath = path.join(root, 'planned-output.pptx');
  const outputPlan = buildCompositionPlan(catalog, {
    base_target: 'alpha',
    composition_mode: 'native',
    pages: [{ page_index: 1, content: { slot_alpha_title: '标题' } }]
  }, { outputPptx: outputPath });
  assert.equal(outputPlan.native_pipeline.output_pptx, outputPath);
});

test('composition rejects non-text fills before writing when COM is unavailable', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-creater-compose-com-'));
  const catalog = makeCatalog(root);
  const plan = buildCompositionPlan(catalog, {
    base_target: 'alpha',
    composition_mode: 'native',
    pages: [{
      page_index: 1,
      content: { slot_alpha_title: '原生首页' },
      object_fills: [{ name: 'object_chart', kind: 'chart_data', value: 42, locator: { series: 1, point: 1 } }]
    }]
  });
  const output = path.join(root, 'blocked.pptx');
  await assert.rejects(
    composeDeck({
      plan,
      outputPath: output,
      powerpoint: { status: () => ({ selected_backend: 'native', reason: 'COM disabled' }) }
    }),
    /PowerPoint COM is required/
  );
  assert.equal(fs.existsSync(output), false);
});

test('composition removes temporary output when COM object fill fails', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-creater-compose-com-'));
  const catalog = makeCatalog(root);
  const plan = buildCompositionPlan(catalog, {
    base_target: 'alpha',
    composition_mode: 'native',
    pages: [{
      page_index: 1,
      content: { slot_alpha_title: '原生首页' },
      object_fills: [{ name: 'object_chart', kind: 'chart_data', value: 42, locator: { series: 1, point: 1 } }]
    }]
  });
  const output = path.join(root, 'failed.pptx');
  await assert.rejects(
    composeDeck({
      plan,
      outputPath: output,
      powerpoint: {
        status: () => ({ selected_backend: 'powerpoint-com' }),
        runScript: () => { throw new Error('COM fill failed'); }
      }
    }),
    /COM fill failed/
  );
  assert.equal(fs.existsSync(output), false);
  assert.equal(fs.readdirSync(root).some((name) => name.includes('.tmp-')), false);
});

test('composition runs full PowerPoint QA transactionally when requested', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-creater-compose-qa-'));
  const catalog = makeCatalog(root);
  const plan = buildCompositionPlan(catalog, {
    base_target: 'alpha',
    composition_mode: 'native',
    pages: [{ page_index: 1, content: { slot_alpha_title: '已验证首页' } }]
  });
  const output = path.join(root, 'verified.pptx');
  const calls = [];
  const powerpoint = {
    status: () => ({ selected_backend: 'powerpoint-com', available: true }),
    runScript: (script, args) => {
      calls.push(script);
      const argument = (name) => args[args.indexOf(name) + 1];
      if (script === 'preview-all-pages.ps1') {
        const directory = argument('-PreviewDirectory');
        fs.mkdirSync(directory, { recursive: true });
        fs.writeFileSync(path.join(directory, 'slide-001.png'), 'png');
        fs.writeFileSync(argument('-ReportPath'), JSON.stringify({
          status: 'verified', slide_count: 1, exported_count: 1, failed_count: 0,
          all_pages_previewed: true, text_overflow: 0, overlap_count: 0
        }));
      } else if (script === 'verify-powerpoint-open.ps1') {
        fs.writeFileSync(argument('-ReportPath'), JSON.stringify({ opened: true, slide_count: 1 }));
      } else if (script === 'qa-full-native.ps1') {
        fs.writeFileSync(argument('-ReportPath'), JSON.stringify({
          status: 'verified', new_shapes: 0, powerpoint_playback: true,
          timeline_signature_match: true
        }));
      }
      return { status: 0, stdout: 'ok', stderr: '' };
    }
  };
  const result = await composeDeck({ plan, outputPath: output, powerpoint, qa: true });
  assert.equal(result.qa.status, 'verified');
  assert.equal(fs.existsSync(output), true);
  assert.deepEqual(calls, [
    'preview-all-pages.ps1', 'preview-all-pages.ps1',
    'verify-powerpoint-open.ps1', 'qa-full-native.ps1'
  ]);
  assert.equal(fs.existsSync(path.join(root, 'verified.plan.json')), true);
  assert.equal(fs.existsSync(path.join(root, 'previews', 'verified', 'master.report.json')), true);
  assert.equal(fs.existsSync(path.join(root, 'previews', 'verified', 'final.report.json')), true);
});

test('composition restores the previous output when full PowerPoint QA fails', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-creater-compose-qa-'));
  const catalog = makeCatalog(root);
  const plan = buildCompositionPlan(catalog, {
    base_target: 'alpha',
    composition_mode: 'native',
    pages: [{ page_index: 1, content: { slot_alpha_title: '新内容' } }]
  });
  const output = path.join(root, 'existing.pptx');
  fs.writeFileSync(output, 'previous output');
  const powerpoint = {
    status: () => ({ selected_backend: 'powerpoint-com', available: true }),
    runScript: (script, args) => {
      if (script === 'preview-all-pages.ps1') {
        const directory = args[args.indexOf('-PreviewDirectory') + 1];
        const reportPath = args[args.indexOf('-ReportPath') + 1];
        const isMasterPreview = reportPath.endsWith('master.report.json');
        fs.mkdirSync(directory, { recursive: true });
        fs.writeFileSync(path.join(directory, 'slide-001.png'), 'png');
        fs.writeFileSync(reportPath, JSON.stringify({
          status: isMasterPreview ? 'verified' : 'review_required',
          slide_count: 1, exported_count: 1, failed_count: 0,
          all_pages_previewed: true, text_overflow: isMasterPreview ? 0 : 1, overlap_count: 0
        }));
      }
      return { status: 0, stdout: 'ok', stderr: '' };
    }
  };
  await assert.rejects(
    composeDeck({ plan, outputPath: output, powerpoint, qa: true }),
    /Full PowerPoint QA preview did not pass/
  );
  assert.equal(fs.readFileSync(output, 'utf8'), 'previous output');
  assert.equal(fs.readdirSync(root).some((name) => name.includes('.tmp-')), false);
});

test('composition copies native pages, fills registered slots, and keeps one family', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-creater-compose-'));
  const catalog = makeCatalog(root);
  const plan = buildCompositionPlan(catalog, {
    base_target: 'alpha',
    composition_mode: 'native',
    pages: [{ page_index: 1, content: { slot_alpha_title: '原生首页' } }]
  });
  const output = path.join(root, 'composed.pptx');
  await composeDeck({ plan, outputPath: output });
  assert.equal(fs.existsSync(output), true);
  const inspect = `from pptx import Presentation\np=Presentation(r'''${output.replace(/\\/g, '\\\\')}''')\nprint(len(p.slides))\nprint('\\n'.join(sh.text for s in p.slides for sh in s.shapes if getattr(sh,'has_text_frame',False)))`;
  const found = execFileSync('python', ['-c', inspect], { encoding: 'utf8' });
  assert.match(found, /^1/m);
  assert.match(found, /原生首页/);
});
