const fs = require('node:fs');
const path = require('node:path');
const { Automizer, modify } = require('pptx-automizer');
const { resolveTypographyTheme, createTypographyModifier } = require('./typography-system');
const { PowerPointComBackend } = require('./powerpoint-com');
const {
  assertNativePipeline,
  buildNativePipelineContract,
  buildOptimizedPages,
  normalizeObjectFill
} = require('./native-pipeline');

const USER_COLOR_TOKENS = Object.freeze(['background', 'text', 'muted', 'accent', 'line', 'darkBackground']);
const HEX = /^[0-9A-F]{6}$/;

function normalizeHex(value, label) {
  const normalized = String(value || '').replace(/^#/, '').toUpperCase();
  if (!HEX.test(normalized)) throw new Error(`${label} must be a six-digit hex color.`);
  return normalized;
}

function normalizeColorMap(colorMap = {}) {
  if (!colorMap || typeof colorMap !== 'object' || Array.isArray(colorMap)) {
    throw new Error('color_map must be an object mapping source hex colors to target hex colors.');
  }
  return Object.fromEntries(Object.entries(colorMap).map(([from, to]) => [
    normalizeHex(from, 'color_map source'), normalizeHex(to, 'color_map target')
  ]));
}

function validateCompletePalette(colors) {
  const missing = USER_COLOR_TOKENS.filter((key) => !colors?.[key]);
  if (missing.length) throw new Error(`Color palette is incomplete: ${USER_COLOR_TOKENS.join(', ')}.`);
  return Object.fromEntries(Object.entries(colors).map(([key, value]) => [key, normalizeHex(value, `colors.${key}`)]));
}

function resolveCatalogPaths(value, catalogPath) {
  if (Array.isArray(value)) return value.map((item) => resolveCatalogPaths(item, catalogPath));
  if (!value || typeof value !== 'object') return value;
  const root = path.resolve(path.dirname(catalogPath), '..');
  return Object.fromEntries(Object.entries(value).map(([key, item]) => {
    if (typeof item === 'string' && /(?:^|_)pptx$/.test(key) && !path.isAbsolute(item)) {
      return [key, path.resolve(root, item)];
    }
    return [key, resolveCatalogPaths(item, catalogPath)];
  }));
}

function loadCompositionCatalog(catalogPath) {
  const catalog = JSON.parse(fs.readFileSync(catalogPath, 'utf8').replace(/^\uFEFF/, ''));
  if (!catalog?.targets) throw new Error('Composition catalog must provide targets.');
  return resolveCatalogPaths(catalog, catalogPath);
}

function resolveTarget(catalog, target) {
  const family = catalog.targets?.[target];
  if (!family) throw new Error(`Unknown composition target: ${target}`);
  if (family.status !== 'production_ready') throw new Error(`Target is not production-ready: ${target}`);
  if (!family.master_pptx || !fs.existsSync(family.master_pptx)) throw new Error(`Master PPTX is unavailable: ${target}`);
  if (!family.source_pptx || !fs.existsSync(family.source_pptx)) throw new Error(`Source PPTX is unavailable: ${target}`);
  return family;
}

function pageDefinition(target, pageIndex) {
  const page = target.pages?.find((item) => item.page_index === pageIndex);
  if (!page) throw new Error(`Target has no registered page_index: ${pageIndex}`);
  return page;
}

function resolveSlotDefinition(catalog, baseTarget, page, slotName) {
  const target = resolveTarget(catalog, page.target || baseTarget);
  const definition = pageDefinition(target, page.page_index);
  return (definition.slots || []).find((slot) => slot.name === slotName) || null;
}

function validatePage(catalog, baseTarget, page, mode) {
  if (mode !== 'native') {
    throw new Error('PPT-creater public generation requires composition_mode "native" and one template family.');
  }
  const targetName = page.target || baseTarget;
  if (targetName !== baseTarget) {
    throw new Error(`Cross-target page ${targetName}:${page.page_index} is not allowed in native composition.`);
  }
  const target = resolveTarget(catalog, targetName);
  const definition = pageDefinition(target, page.page_index);
  const slots = new Map((definition.slots || []).map((slot) => [slot.name, slot]));
  const objects = new Map((definition.objects || []).map((object) => [object.name, object]));
  for (const slotName of Object.keys(page.content || {})) {
    if (!slots.has(slotName)) throw new Error(`Page ${targetName}:${page.page_index} has no registered slot: ${slotName}`);
  }
  const objectFills = (page.object_fills || []).map(normalizeObjectFill);
  for (const fill of objectFills) {
    if (!slots.has(fill.name) && !objects.has(fill.name)) {
      throw new Error(`Page ${targetName}:${page.page_index} has no registered native object: ${fill.name}`);
    }
  }
  return { targetName, target, definition, slots, objects, objectFills };
}

function buildCompositionPlan(catalog, request, { optimize = true, outputPptx = 'pending-output.pptx' } = {}) {
  const baseTarget = request.base_target;
  const mode = request.composition_mode || 'native';
  if (mode !== 'native') throw new Error('composition_mode must be "native".');
  if (!Array.isArray(request.pages) || request.pages.length === 0) throw new Error('pages must contain at least one page.');
  const base = resolveTarget(catalog, baseTarget);
  const typographyTheme = request.typography_theme_id ? resolveTypographyTheme(request.typography_theme_id) : null;
  const rawPages = request.pages.map((page, index) => {
    if (!Number.isInteger(page.page_index) || page.page_index < 1) {
      throw new Error(`pages[${index}].page_index must be a positive integer.`);
    }
    validatePage(catalog, baseTarget, page, mode);
    return { ...page, content: { ...(page.content || {}) }, object_fills: [...(page.object_fills || [])] };
  });

  const optimization = optimize
    ? buildOptimizedPages(rawPages, (page, slotName) => resolveSlotDefinition(catalog, baseTarget, page, slotName))
    : { pages: rawPages, changes: [], unresolved: [] };
  if (optimization.unresolved.length > 0) {
    const unresolved = optimization.unresolved.map((item) => `${item.page_index}:${item.slot_name}`).join(', ');
    throw new Error(`Text optimization could not fit registered slots: ${unresolved}.`);
  }
  if (optimize) {
    for (const page of optimization.pages) {
      for (const [slotName, value] of Object.entries(page.content || {})) {
        const slot = resolveSlotDefinition(catalog, baseTarget, page, slotName);
        if (slot && String(value).length > slot.max_chars_cn) {
          throw new Error(`Optimized content exceeds ${slotName} capacity (${slot.max_chars_cn} characters).`);
        }
      }
    }
  }

  const changeMap = new Map(optimization.changes.map((item) => [`${item.page_index}:${item.slot_name}`, item]));
  const pages = optimization.pages.map((page, index) => {
    const validated = validatePage(catalog, baseTarget, page, mode);
    const contentOriginal = { ...(rawPages[index].content || {}) };
    const contentFinal = { ...(page.content || {}) };
    const slotOptimization = {};
    for (const slotName of Object.keys(contentOriginal)) {
      slotOptimization[slotName] = changeMap.get(`${page.page_index}:${slotName}`) || {
        text: contentFinal[slotName], changed: false, unresolved: false, strategy: 'unchanged'
      };
    }
    return {
      index: index + 1,
      target: validated.targetName,
      master_pptx: validated.target.master_pptx,
      page_index: page.page_index,
      animation_effects: validated.definition.animation_effects || 0,
      content: contentFinal,
      content_original: contentOriginal,
      content_final: contentFinal,
      optimization: slotOptimization,
      object_fills: validated.objectFills
    };
  });

  const sourcePptx = base.source_pptx;
  return {
    version: 1,
    base_target: baseTarget,
    base_master_pptx: base.master_pptx,
    source_pptx: sourcePptx,
    composition_mode: 'native',
    colors: null,
    color_map: {},
    typography_theme_id: request.typography_theme_id || null,
    typography_theme: typographyTheme,
    native_pipeline: buildNativePipelineContract({
      sourcePptx,
      masterPptx: base.master_pptx,
      outputPptx,
      compositionMode: 'native'
    }),
    text_optimization: {
      attempted: optimize,
      changed_count: optimization.changes.filter((item) => item.changed).length,
      unresolved_count: optimization.unresolved.length,
      changes: optimization.changes
    },
    optimization: {
      attempted: optimize,
      changed_count: optimization.changes.filter((item) => item.changed).length,
      unresolved_count: optimization.unresolved.length,
      changes: optimization.changes
    },
    pages
  };
}

function objectFillRequiresCom(fill) {
  return fill.kind !== 'text';
}

function readJsonFile(filePath) {
  return JSON.parse(fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, ''));
}

function writeJsonFile(filePath, value) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function qaPaths(output) {
  const outputDir = path.dirname(output);
  const stem = path.basename(output, path.extname(output));
  const previewRoot = path.join(outputDir, 'previews', stem);
  return {
    plan: path.join(outputDir, `${stem}.plan.json`),
    master_preview_directory: path.join(previewRoot, 'master'),
    master_preview_report: path.join(previewRoot, 'master.report.json'),
    preview_directory: path.join(previewRoot, 'final'),
    preview_report: path.join(previewRoot, 'final.report.json'),
    powerpoint_report: path.join(outputDir, `${stem}.powerpoint-qa.json`),
    qa_report: path.join(outputDir, `${stem}.qa.json`)
  };
}

function runFullPowerPointQa({ output, masterPptx, plan, powerpoint }) {
  if (!powerpoint) throw new Error('Full PowerPoint QA requires the PowerPoint COM backend.');
  const status = powerpoint.status();
  if (status.selected_backend !== 'powerpoint-com') {
    throw new Error(`Full PowerPoint QA requires PowerPoint COM: ${status.reason || 'backend unavailable.'}`);
  }
  const paths = qaPaths(output);
  writeJsonFile(paths.plan, {
    ...plan,
    qa_paths: paths,
    qa_pipeline: ['full_page_preview', 'powerpoint_open', 'native_normalization_and_playback']
  });
  powerpoint.runScript('preview-all-pages.ps1', [
    '-InputPath', masterPptx,
    '-PreviewDirectory', paths.master_preview_directory,
    '-ReportPath', paths.master_preview_report
  ]);
  const masterPreview = readJsonFile(paths.master_preview_report);
  if (masterPreview.status !== 'verified') {
    throw new Error(`Full PowerPoint QA master preview did not pass: ${paths.master_preview_report}`);
  }
  powerpoint.runScript('preview-all-pages.ps1', [
    '-InputPath', output,
    '-PreviewDirectory', paths.preview_directory,
    '-ReportPath', paths.preview_report
  ]);
  const preview = readJsonFile(paths.preview_report);
  if (preview.status !== 'verified') {
    throw new Error(`Full PowerPoint QA preview did not pass: ${paths.preview_report}`);
  }
  powerpoint.runScript('verify-powerpoint-open.ps1', [
    '-InputPath', output,
    '-ReportPath', paths.powerpoint_report
  ]);
  const open = readJsonFile(paths.powerpoint_report);
  if (open.opened !== true) {
    throw new Error(`PowerPoint open verification did not pass: ${paths.powerpoint_report}`);
  }
  powerpoint.runScript('qa-full-native.ps1', [
    '-SourcePath', masterPptx,
    '-OutputPath', output,
    '-PlanPath', paths.plan,
    '-ReportPath', paths.qa_report,
    '-BeforePreviewReportPath', paths.master_preview_report,
    '-AfterPreviewReportPath', paths.preview_report
  ]);
  const qa = readJsonFile(paths.qa_report);
  if (qa.status !== 'verified') {
    throw new Error(`Full native PowerPoint QA did not pass: ${paths.qa_report}`);
  }
  return {
    status: 'verified',
    reports: paths,
    preview,
    open,
    qa
  };
}

function applyComObjectFills(output, plan, fills, powerpoint) {
  if (fills.length === 0) return { applied: true, count: 0 };
  if (!powerpoint) throw new Error('Native COM object fills require the PowerPoint COM backend.');
  const tempPlan = `${output}.native-object-fills.json`;
  fs.writeFileSync(tempPlan, JSON.stringify({
    pages: plan.pages.map((page) => ({
      index: page.index,
      page_index: page.page_index,
      object_fills: (page.object_fills || []).filter(objectFillRequiresCom)
    }))
  }, null, 2), 'utf8');
  try {
    const result = powerpoint.runScript('apply-native-object-fills.ps1', [
      '-InputPath', output, '-PlanPath', tempPlan
    ]);
    return { applied: true, count: fills.length, output: result.stdout?.trim() || null };
  } finally {
    try { fs.unlinkSync(tempPlan); } catch {}
  }
}

async function composeDeck({ plan, outputPath, draft = false, powerpoint = null, powerpointConfig = {}, qa = false }) {
  const output = path.resolve(outputPath);
  const tempOutput = `${output}.tmp-${process.pid}-${Date.now()}.pptx`;
  const sourcePptx = plan?.source_pptx || plan?.native_pipeline?.source_pptx;
  const masterPptx = plan?.base_master_pptx || plan?.native_pipeline?.master_pptx;
  if (!sourcePptx || !fs.existsSync(sourcePptx)) throw new Error('Plan source_pptx is unavailable.');
  if (!masterPptx || !fs.existsSync(masterPptx)) throw new Error('Plan master_pptx is unavailable.');
  const nativePipeline = buildNativePipelineContract({
    sourcePptx, masterPptx, outputPptx: output, compositionMode: plan.composition_mode || 'native'
  });
  assertNativePipeline(nativePipeline);
  if (!draft && plan.text_optimization?.attempted !== true) {
    throw new Error('Final native composition requires text optimization before writing.');
  }

  const outputDir = path.dirname(output);
  fs.mkdirSync(outputDir, { recursive: true });
  const masterPath = path.resolve(masterPptx);
  const pageMasters = new Set((plan.pages || []).map((page) => path.resolve(page.master_pptx)));
  if ([...pageMasters].some((pageMaster) => pageMaster !== masterPath)) {
    throw new Error('A native deck may use only one registered template family.');
  }

  const comFills = (plan.pages || []).flatMap((page) =>
    (page.object_fills || []).filter(objectFillRequiresCom)
  );
  const comBackend = powerpoint || ((comFills.length > 0 || qa) ? new PowerPointComBackend(powerpointConfig) : null);
  if (comFills.length > 0 || qa) {
    if (!comBackend) throw new Error('Native COM object fills and full QA require the PowerPoint COM backend.');
    const status = comBackend.status();
    if (status.selected_backend !== 'powerpoint-com') {
      throw new Error(`PowerPoint COM is required for native object fills or full QA: ${status.reason || 'backend unavailable.'}`);
    }
  }

  const backupOutput = `${output}.backup-${process.pid}-${Date.now()}.pptx`;
  let originalMoved = false;
  let published = false;
  try {
    const automizer = new Automizer({ outputDir, removeExistingSlides: true, cleanup: true });
    const root = fs.readFileSync(masterPath);
    const presentation = automizer.loadRoot(root).load(root, 'family');
    for (const page of plan.pages || []) {
      presentation.addSlide('family', page.page_index, (slide) => {
        for (const [slotName, value] of Object.entries(page.content_final || page.content || {})) {
          const typography = plan.typography_theme ? createTypographyModifier(plan.typography_theme, slotName) : null;
          slide.modifyElement(slotName, [modify.setText(String(value)), ...(typography ? [typography] : [])]);
        }
        for (const fill of page.object_fills || []) {
          if (objectFillRequiresCom(fill)) continue;
          slide.modifyElement(fill.name, [modify.setText(String(fill.value))]);
        }
      });
    }
    await presentation.write(path.basename(tempOutput));
    const objectFillResult = applyComObjectFills(tempOutput, plan, comFills, comBackend);

    if (fs.existsSync(output)) {
      fs.renameSync(output, backupOutput);
      originalMoved = true;
    }
    fs.renameSync(tempOutput, output);
    published = true;

    const qaResult = qa
      ? runFullPowerPointQa({ output, masterPptx: masterPath, plan, powerpoint: comBackend })
      : null;

    if (originalMoved) {
      fs.unlinkSync(backupOutput);
      originalMoved = false;
    }
    return { output, object_fill_result: objectFillResult, ...(qaResult ? { qa: qaResult } : {}) };
  } catch (error) {
    try { fs.unlinkSync(tempOutput); } catch {}
    if (published) {
      try { fs.unlinkSync(output); } catch {}
    }
    if (originalMoved) {
      try { fs.renameSync(backupOutput, output); originalMoved = false; } catch {}
    }
    throw error;
  } finally {
    if (originalMoved) {
      try { fs.renameSync(backupOutput, output); } catch {}
    }
    try { fs.unlinkSync(backupOutput); } catch {}
  }
}

module.exports = {
  USER_COLOR_TOKENS,
  normalizeColorMap,
  validateCompletePalette,
  loadCompositionCatalog,
  buildCompositionPlan,
  composeDeck
};
