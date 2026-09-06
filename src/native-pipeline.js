const fs = require('node:fs');

const NATIVE_PIPELINE_VERSION = 1;

const NATIVE_OBJECT_KINDS = Object.freeze([
  'text',
  'group_text',
  'table_cell',
  'smartart_text',
  'chart_title',
  'chart_label',
  'chart_data',
  'image_placeholder',
  'image_fill',
  'shape_fill'
]);

function buildNativePipelineContract({ sourcePptx, masterPptx, outputPptx, compositionMode = 'native' } = {}) {
  if (!sourcePptx || !masterPptx || !outputPptx) {
    throw new Error('Native-template pipeline requires sourcePptx, masterPptx, and outputPptx.');
  }
  if (compositionMode !== 'native') {
    throw new Error('Public PPT-creater generation must use composition_mode "native".');
  }
  return {
    version: NATIVE_PIPELINE_VERSION,
    source_pptx: sourcePptx,
    master_pptx: masterPptx,
    output_pptx: outputPptx,
    composition_mode: compositionMode,
    pipeline: {
      name: 'native-template',
      version: NATIVE_PIPELINE_VERSION,
      native_only: true,
      require_native_routing: true,
      require_native_page_copy: true,
      require_native_object_fill: true,
      require_preview_all_pages: true,
      require_text_optimization: true,
      require_powerpoint_normalization: true,
      require_powerpoint_qa: true,
      require_animation_playback: true
    },
    strict_native_only: true
  };
}

function assertNativePipeline(contract = {}) {
  const pipeline = contract.pipeline || {};
  if (
    pipeline.native_only !== true ||
    pipeline.require_preview_all_pages !== true ||
    pipeline.require_text_optimization !== true ||
    pipeline.require_powerpoint_qa !== true
  ) {
    throw new Error('All public PPT-creater generation must use the native-template pipeline.');
  }
  if (
    contract.composition_mode !== 'native' ||
    !contract.source_pptx ||
    !contract.master_pptx ||
    !contract.output_pptx
  ) {
    throw new Error('Native-template pipeline requires native composition and source/master/output PPTX paths.');
  }
  return { version: NATIVE_PIPELINE_VERSION, valid: true };
}

function normalizeObjectFill(fill = {}) {
  if (!fill || typeof fill !== 'object' || Array.isArray(fill)) {
    throw new Error('Native object fill must be an object.');
  }
  const name = String(fill.name || '').trim();
  const kind = String(fill.kind || '').trim();
  if (!name) throw new Error('Native object fill requires name.');
  if (!NATIVE_OBJECT_KINDS.includes(kind)) {
    throw new Error(`Unsupported native object kind: ${kind}`);
  }
  if (!Object.prototype.hasOwnProperty.call(fill, 'value')) {
    throw new Error(`Native object fill ${name} requires value.`);
  }
  const normalized = { name, kind, value: fill.value };
  if (fill.locator !== undefined) {
    if (!fill.locator || typeof fill.locator !== 'object' || Array.isArray(fill.locator)) {
      throw new Error(`Native object fill ${name} locator must be an object.`);
    }
    normalized.locator = { ...fill.locator };
  }
  return normalized;
}

function uniqueCandidates(values) {
  return [...new Set(values.map((value) => String(value).trim()).filter(Boolean))];
}

function semanticTextCandidates(value, role = 'body') {
  const source = String(value).replace(/\s+/g, ' ').trim();
  const candidates = [source];
  const withoutLead = source
    .replace(/^(建议|结论|核心判断|观察重点|投资判断|管理动作)[:：]\s*/u, '')
    .replace(/^(我们认为|需要|应当|可以|能够)\s*/u, '');
  candidates.push(withoutLead);

  const clauses = withoutLead.split(/[：:，。；、]/u).map((part) => part.trim()).filter(Boolean);
  if (clauses.length > 1) {
    candidates.push(clauses.slice(0, 2).join('、'));
    candidates.push(clauses.slice(-2).join('、'));
    candidates.push(clauses[0]);
  }
  const keywordPairs = [
    ['闭环', /(闭环)/u],
    ['业务结果', /(业务结果|结果)/u],
    ['质量成本', /(质量|成本)/u],
    ['工作流', /(工作流|流程)/u],
    ['数据反馈', /(数据|反馈)/u]
  ];
  const matchedKeywords = keywordPairs.filter(([, pattern]) => pattern.test(withoutLead)).map(([label]) => label);
  if (matchedKeywords.length > 0) candidates.push(matchedKeywords.slice(0, 3).join('与'));

  candidates.push(withoutLead
    .replace(/(正在|持续|共同|直接|真实|关键的|重要的|主要的|可持续的|可量化的|长期的)/gu, '')
    .replace(/(成为|形成|实现|推动|支撑|决定|纳入|围绕)/gu, '')
    .replace(/\s+/g, ' ')
    .trim());

  if (role === 'title' || role === 'heading') {
    candidates.push(withoutLead.replace(/[：，。；、]/gu, ' ').replace(/\s+/g, ' ').trim());
  }
  return uniqueCandidates(candidates);
}

function optimizeTextForSlot(value, { maxChars = Infinity, role = 'body' } = {}) {
  const source = String(value ?? '').trim();
  if (!Number.isFinite(maxChars) || maxChars < 1 || source.length <= maxChars) {
    return { text: source, changed: false, unresolved: false, strategy: 'unchanged' };
  }
  const candidates = semanticTextCandidates(source, role)
    .filter((candidate) => candidate.length <= maxChars);
  if (candidates.length > 0) {
    const selected = candidates.sort((a, b) => {
      const score = (text) => (text.includes('闭环') ? 3 : 0) + (text.includes('结果') ? 2 : 0)
        + (text.includes('质量') || text.includes('成本') || text.includes('业务') ? 1 : 0);
      return score(b) - score(a) || b.length - a.length;
    })[0];
    return { text: selected, changed: selected !== source, unresolved: false, strategy: 'semantic_rewrite' };
  }
  return { text: source, changed: false, unresolved: true, strategy: 'unresolved' };
}

function buildOptimizedPages(pages = [], slotResolver) {
  const changes = [];
  const optimizedPages = pages.map((page) => {
    const content = { ...(page.content || {}) };
    for (const [slotName, value] of Object.entries(content)) {
      const slot = slotResolver(page, slotName) || {};
      const result = optimizeTextForSlot(value, {
        maxChars: slot.max_chars_cn,
        role: slot.role || 'body'
      });
      content[slotName] = result.text;
      if (result.changed || result.unresolved) {
        changes.push({
          page_index: page.page_index,
          slot_name: slotName,
          before: String(value),
          after: result.text,
          ...result
        });
      }
    }
    return { ...page, content };
  });
  return {
    pages: optimizedPages,
    changes,
    unresolved: changes.filter((change) => change.unresolved)
  };
}

function buildPreviewContract({ slideCount, files = [], previewDirectory = null } = {}) {
  if (!Number.isInteger(slideCount) || slideCount < 1) {
    throw new Error('Preview contract requires a positive slide count.');
  }
  const normalizedFiles = files.map((file) => String(file)).filter(Boolean);
  if (normalizedFiles.length !== slideCount) {
    throw new Error(`Preview must include every slide: expected ${slideCount}, received ${normalizedFiles.length}.`);
  }
  const missing = normalizedFiles.filter((file) => !fs.existsSync(file));
  if (missing.length) throw new Error(`Preview files are missing: ${missing.join(', ')}.`);
  const empty = normalizedFiles.filter((file) => fs.statSync(file).size === 0);
  if (empty.length) throw new Error(`Preview files are empty: ${empty.join(', ')}.`);
  return {
    slide_count: slideCount,
    preview_count: normalizedFiles.length,
    all_pages_previewed: true,
    ...(previewDirectory ? { preview_directory: previewDirectory } : {})
  };
}

module.exports = {
  NATIVE_PIPELINE_VERSION,
  NATIVE_OBJECT_KINDS,
  buildNativePipelineContract,
  assertNativePipeline,
  normalizeObjectFill,
  optimizeTextForSlot,
  buildOptimizedPages,
  buildPreviewContract
};
