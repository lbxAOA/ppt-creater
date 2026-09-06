const fs = require('node:fs');
const path = require('node:path');
const {
  loadCompositionCatalog,
  buildCompositionPlan,
  composeDeck
} = require('./deck-composer');
const { PowerPointComBackend } = require('./powerpoint-com');

const ROOT = path.resolve(__dirname, '..');
const catalogPath = path.join(ROOT, 'catalog', 'complete-production-template-library.json');
const outputPath = path.join(ROOT, 'output', '2026-ai-industry-report-full-native.pptx');
const draftPath = path.join(ROOT, 'output', '2026-ai-industry-report-full-native-draft.pptx');
const previewRoot = path.join(ROOT, 'output', 'previews', '2026-ai-industry-report-full-native');
const masterPreviewDirectory = path.join(previewRoot, 'master');
const draftPreviewDirectory = path.join(previewRoot, 'draft');
const finalPreviewDirectory = path.join(previewRoot, 'final');
const masterPreviewReportPath = path.join(previewRoot, 'master.report.json');
const draftPreviewReportPath = path.join(previewRoot, 'draft.report.json');
const finalPreviewReportPath = path.join(previewRoot, 'final.report.json');
const qaPath = path.join(ROOT, 'output', '2026-ai-industry-report-full-native.qa.json');
const openQaPath = path.join(ROOT, 'output', '2026-ai-industry-report-full-native.powerpoint-qa.json');
const planPath = path.join(ROOT, 'output', '2026-ai-industry-report-full-native.plan.json');
const draftPlanPath = path.join(ROOT, 'output', '2026-ai-industry-report-full-native-draft.plan.json');
const configPath = process.env.PPT_CREATER_CONFIG || path.join(ROOT, 'config', 'ppt-creater.config.json');

const pages = [
  { page_index: 1, content: { slot_s01_04: 'AI行业', slot_s01_06: '研究简报' } },
  { page_index: 2, content: { slot_s02_07: '目录', slot_s02_10: '价值栈', slot_s02_13: '闭环', slot_s02_16: '商业化', slot_s02_19: '结论' } },
  { page_index: 3, content: { slot_s03_05: 'AI 01', slot_s03_06: '系统价值' } },
  { page_index: 4, content: {
    slot_s04_01: '应用与工作流', slot_s04_02: '行业智能体、Copilot与自动化系统直接承担业务结果。',
    slot_s04_03: '模型与平台', slot_s04_04: '基础模型、多模态、推理、评测与部署共同决定平台能力。',
    slot_s04_05: '数据与工具链', slot_s04_06: '数据治理、检索、观测与安全构成规模化落地底座。',
    slot_s04_07: '基础设施', slot_s04_08: '芯片、云、网络、存储和能耗优化决定规模边界。'
  } },
  { page_index: 5, content: {
    slot_s05_44: '感知与上下文', slot_s05_45: '多模态输入、企业知识与业务状态。',
    slot_s05_46: '推理与建模', slot_s05_47: '理解任务并预测候选行动的后果。',
    slot_s05_48: '规划与调用', slot_s05_49: '分解目标，调用工具并协调人机协作。',
    slot_s05_50: '执行与反馈', slot_s05_51: '写入真实流程，监控质量、成本并纠错。',
    slot_s05_52: '持续学习', slot_s05_53: '将部署反馈沉淀为可复用的数据闭环。',
    slot_s05_71: '感知', slot_s05_72: '推理', slot_s05_73: '规划', slot_s05_74: '执行', slot_s05_75: '学习'
  } },
  { page_index: 6, content: {
    slot_s06_11: '模型供给', slot_s06_12: '模型能力商品化加速，价格与推理效率成为关键。', slot_s06_13: '01',
    slot_s06_14: '平台集成', slot_s06_15: '数据、权限、工作流与可观测性决定部署成本。', slot_s06_16: '02',
    slot_s06_17: '行业应用', slot_s06_18: '围绕高频刚需和可量化结果的场景更容易付费。', slot_s06_19: '03',
    slot_s06_20: '服务与治理', slot_s06_21: '安全、合规、评测与变更管理支撑长期运行。', slot_s06_22: '04'
  } },
  { page_index: 7, content: {
    slot_s07_05: 'ROI', slot_s07_06: '价值', slot_s07_07: '深度', slot_s07_08: '规模',
    slot_s07_13: '可验证结果', slot_s07_14: '首个成果', slot_s07_15: '工作流嵌入', slot_s07_16: '日常使用',
    slot_s07_17: '数据飞轮', slot_s07_18: '专有反馈形成长期差异化。', slot_s07_19: '单位经济', slot_s07_20: '推理与交付成本必须随规模下降。'
  } },
  { page_index: 8, content: {
    slot_s08_02: '投资判断：以业务结果、使用深度、交付成本与数据闭环替代泛化的参数叙事。',
    slot_s08_03: '可验证ROI', slot_s08_04: '90天', slot_s08_08: '工作流嵌入', slot_s08_11: '数据飞轮', slot_s08_14: '单位经济'
  } },
  { page_index: 9, content: {
    slot_s09_05: '模型幻觉与可靠性', slot_s09_06: '高风险场景需要评测、人工兜底与回滚。',
    slot_s09_07: '推理成本失控', slot_s09_08: '持续观测单位成本和规模效率。',
    slot_s09_09: '高', slot_s09_10: '中', slot_s09_11: '中', slot_s09_12: '高'
  } },
  { page_index: 10, content: {
    slot_s10_05: '90天', slot_s10_06: '验证窗口', slot_s10_07: '4项', slot_s10_08: '结果、深度、成本、数据',
    slot_s10_09: '闭环', slot_s10_10: '将质量、时延、成本与人工介入纳入同一仪表盘。',
    slot_s10_11: '复制', slot_s10_12: '从受限试点扩展到跨团队工作流。', slot_s10_13: '建议'
  } },
  { page_index: 11, content: {
    slot_s11_13: '可靠性', slot_s11_14: '成本', slot_s11_15: '组织', slot_s11_16: '治理',
    slot_s11_17: '模型风险', slot_s11_18: '建立高风险场景准入与人工兜底。', slot_s11_19: '成本风险', slot_s11_20: '成本观测',
    slot_s11_21: '组织风险', slot_s11_22: '以业务结果推动流程改造与采用。', slot_s11_23: '治理风险', slot_s11_24: '先定义权限、责任、评测和回滚机制。'
  } },
  { page_index: 12, content: { slot_s12_01: '结论：AI的长期机会在于将模型能力嵌入真实数据、工作流与结果责任。', slot_s12_06: 'THANKS' } }
];

function writeJson(filePath, value) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, ''));
}

function runPowerShell(powerpoint, script, args) {
  try {
    return powerpoint.runScript(script, args);
  } catch (error) {
    const detail = error.stderr || error.stdout || error.message;
    throw new Error(`${script} failed: ${detail}`);
  }
}

async function main() {
  const config = readJson(configPath);
  const powerpoint = new PowerPointComBackend(config);
  const powerpointStatus = powerpoint.status();
  if (powerpointStatus.selected_backend !== 'powerpoint-com') {
    throw new Error(`Strict native report generation requires PowerPoint COM: ${powerpointStatus.reason || 'configure backend as auto or powerpoint-com.'}`);
  }
  const catalog = loadCompositionCatalog(catalogPath);
  const target = catalog.targets['ai-industry'];
  if (!target || target.status !== 'production_ready') throw new Error('ai-industry native production target is unavailable.');
  fs.mkdirSync(previewRoot, { recursive: true });

  // Gate 1: inspect every native master page before any filling.
  runPowerShell(powerpoint, 'preview-all-pages.ps1', [
    '-InputPath', target.master_pptx,
    '-PreviewDirectory', masterPreviewDirectory,
    '-ReportPath', masterPreviewReportPath
  ]);
  const masterPreview = readJson(masterPreviewReportPath);
  if (masterPreview.status !== 'verified') throw new Error('Native master preview did not pass.');

  // Gate 2: create an untouched native draft and preview it separately.
  const draftPlan = buildCompositionPlan(catalog, {
    base_target: 'ai-industry', composition_mode: 'native', pages
  }, { optimize: false, outputPptx: draftPath });
  writeJson(draftPlanPath, {
    ...draftPlan,
    catalog_path: catalogPath,
    pipeline_stages: ['native_master_preview', 'native_draft_copy', 'draft_preview', 'text_optimization', 'native_final_copy', 'final_preview', 'powerpoint_qa']
  });
  await composeDeck({ plan: draftPlan, outputPath: draftPath, draft: true, powerpoint });
  runPowerShell(powerpoint, 'preview-all-pages.ps1', [
    '-InputPath', draftPath,
    '-PreviewDirectory', draftPreviewDirectory,
    '-ReportPath', draftPreviewReportPath,
    '-AllowReview'
  ]);
  const draftPreview = readJson(draftPreviewReportPath);
  if (!draftPreview.all_pages_previewed || draftPreview.failed_count !== 0) {
    throw new Error('Native draft preview did not export every page.');
  }

  // Gate 3: optimize content against registered slot capacities, then fill only native objects.
  const plan = buildCompositionPlan(catalog, {
    base_target: 'ai-industry', composition_mode: 'native', pages
  }, { optimize: true, outputPptx: outputPath });
  writeJson(planPath, {
    ...plan,
    catalog_path: catalogPath,
    pipeline_stages: ['native_master_preview', 'native_draft_copy', 'draft_preview', 'text_optimization', 'native_final_copy', 'final_preview', 'powerpoint_qa'],
    preview_reports: {
      master: masterPreviewReportPath,
      draft: draftPreviewReportPath,
      final: finalPreviewReportPath
    }
  });
  await composeDeck({ plan, outputPath, powerpoint });

  // Gate 4: preview every final page after filling and before Office QA.
  runPowerShell(powerpoint, 'preview-all-pages.ps1', [
    '-InputPath', outputPath,
    '-PreviewDirectory', finalPreviewDirectory,
    '-ReportPath', finalPreviewReportPath
  ]);
  const finalPreview = readJson(finalPreviewReportPath);
  if (finalPreview.status !== 'verified') throw new Error('Final native preview did not pass.');

  // Gate 5: structural validation and PowerPoint open/playback verification.
  runPowerShell(powerpoint, 'verify-powerpoint-open.ps1', [
    '-InputPath', outputPath,
    '-ReportPath', openQaPath
  ]);
  runPowerShell(powerpoint, 'qa-full-native.ps1', [
    '-SourcePath', target.master_pptx,
    '-OutputPath', outputPath,
    '-PlanPath', planPath,
    '-ReportPath', qaPath,
    '-BeforePreviewReportPath', masterPreviewReportPath,
    '-AfterPreviewReportPath', finalPreviewReportPath
  ]);

  console.log(JSON.stringify({
    output_pptx: outputPath,
    plan_json: planPath,
    qa_json: qaPath,
    open_qa_json: openQaPath,
    master_preview_report: masterPreviewReportPath,
    draft_preview_report: draftPreviewReportPath,
    final_preview_report: finalPreviewReportPath,
    base_target: plan.base_target,
    slide_count: plan.pages.length,
    composition_mode: plan.composition_mode,
    text_changes: plan.text_optimization.changed_count,
    source_pptx: plan.base_master_pptx
  }));
}

main().catch((error) => {
  console.error(error.stack || error.message || error);
  process.exitCode = 1;
});
