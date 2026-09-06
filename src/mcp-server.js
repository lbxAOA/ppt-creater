const fs = require('node:fs');
const path = require('node:path');
const { McpServer } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const { z } = require('zod');
const { loadCatalog, searchTemplates } = require('./catalog');
const { chooseFamily, buildDeckPlan } = require('./planner');
const { loadCompositionCatalog, buildCompositionPlan, composeDeck } = require('./deck-composer');
const { listTypographyThemes } = require('./typography-system');
const { getPowerPointComStatus, PowerPointComBackend } = require('./powerpoint-com');

const ROOT = path.resolve(__dirname, '..');
const configPath = process.env.PPT_CREATER_CONFIG || path.join(ROOT, 'config', 'ppt-creater.config.json');
const familyPath = process.env.PPT_CREATER_FAMILIES || path.join(ROOT, 'catalog', 'families.json');
const compositionCatalogPath = process.env.PPT_CREATER_COMPOSITION_CATALOG
  || path.join(ROOT, 'catalog', 'complete-production-template-library.json');
const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));

function text(value) {
  return { content: [{ type: 'text', text: JSON.stringify(value, null, 2) }] };
}
function families() {
  if (!fs.existsSync(familyPath)) return [];
  return JSON.parse(fs.readFileSync(familyPath, 'utf8')).families || [];
}
function catalogPath() {
  return path.resolve(ROOT, config.catalog_root, 'templates.jsonl');
}
function compositionCatalog() {
  return loadCompositionCatalog(compositionCatalogPath);
}
function compositionPageCatalog(catalog) {
  return Object.fromEntries(Object.entries(catalog.targets).map(([target, family]) => [target, {
    master_pptx: family.master_pptx,
    source_pptx: family.source_pptx,
    status: family.status,
    pages: (family.pages || []).map((page) => ({
      page_index: page.page_index,
      animation_effects: page.animation_effects || 0,
      slots: (page.slots || []).map((slot) => ({
        name: slot.name,
        max_chars_cn: slot.max_chars_cn,
        animation_locked: slot.animation_locked
      }))
    }))
  }]));
}
function outputPath(outputName) {
  const safeName = path.basename(outputName);
  if (!safeName.toLowerCase().endsWith('.pptx')) throw new Error('output_name must end with .pptx.');
  return path.join(ROOT, config.output_root, safeName);
}

const server = new McpServer({ name: 'ppt-creater', version: '0.2.0' });

server.registerTool('ppt_creater_list_typography_themes', {
  description: 'List built-in typography themes with coordinated title, body, caption, and number font styles for PPT generation.',
  inputSchema: {}
}, async () => text({ themes: listTypographyThemes() }));

server.registerTool('ppt_creater_status', {
  description: 'Report local PPT template-library configuration and readiness.'
}, async () => text({
  project_root: ROOT,
  config_path: configPath,
  library_root: config.library_root || null,
  catalog_path: catalogPath(),
  catalog_ready: fs.existsSync(catalogPath()),
  family_registry_path: familyPath,
  family_count: families().length,
  composition_catalog_path: compositionCatalogPath,
  composition_catalog_ready: fs.existsSync(compositionCatalogPath),
  powerpoint_com: getPowerPointComStatus({ config, probe: false }),
  note: config.library_root ? 'Ready to scan, plan, compose, and export approved templates.' : 'Set library_root after the template library is provided.'
}));

server.registerTool('ppt_creater_powerpoint_com_status', {
  description: 'Probe the built-in PowerPoint COM backend without opening a user presentation.'
}, async () => text(getPowerPointComStatus({ config, probe: true })));

server.registerTool('ppt_creater_search_templates', {
  description: 'Search the non-destructive local PPTX template catalog.',
  inputSchema: {
    text: z.string().optional(),
    require_animation: z.boolean().optional(),
    require_slots: z.boolean().optional(),
    min_slides: z.number().int().positive().optional(),
    limit: z.number().int().positive().max(50).optional()
  }
}, async (query) => text(searchTemplates(loadCatalog(catalogPath()), query)));

server.registerTool('ppt_creater_choose_family', {
  description: 'Choose one coherent template family for a deck; never mixes unrelated visual families.',
  inputSchema: {
    use_case: z.string(),
    visual_tags: z.array(z.string()).default([]),
    aspect_ratio: z.string().default('16:9')
  }
}, async (request) => text(chooseFamily(families(), request)));

server.registerTool('ppt_creater_plan_deck', {
  description: 'Validate and create a one-family deck plan that preserves registered animation patterns.',
  inputSchema: {
    family_id: z.string().optional(),
    title: z.string(),
    use_case: z.string().optional(),
    visual_tags: z.array(z.string()).optional(),
    aspect_ratio: z.string().optional(),
    color_mode: z.enum(['template', 'user']).optional(),
    color_theme_name: z.string().optional(),
    typography_theme_id: z.string().optional(),
    colors: z.object({
      background: z.string().optional(), text: z.string().optional(), muted: z.string().optional(),
      accent: z.string().optional(), line: z.string().optional(), darkBackground: z.string().optional(),
      surface: z.string().optional(), inverseText: z.string().optional(), darkLine: z.string().optional()
    }).optional(),
    slides: z.array(z.object({
      type: z.string(), animation: z.string().optional(), content: z.record(z.string(), z.string())
    })).min(1)
  }
}, async (request) => text(buildDeckPlan(families(), request)));

const colorSchema = z.object({
  background: z.string(), text: z.string(), muted: z.string(), accent: z.string(), line: z.string(), darkBackground: z.string()
});
const nativeObjectFillSchema = z.object({
  name: z.string().min(1),
  kind: z.enum([
    'text', 'group_text', 'table_cell', 'smartart_text',
    'chart_title', 'chart_label', 'chart_data', 'image_placeholder', 'image_fill', 'shape_fill'
  ]),
  value: z.unknown(),
  locator: z.record(z.string(), z.unknown()).optional()
});
const pageSchema = z.object({
  target: z.string().optional(),
  page_index: z.number().int().positive(),
  content: z.record(z.string(), z.string()).default({}),
  object_fills: z.array(nativeObjectFillSchema).default([])
});
const compositionSchema = {
  base_target: z.string(),
  composition_mode: z.literal('native').default('native'),
  typography_theme_id: z.string().optional(),
  colors: colorSchema.optional(),
  color_map: z.record(z.string(), z.string()).optional(),
  pages: z.array(pageSchema).min(1)
};

server.registerTool('ppt_creater_list_composition_pages', {
  description: 'List production-ready native pages and their fillable slot capacities for template expansion and cross-target structural selection.',
  inputSchema: { target: z.string().optional() }
}, async ({ target } = {}) => {
  const catalog = compositionCatalog();
  if (target) {
    if (!catalog.targets[target]) throw new Error(`Unknown composition target: ${target}`);
    return text({ [target]: compositionPageCatalog({ targets: { [target]: catalog.targets[target] } })[target] });
  }
  return text(compositionPageCatalog(catalog));
});

server.registerTool('ppt_creater_plan_composition', {
  description: 'Plan a strict native-template composition. One production-ready family only; existing native text, groups, tables, SmartArt, charts, and images may be filled through registered object locators.',
  inputSchema: compositionSchema
}, async (request) => text(buildCompositionPlan(compositionCatalog(), request)));

server.registerTool('ppt_creater_compose_deck', {
  description: 'Copy registered native pages, fill approved native objects, then require full-page preview, PowerPoint opening, native normalization, animation playback, and QA before returning the editable PPTX.',
  inputSchema: { ...compositionSchema, output_name: z.string().min(6) }
}, async (request) => {
  const plan = buildCompositionPlan(compositionCatalog(), request);
  const output = outputPath(request.output_name);
  const powerpoint = new PowerPointComBackend(config);
  const status = powerpoint.status();
  if (status.selected_backend !== 'powerpoint-com') {
    throw new Error(`ppt_creater_compose_deck requires PowerPoint COM for full QA: ${status.reason || 'backend unavailable.'}`);
  }
  const result = await composeDeck({ plan, outputPath: output, powerpoint, qa: true });
  return text({
    output_path: result.output,
    slide_count: plan.pages.length,
    composition_mode: plan.composition_mode,
    base_target: plan.base_target,
    imported_targets: [...new Set(plan.pages.map((page) => page.target))],
    object_fill_result: result.object_fill_result,
    qa: {
      status: result.qa.status,
      reports: result.qa.reports,
      new_shapes: result.qa.qa.new_shapes,
      powerpoint_playback: result.qa.qa.powerpoint_playback,
      timeline_signature_match: result.qa.qa.timeline_signature_match
    }
  });
});

async function main() {
  await server.connect(new StdioServerTransport());
}
main().catch((error) => { console.error(error); process.exit(1); });
