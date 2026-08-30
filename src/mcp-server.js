const fs = require('node:fs');
const path = require('node:path');
const { McpServer } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const { z } = require('zod');
const { loadCatalog, searchTemplates } = require('./catalog');
const { chooseFamily, buildDeckPlan } = require('./planner');

const ROOT = path.resolve(__dirname, '..');
const configPath = process.env.PPT_CREATER_CONFIG || path.join(ROOT, 'config', 'ppt-creater.config.json');
const familyPath = process.env.PPT_CREATER_FAMILIES || path.join(ROOT, 'catalog', 'families.json');
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

const server = new McpServer({ name: 'ppt-creater', version: '0.1.0' });

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
  note: config.library_root ? 'Ready to scan or search.' : 'Set library_root after the template library is provided.'
}));

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
    slides: z.array(z.object({
      type: z.string(),
      animation: z.string().optional(),
      content: z.record(z.string(), z.string())
    })).min(1)
  }
}, async (request) => text(buildDeckPlan(families(), request)));

// This server creates and validates plans. The safe template filler remains local
// until a real template family has passed playback verification.
// Use src/fill-template.js after named slot_* shapes and an approved registry exist.

async function main() {
  await server.connect(new StdioServerTransport());
}
main().catch((error) => { console.error(error); process.exit(1); });
