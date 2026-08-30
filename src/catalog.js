const fs = require('node:fs');
const path = require('node:path');

function loadCatalog(catalogPath) {
  if (!fs.existsSync(catalogPath)) return [];
  return fs.readFileSync(catalogPath, 'utf8')
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => JSON.parse(line));
}

function searchTemplates(catalog, query = {}) {
  const tokens = (query.text || '').toLowerCase().split(/\s+/).filter(Boolean);
  return catalog
    .filter((item) => item.status === 'ok')
    .filter((item) => !query.require_animation || item.has_animation_xml)
    .filter((item) => !query.require_slots || item.has_named_slots)
    .filter((item) => !query.min_slides || item.slide_count >= query.min_slides)
    .map((item) => {
      const haystack = JSON.stringify(item).toLowerCase();
      let score = 0;
      if (item.has_animation_xml) score += 4;
      if (item.has_named_slots) score += 4;
      score += Math.min(item.slide_count || 0, 20) / 10;
      score += tokens.filter((token) => haystack.includes(token)).length * 2;
      return { ...item, score: Number(score.toFixed(2)) };
    })
    .sort((a, b) => b.score - a.score)
    .slice(0, query.limit || 20);
}

function validateFamily(family) {
  if (!family.family_id || !family.source_template) throw new Error('Family requires family_id and source_template.');
  const ids = new Set();
  for (const slide of family.slides || []) {
    if (ids.has(slide.slide_id)) throw new Error(`Duplicate slide_id: ${slide.slide_id}`);
    ids.add(slide.slide_id);
    if (slide.animation_pattern && slide.animation_pattern !== 'none') {
      for (const slot of slide.slots || []) {
        if (!slot.animation_locked) {
          throw new Error(`Animated slide ${slide.slide_id}: slot ${slot.name} must be animation_locked.`);
        }
      }
    }
  }
}

function registerFamily(registry, family) {
  validateFamily(family);
  const copy = structuredClone(family);
  const existing = registry.families.findIndex((item) => item.family_id === copy.family_id);
  if (existing >= 0) registry.families[existing] = copy;
  else registry.families.push(copy);
  return registry;
}

function saveRegistry(filePath, registry) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, JSON.stringify(registry, null, 2) + '\n', 'utf8');
}

module.exports = { loadCatalog, searchTemplates, validateFamily, registerFamily, saveRegistry };
