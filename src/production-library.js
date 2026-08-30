const fs = require('node:fs');

function loadProductionLibrary(registryPath) {
  const raw = fs.readFileSync(registryPath, 'utf8').replace(/^\uFEFF/, '');
  return JSON.parse(raw);
}

function assertVerifiedTemplate(item, target) {
  if (!item) throw new Error(`No production template for target: ${target}`);
  if (item.status !== 'production_ready') throw new Error(`Template is not production-ready: ${target}`);
  if (item.strict_native_only !== true || item.new_shapes !== 0) {
    throw new Error(`Template violates strict-native policy: ${target}`);
  }
  if (item.powerpoint_playback !== true) throw new Error(`Template playback is not verified: ${target}`);
  if (!item.master_pptx || !fs.existsSync(item.master_pptx)) throw new Error(`Master PPTX is unavailable: ${target}`);
}

function chooseProductionTemplate(registry, target, { allowExternal = false, templateId } = {}) {
  const templates = registry?.templates;
  const routes = registry?.routes;

  // Backward compatibility for the previous local-only registry.
  if (!templates || !routes) {
    const item = registry?.targets?.[target];
    if (!item) throw new Error(`No production template for target: ${target}`);
    if (item.animation_status !== 'verified_preserved') throw new Error(`Template animation has not been verified: ${target}`);
    if (!item.master_pptx || !fs.existsSync(item.master_pptx)) throw new Error(`Master PPTX is unavailable: ${target}`);
    if (!item.slot_count) throw new Error(`No text slots were mapped: ${target}`);
    return item;
  }

  const route = routes[target];
  if (!route) throw new Error(`No production template for target: ${target}`);
  const selectedId = templateId || route.primary;
  if (![route.primary, ...(route.alternatives || [])].includes(selectedId)) {
    throw new Error(`Template ${selectedId} is not registered for target: ${target}`);
  }
  const item = templates[selectedId];
  assertVerifiedTemplate(item, target);
  if (item.template_origin === 'external' && !allowExternal) {
    throw new Error(`Template ${selectedId} requires external-license approval.`);
  }
  return item;
}

module.exports = { loadProductionLibrary, chooseProductionTemplate };
