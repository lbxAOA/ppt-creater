const fs = require('node:fs');

function loadRegistry(registryPath) {
  return JSON.parse(fs.readFileSync(registryPath, 'utf8').replace(/^\uFEFF/, ''));
}

function resolveTarget(registry, target, { allowExternal = false, templateId } = {}) {
  // Support the historical visual-review registry while the caller migrates.
  if (!registry.routes || !registry.templates) {
    const group = registry.targets[target];
    if (!group) throw new Error(`Unknown generation target: ${target}`);
    const candidates = [...(group.primary || []), ...(group.secondary || [])];
    if (!candidates.length) throw new Error(`No template candidate is registered for: ${target}`);
    return { target, family_id: candidates[0], primary: group.primary || [], secondary: group.secondary || [], notes: group.notes || '' };
  }

  const route = registry.routes[target];
  if (!route) throw new Error(`Unknown generation target: ${target}`);
  const familyId = templateId || route.primary;
  const candidates = [route.primary, ...(route.alternatives || [])];
  if (!candidates.includes(familyId)) throw new Error(`Template ${familyId} is not registered for: ${target}`);
  const family = registry.templates[familyId];
  if (!family || family.status !== 'production_ready') throw new Error(`Template is not production-ready: ${familyId}`);
  if (family.template_origin === 'external' && !allowExternal) {
    throw new Error(`Template ${familyId} requires external-license approval.`);
  }
  return { target, family_id: familyId, primary: route.primary, alternatives: route.alternatives || [], family };
}

module.exports = { loadRegistry, resolveTarget };
