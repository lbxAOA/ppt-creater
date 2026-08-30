const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { loadRegistry, resolveTarget } = require('../src/target-router');

const registry = loadRegistry(path.join(__dirname, '..', 'catalog', 'unified-production-template-library.json'));

test('routes academic work to the user-owned modern-report family', () => {
  const route = resolveTarget(registry, 'academic-clean');
  assert.equal(route.family_id, 'local-modern-report');
});

test('blocks external primary templates until external-license approval is supplied', () => {
  assert.throws(() => resolveTarget(registry, 'ai-industry'), /external-license approval/);
  const route = resolveTarget(registry, 'ai-industry', { allowExternal: true });
  assert.equal(route.family_id, 'external-ai-tool-pitch');
});

test('routes company profiles to the professional external family after approval', () => {
  const route = resolveTarget(registry, 'brand-company-profile', { allowExternal: true });
  assert.equal(route.family_id, 'external-company-professional');
});

test('routes wedding events to the user-owned album family without external approval', () => {
  const route = resolveTarget(registry, 'wedding-bridal');
  assert.equal(route.family_id, 'local-wedding-album');
});

test('rejects unregistered target categories', () => {
  assert.throws(() => resolveTarget(registry, 'unknown'), /Unknown generation target/);
});
