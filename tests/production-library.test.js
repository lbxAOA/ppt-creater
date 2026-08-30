const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { loadProductionLibrary, chooseProductionTemplate } = require('../src/production-library');

const registry = loadProductionLibrary(path.join(__dirname, '..', 'catalog', 'unified-production-template-library.json'));

test('chooses a user-owned animation-preserved native master without an external license override', () => {
  const item = chooseProductionTemplate(registry, 'academic-clean');
  assert.equal(item.family_id, 'local-modern-report');
  assert.equal(item.template_origin, 'local');
  assert.ok(item.master_pptx.endsWith('local-modern-report\\master.pptx'));
});

test('requires explicit external-license approval before selecting an external family', () => {
  assert.throws(() => chooseProductionTemplate(registry, 'ai-industry'), /external-license approval/);
  const item = chooseProductionTemplate(registry, 'ai-industry', { allowExternal: true });
  assert.equal(item.family_id, 'external-ai-tool-pitch');
  assert.equal(item.template_origin, 'external');
});

test('allows a registered route alternative only with explicit external-license approval', () => {
  const item = chooseProductionTemplate(registry, 'technical-product', {
    allowExternal: true,
    templateId: 'external-ai-tool-pitch'
  });
  assert.equal(item.family_id, 'external-ai-tool-pitch');
  assert.throws(
    () => chooseProductionTemplate(registry, 'technical-product', { allowExternal: true, templateId: 'local-modern-report' }),
    /not registered/
  );
});

test('refuses a target not provisioned in the production library', () => {
  assert.throws(() => chooseProductionTemplate(registry, 'unlisted'), /No production template/);
});
