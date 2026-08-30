const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const { loadCatalog, searchTemplates, registerFamily } = require('../src/catalog');

test('searchTemplates prioritizes animated and slot-ready templates', () => {
  const catalog = [
    { relative_path: 'plain.pptx', status: 'ok', slide_count: 4, has_animation_xml: false, has_named_slots: false, shape_kinds: {} },
    { relative_path: 'animated.pptx', status: 'ok', slide_count: 8, has_animation_xml: true, has_named_slots: true, shape_kinds: { chart: 2 } }
  ];
  const results = searchTemplates(catalog, { require_animation: true, require_slots: true, limit: 5 });
  assert.equal(results.length, 1);
  assert.equal(results[0].relative_path, 'animated.pptx');
});

test('registerFamily rejects animation slots that are not marked locked', () => {
  assert.throws(() => registerFamily({ families: [] }, {
    family_id: 'broken',
    source_template: 'a.pptx',
    slides: [{ slide_id: 'architecture', type: 'architecture_sequence', animation_pattern: 'left_to_right_sequence', slots: [{ name: 'world_model', shape_name: 'slot_world_model', animation_locked: false }] }]
  }), /must be animation_locked/);
});

test('loadCatalog reads JSONL entries', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-creater-'));
  const catalogPath = path.join(dir, 'catalog.jsonl');
  fs.writeFileSync(catalogPath, '{"relative_path":"a.pptx","status":"ok"}\n', 'utf8');
  assert.deepEqual(loadCatalog(catalogPath), [{ relative_path: 'a.pptx', status: 'ok' }]);
});
