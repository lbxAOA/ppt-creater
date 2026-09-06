const assert = require('node:assert/strict');
const test = require('node:test');
const {
  TYPOGRAPHY_ROLES,
  listTypographyThemes,
  resolveTypographyTheme,
  inferTypographyRole,
  createTypographyTheme
} = require('../src/typography-system');
const { buildDeckPlan } = require('../src/planner');

const family = {
  family_id: 'tech-navy-01',
  use_cases: ['AI行业研究'],
  visual_tags: ['深色'],
  aspect_ratio: '16:9',
  theme: { colors: ['172033', 'F5F8FC', '0B1F3A', 'DCE5EF', '00A6A6'], fonts: ['Aptos'] },
  animation_policy: { allowed_patterns: ['none'] },
  slides: [{ slide_id: 'cover', type: 'cover', animation_pattern: 'none', slots: ['title', 'subtitle'] }]
};

test('built-in typography themes expose coordinated roles for Chinese, Latin, and numbers', () => {
  const themes = listTypographyThemes();
  assert.ok(themes.length >= 6);
  assert.ok(themes.some((theme) => theme.id === 'modern-business'));
  const theme = resolveTypographyTheme('tech-precision');
  for (const role of TYPOGRAPHY_ROLES) {
    assert.equal(typeof theme.roles[role].fontFamily, 'string');
    assert.equal(typeof theme.roles[role].fontSize, 'number');
    assert.ok(theme.roles[role].fontWeight >= 400);
  }
  assert.equal(theme.roles.number.fontFamily, 'JetBrains Mono');
  assert.match(theme.roles.body.fallbackFontFamily, /Microsoft YaHei/);
});

test('typography role inference uses semantic slot names without changing content', () => {
  assert.equal(inferTypographyRole('slot_title'), 'title');
  assert.equal(inferTypographyRole('slot_kpi_number'), 'number');
  assert.equal(inferTypographyRole('slot_source_caption'), 'caption');
  assert.equal(inferTypographyRole('slot_unknown'), 'body');
});

test('custom typography themes reject missing roles and invalid values', () => {
  assert.throws(() => createTypographyTheme({ id: 'broken', name: 'Broken', roles: {} }), /requires roles/);
  assert.throws(() => createTypographyTheme({
    id: 'broken', name: 'Broken', roles: Object.fromEntries(TYPOGRAPHY_ROLES.map((role) => [role, { fontFamily: 'Arial', fontSize: 0, fontWeight: 400 }]))
  }), /fontSize/);
});

test('deck planning resolves an explicitly selected typography theme', () => {
  const plan = buildDeckPlan([family], {
    family_id: family.family_id,
    title: '具身智能',
    typography_theme_id: 'modern-business',
    slides: [{ type: 'cover', content: { title: '具身智能', subtitle: '技术与商业机会' } }]
  });
  assert.equal(plan.typography_theme.id, 'modern-business');
  assert.equal(plan.typography_theme.roles.title.fontWeight, 700);
});

test('deck planning keeps native typography when no preset is selected', () => {
  const plan = buildDeckPlan([family], {
    family_id: family.family_id,
    title: '具身智能',
    slides: [{ type: 'cover', content: { title: '具身智能', subtitle: '技术与商业机会' } }]
  });
  assert.equal(plan.typography_theme, null);
});
