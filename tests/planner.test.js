const assert = require('node:assert/strict');
const test = require('node:test');
const { chooseFamily, buildDeckPlan } = require('../src/planner');

const families = [
  {
    family_id: 'tech-navy-01',
    use_cases: ['AI行业研究', '技术方案', '投资人技术尽调'],
    visual_tags: ['深色', '科技', '正式'],
    aspect_ratio: '16:9',
    animation_policy: { allowed_patterns: ['fade', 'left_to_right_sequence'] },
    slides: [
      { slide_id: 'cover', type: 'cover', animation_pattern: 'fade', slots: ['title', 'subtitle'] },
      { slide_id: 'architecture', type: 'architecture_sequence', animation_pattern: 'left_to_right_sequence', slots: ['title', 'world_model'] },
      { slide_id: 'summary', type: 'summary', animation_pattern: 'none', slots: ['title', 'insight'] }
    ]
  },
  {
    family_id: 'academic-clean-01',
    use_cases: ['科研答辩'],
    visual_tags: ['浅色', '学术'],
    aspect_ratio: '16:9',
    animation_policy: { allowed_patterns: ['fade'] },
    slides: [{ slide_id: 'cover', type: 'cover', animation_pattern: 'fade', slots: ['title'] }]
  }
];

test('chooseFamily keeps a deck in the best matching family', () => {
  const family = chooseFamily(families, {
    use_case: 'AI行业研究',
    visual_tags: ['深色', '科技'],
    aspect_ratio: '16:9'
  });
  assert.equal(family.family_id, 'tech-navy-01');
});

test('buildDeckPlan rejects a requested animation not approved by the family', () => {
  assert.throws(() => buildDeckPlan(families, {
    family_id: 'tech-navy-01',
    title: 'Test',
    slides: [{ type: 'architecture_sequence', animation: 'spin', content: { title: 'A', world_model: 'B' } }]
  }), /not allowed/);
});

test('buildDeckPlan carries one resolved template theme for deck content and navigation', () => {
  const themedFamilies = [{
    ...families[0],
    theme: { colors: ['172033', 'F5F8FC', '0B1F3A', 'DCE5EF', '00A6A6'], fonts: ['Microsoft YaHei'] }
  }];

  const plan = buildDeckPlan(themedFamilies, {
    family_id: 'tech-navy-01',
    title: '具身智能',
    slides: [{ type: 'cover', content: { title: '具身智能', subtitle: '技术与商业机会' } }]
  });

  assert.equal(plan.color_theme.source.mode, 'template');
  assert.equal(plan.color_theme.colors.accent, '00A6A6');
  assert.equal(plan.color_theme.colors.nav.active, '00A6A6');
});
