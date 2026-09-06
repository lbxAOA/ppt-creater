const assert = require('node:assert/strict');
const test = require('node:test');
const { createColorTheme, resolveColor } = require('../src/color-system');
const { addNavigation, normalizeNavigation } = require('../src/navigation');

function makeSlide() {
  return {
    shapes: [],
    texts: [],
    addShape(type, options) { this.shapes.push({ type, options }); },
    addText(text, options) { this.texts.push({ text, options }); }
  };
}

test('color system derives every navigation color from semantic theme tokens', () => {
  const theme = createColorTheme({
    name: 'violet',
    colors: { accent: '7C3AED', background: 'F7F5FF', text: '211A33' }
  });

  assert.equal(resolveColor(theme, 'nav.active'), '7C3AED');
  assert.equal(resolveColor(theme, 'nav.track'), theme.colors.line);
  assert.equal(resolveColor(theme, 'nav.inactive'), theme.colors.muted);
  assert.equal(theme.colors.accentWeak.length, 6);
});

test('number-static renders a current-page count using semantic colors', () => {
  const slide = makeSlide();
  const theme = createColorTheme();
  const navigation = normalizeNavigation({ style: 'number-static', page: 3, total: 10 });

  addNavigation(slide, navigation, theme);

  assert.equal(slide.texts.at(-1).text, '03 / 10');
  assert.equal(slide.texts.at(-1).options.color, theme.colors.accent);
  assert.ok(slide.shapes.some((shape) => shape.options.fill?.color === theme.colors.accent));
});

test('rail-dynamic highlights the active section after final pagination', () => {
  const slide = makeSlide();
  const theme = createColorTheme();
  const navigation = normalizeNavigation({
    style: 'rail-dynamic',
    page: 6,
    total: 10,
    sections: [
      { title: '背景', start: 1, end: 2 },
      { title: '市场', start: 3, end: 5 },
      { title: '技术', start: 6, end: 8 },
      { title: '结论', start: 9, end: 10 }
    ]
  });

  addNavigation(slide, navigation, theme);

  const technicalLabel = slide.texts.find((entry) => entry.text === '技术');
  assert.equal(technicalLabel.options.color, theme.colors.accent);
  assert.ok(slide.texts.some((entry) => entry.text === '✓'));
  assert.ok(slide.shapes.some((shape) => shape.options.fill?.color === theme.colors.accent));
});

test('navigation rejects a page outside the final pagination range', () => {
  assert.throws(
    () => normalizeNavigation({ style: 'rail-static', page: 11, total: 10 }),
    /page must be an integer from 1 to total/
  );
});
