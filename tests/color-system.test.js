const assert = require('node:assert/strict');
const test = require('node:test');
const { createTemplateColorTheme, createUserColorTheme, createLegacyReportColors, resolveDeckColorTheme } = require('../src/color-system');

test('template color mode derives the navigation theme from the selected template palette', () => {
  const family = {
    family_id: 'template-blue',
    theme: {
      colors: ['1F2937', 'FFFFFF', '334155', 'E2E8F0', '2563EB', 'F59E0B'],
      fonts: ['Aptos']
    }
  };

  const theme = createTemplateColorTheme(family);

  assert.equal(theme.source.mode, 'template');
  assert.equal(theme.colors.background, 'FFFFFF');
  assert.equal(theme.colors.text, '1F2937');
  assert.equal(theme.colors.darkBackground, '334155');
  assert.equal(theme.colors.accent, '2563EB');
  assert.equal(theme.colors.nav.active, '2563EB');
  assert.equal(theme.fontFace, 'Aptos');
});

test('user color mode applies one semantic palette to content and navigation', () => {
  const family = { family_id: 'template-blue', theme: { colors: ['1F2937', 'FFFFFF', '334155', 'E2E8F0', '2563EB'] } };

  const theme = resolveDeckColorTheme({
    family,
    colorMode: 'user',
    colors: { background: 'FFF7ED', text: '431407', muted: '9A3412', accent: 'EA580C', line: 'FED7AA', darkBackground: '7C2D12' }
  });

  assert.equal(theme.source.mode, 'user');
  assert.equal(theme.colors.accent, 'EA580C');
  assert.equal(theme.colors.nav.active, 'EA580C');
  assert.equal(theme.colors.nav.track, 'FED7AA');
  assert.equal(theme.colors.text, '431407');
});

test('user color mode rejects incomplete palettes instead of silently mixing template colors', () => {
  assert.throws(() => resolveDeckColorTheme({
    family: { family_id: 'template-blue', theme: { colors: ['1F2937', 'FFFFFF', '334155', 'E2E8F0', '2563EB'] } },
    colorMode: 'user',
    colors: { accent: 'EA580C' }
  }), /requires all semantic color tokens/);
});

test('template color mode rejects a family without a usable native palette', () => {
  assert.throws(() => createTemplateColorTheme({ family_id: 'missing-theme' }), /no usable theme colors/);
});

test('legacy report colors are derived from the same semantic theme', () => {
  const theme = createUserColorTheme({
    colors: { background: 'FFF7ED', text: '431407', muted: '9A3412', accent: 'EA580C', line: 'FED7AA', darkBackground: '7C2D12' }
  });

  const colors = createLegacyReportColors(theme);

  assert.equal(colors.bg, theme.colors.background);
  assert.equal(colors.navy, theme.colors.darkBackground);
  assert.equal(colors.ink, theme.colors.text);
  assert.equal(colors.accent, theme.colors.accent);
  assert.equal(colors.line, theme.colors.line);
  assert.equal(colors.orange, theme.colors.accentMuted);
  assert.equal(colors.red, theme.colors.accentMuted);
});
