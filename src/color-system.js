const DEFAULT_COLORS = Object.freeze({
  background: 'F5F8FC',
  surface: 'FFFFFF',
  text: '172033',
  muted: '64748B',
  accent: '00A6A6',
  line: 'DCE5EF',
  inverseText: 'FFFFFF',
  darkBackground: '0B1F3A',
  darkLine: '315476'
});

function normalizeHex(value, name) {
  if (typeof value !== 'string' || !/^[0-9A-Fa-f]{6}$/.test(value)) {
    throw new Error(`${name} must be a six-digit hex color without #.`);
  }
  return value.toUpperCase();
}

function mixHex(base, target, amount) {
  const from = normalizeHex(base, 'base color');
  const to = normalizeHex(target, 'target color');
  const ratio = Math.max(0, Math.min(1, Number(amount)));
  return [0, 2, 4].map((offset) => {
    const start = Number.parseInt(from.slice(offset, offset + 2), 16);
    const end = Number.parseInt(to.slice(offset, offset + 2), 16);
    return Math.round(start + (end - start) * ratio).toString(16).padStart(2, '0');
  }).join('').toUpperCase();
}

function createColorTheme({ name = 'default', fontFace = 'Microsoft YaHei', colors = {}, source = { mode: 'default' } } = {}) {
  const base = {};
  for (const [key, value] of Object.entries({ ...DEFAULT_COLORS, ...colors })) {
    base[key] = normalizeHex(value, `colors.${key}`);
  }
  return {
    name,
    fontFace,
    source,
    colors: {
      ...base,
      accentWeak: mixHex(base.accent, base.background, 0.84),
      accentMuted: mixHex(base.accent, base.muted, 0.48),
      nav: {
        active: base.accent,
        completed: mixHex(base.accent, base.background, 0.35),
        inactive: base.muted,
        track: base.line,
        label: base.muted,
        onActive: base.inverseText
      }
    }
  };
}

function colorAt(colors, index, fallback) {
  const value = colors[index];
  return typeof value === 'string' && /^[0-9A-Fa-f]{6}$/.test(value) ? value : fallback;
}

function createTemplateColorTheme(family) {
  const palette = family?.theme?.colors;
  if (!Array.isArray(palette) || palette.length < 5) {
    throw new Error(`Template family ${family?.family_id || '(unknown)'} has no usable theme colors.`);
  }
  const [text, background, darkBackground, line, accent] = palette;
  return createColorTheme({
    name: family.family_id || 'template',
    fontFace: family.theme?.fonts?.[0] || 'Microsoft YaHei',
    source: { mode: 'template', familyId: family.family_id || null },
    colors: {
      background,
      surface: background,
      text,
      muted: colorAt(palette, 2, DEFAULT_COLORS.muted),
      accent,
      line,
      inverseText: background,
      darkBackground,
      darkLine: colorAt(palette, 3, DEFAULT_COLORS.darkLine)
    }
  });
}

const USER_COLOR_TOKENS = Object.freeze(['background', 'text', 'muted', 'accent', 'line', 'darkBackground']);

function createUserColorTheme({ name = 'user', fontFace, colors } = {}) {
  const missing = USER_COLOR_TOKENS.filter((token) => !colors || !colors[token]);
  if (missing.length) {
    throw new Error(`User color mode requires all semantic color tokens: ${USER_COLOR_TOKENS.join(', ')}.`);
  }
  return createColorTheme({
    name,
    fontFace,
    source: { mode: 'user' },
    colors: {
      ...colors,
      surface: colors.surface || colors.background,
      inverseText: colors.inverseText || DEFAULT_COLORS.inverseText,
      darkLine: colors.darkLine || mixHex(colors.darkBackground, colors.background, 0.45)
    }
  });
}

function resolveDeckColorTheme({ family, colorMode = 'template', colors, name, fontFace } = {}) {
  if (colorMode === 'template') return createTemplateColorTheme(family);
  if (colorMode === 'user') return createUserColorTheme({ name, fontFace: fontFace || family?.theme?.fonts?.[0], colors });
  throw new Error(`Unsupported color mode: ${colorMode}`);
}

function createLegacyReportColors(theme) {
  if (!theme?.colors?.nav) throw new Error('theme must provide semantic colors.');
  const { colors } = theme;
  return {
    bg: colors.background,
    navy: colors.darkBackground,
    navy2: mixHex(colors.darkBackground, colors.accent, 0.22),
    ink: colors.text,
    muted: colors.muted,
    accent: colors.accent,
    accentLight: colors.accentWeak,
    orange: colors.reportAccentSecondary || colors.accentMuted,
    line: colors.line,
    white: colors.inverseText,
    red: colors.reportRisk || colors.accentMuted,
    darkMuted: mixHex(colors.inverseText, colors.darkBackground, 0.32),
    darkAccent: mixHex(colors.inverseText, colors.accent, 0.42),
    darkLine: colors.darkLine
  };
}

function resolveColor(theme, token) {
  const value = token.split('.').reduce((current, key) => current?.[key], theme.colors);
  if (typeof value !== 'string') throw new Error(`Unknown color token: ${token}`);
  return value;
}

module.exports = {
  createColorTheme,
  createTemplateColorTheme,
  createUserColorTheme,
  resolveDeckColorTheme,
  createLegacyReportColors,
  resolveColor,
  mixHex
};
