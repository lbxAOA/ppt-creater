const { ModifyTextHelper, XmlHelper } = require('pptx-automizer');

const TYPOGRAPHY_ROLES = Object.freeze(['title', 'subtitle', 'heading', 'body', 'caption', 'number']);

function validateRole(role, roleName) {
  if (!role || typeof role !== 'object') throw new Error(`Typography role ${roleName} must be an object.`);
  if (typeof role.fontFamily !== 'string' || !role.fontFamily.trim()) {
    throw new Error(`Typography role ${roleName} requires fontFamily.`);
  }
  if (!Number.isFinite(role.fontSize) || role.fontSize <= 0) {
    throw new Error(`Typography role ${roleName} fontSize must be positive.`);
  }
  if (!Number.isInteger(role.fontWeight) || role.fontWeight < 100 || role.fontWeight > 900) {
    throw new Error(`Typography role ${roleName} fontWeight must be an integer from 100 to 900.`);
  }
  if (!Number.isFinite(role.lineHeight) || role.lineHeight <= 0) {
    throw new Error(`Typography role ${roleName} lineHeight must be positive.`);
  }
}

function createTypographyTheme({ id, name, description = '', category = 'general', roles } = {}) {
  if (typeof id !== 'string' || !id.trim()) throw new Error('Typography theme requires id.');
  if (typeof name !== 'string' || !name.trim()) throw new Error('Typography theme requires name.');
  const missing = TYPOGRAPHY_ROLES.filter((role) => !roles || !roles[role]);
  if (missing.length) throw new Error(`Typography theme requires roles: ${missing.join(', ')}.`);
  const normalizedRoles = {};
  for (const roleName of TYPOGRAPHY_ROLES) {
    const role = { ...roles[roleName] };
    validateRole(role, roleName);
    role.fontFamily = role.fontFamily.trim();
    role.eastAsiaFontFamily = (role.eastAsiaFontFamily || 'Noto Sans SC').trim();
    role.latinFontFamily = (role.latinFontFamily || role.fontFamily).trim();
    if (role.fallbackFontFamily) role.fallbackFontFamily = role.fallbackFontFamily.trim();
    else role.fallbackFontFamily = `${role.eastAsiaFontFamily}, Microsoft YaHei, Arial, sans-serif`;
    normalizedRoles[roleName] = Object.freeze(role);
  }
  return Object.freeze({
    id,
    name,
    description,
    category,
    roles: Object.freeze(normalizedRoles),
    fonts: Object.freeze({
      heading: normalizedRoles.title.fontFamily,
      body: normalizedRoles.body.fontFamily,
      caption: normalizedRoles.caption.fontFamily,
      number: normalizedRoles.number.fontFamily
    })
  });
}

const role = (fontFamily, eastAsiaFontFamily, fontSize, fontWeight, lineHeight, extra = {}) => ({
  fontFamily,
  eastAsiaFontFamily,
  fontSize,
  fontWeight,
  lineHeight,
  ...extra
});

const BUILT_IN_THEMES = Object.freeze([
  createTypographyTheme({
    id: 'modern-business', name: '现代商务', category: 'business',
    description: '清晰、稳重，适合商业汇报、项目总结与公司介绍。',
    roles: {
      title: role('Inter', 'Noto Sans SC', 30, 700, 1.2, { letterSpacing: -0.3 }),
      subtitle: role('Inter', 'Noto Sans SC', 17, 500, 1.4),
      heading: role('Inter', 'Noto Sans SC', 22, 700, 1.3),
      body: role('Inter', 'Noto Sans SC', 16, 400, 1.55),
      caption: role('Inter', 'Noto Sans SC', 11, 400, 1.4),
      number: role('Inter', 'Noto Sans SC', 28, 700, 1.15)
    }
  }),
  createTypographyTheme({
    id: 'tech-precision', name: '极简科技', category: 'tech',
    description: '高对比、强层级，适合 AI、技术架构与数据产品。',
    roles: {
      title: role('Space Grotesk', 'Noto Sans SC', 32, 700, 1.15, { letterSpacing: -0.5 }),
      subtitle: role('Inter', 'Noto Sans SC', 17, 400, 1.4),
      heading: role('Space Grotesk', 'Noto Sans SC', 23, 700, 1.25),
      body: role('Inter', 'Noto Sans SC', 15, 400, 1.55),
      caption: role('Inter', 'Noto Sans SC', 10, 400, 1.35),
      number: role('JetBrains Mono', 'Noto Sans SC', 29, 700, 1.1)
    }
  }),
  createTypographyTheme({
    id: 'editorial-serif', name: '人文杂志', category: 'editorial',
    description: '标题有质感、正文易阅读，适合品牌故事与内容展示。',
    roles: {
      title: role('Noto Serif SC', 'Noto Serif SC', 31, 700, 1.25),
      subtitle: role('Noto Sans SC', 'Noto Sans SC', 17, 400, 1.5),
      heading: role('Noto Serif SC', 'Noto Serif SC', 23, 700, 1.3),
      body: role('Noto Sans SC', 'Noto Sans SC', 16, 400, 1.65),
      caption: role('Noto Sans SC', 'Noto Sans SC', 11, 400, 1.45),
      number: role('Inter', 'Noto Sans SC', 28, 600, 1.2)
    }
  }),
  createTypographyTheme({
    id: 'premium-finance', name: '高端金融', category: 'finance',
    description: '克制、正式，适合投资人材料、董事会与财务报告。',
    roles: {
      title: role('Noto Serif SC', 'Noto Serif SC', 29, 700, 1.25),
      subtitle: role('Noto Sans SC', 'Noto Sans SC', 16, 500, 1.45),
      heading: role('Noto Sans SC', 'Noto Sans SC', 21, 700, 1.3),
      body: role('Noto Sans SC', 'Noto Sans SC', 15, 400, 1.55),
      caption: role('Noto Sans SC', 'Noto Sans SC', 10, 400, 1.4),
      number: role('IBM Plex Mono', 'Noto Sans SC', 27, 700, 1.15)
    }
  }),
  createTypographyTheme({
    id: 'vibrant-creative', name: '活力创意', category: 'creative',
    description: '轻快、有节奏，适合营销、活动与年轻产品发布。',
    roles: {
      title: role('Manrope', 'Noto Sans SC', 33, 800, 1.15),
      subtitle: role('Manrope', 'Noto Sans SC', 18, 500, 1.4),
      heading: role('Manrope', 'Noto Sans SC', 24, 700, 1.25),
      body: role('Inter', 'Noto Sans SC', 16, 400, 1.55),
      caption: role('Inter', 'Noto Sans SC', 11, 400, 1.4),
      number: role('Manrope', 'Noto Sans SC', 31, 800, 1.1)
    }
  }),
  createTypographyTheme({
    id: 'education-clear', name: '教育清晰', category: 'education',
    description: '阅读负担低、层级明确，适合课程、培训与学术汇报。',
    roles: {
      title: role('Noto Sans SC', 'Noto Sans SC', 30, 700, 1.25),
      subtitle: role('Noto Sans SC', 'Noto Sans SC', 18, 400, 1.5),
      heading: role('Noto Sans SC', 'Noto Sans SC', 22, 700, 1.35),
      body: role('Noto Sans SC', 'Noto Sans SC', 17, 400, 1.65),
      caption: role('Noto Sans SC', 'Noto Sans SC', 11, 400, 1.45),
      number: role('Inter', 'Noto Sans SC', 29, 700, 1.2)
    }
  })
]);

function listTypographyThemes() {
  return BUILT_IN_THEMES.map(({ id, name, description, category, fonts }) => ({
    id, name, description, category, fonts: { ...fonts }
  }));
}

function resolveTypographyTheme(themeId) {
  if (!themeId) return null;
  const theme = BUILT_IN_THEMES.find((item) => item.id === themeId);
  if (!theme) throw new Error(`Unknown typography theme: ${themeId}`);
  return theme;
}

function inferTypographyRole(slotName = '') {
  const value = String(slotName).toLowerCase();
  if (/(^|[_-])(title|headline|cover)([_-]|$)/.test(value)) return 'title';
  if (/(subtitle|sub_title|tagline|deck)/.test(value)) return 'subtitle';
  if (/(number|kpi|metric|stat|percent|percentage|amount|count|value)/.test(value)) return 'number';
  if (/(caption|source|footnote|citation|note|eyebrow|label)/.test(value)) return 'caption';
  if (/(heading|section|chapter)/.test(value)) return 'heading';
  return 'body';
}

function setTypeface(element, tagName, value) {
  let child = XmlHelper.getFirstDirectChild(element, [tagName]);
  if (!child) {
    child = element.ownerDocument.createElement(tagName);
    XmlHelper.insertInSchemaOrder(element, child, [
      'a:ln', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill',
      'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn', 'a:uFillTx', 'a:uFill',
      'a:latin', 'a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst'
    ]);
  }
  child.setAttribute('typeface', String(value));
}

function textRunProperties(element) {
  const properties = [
    ...Array.from(element.getElementsByTagName('a:rPr')),
    ...Array.from(element.getElementsByTagName('a:defRPr')),
    ...Array.from(element.getElementsByTagName('a:endParaRPr'))
  ];
  const runs = Array.from(element.getElementsByTagName('a:r'));
  for (const run of runs) {
    if (XmlHelper.getFirstDirectChild(run, ['a:rPr'])) continue;
    const property = element.ownerDocument.createElement('a:rPr');
    const textNode = XmlHelper.getFirstDirectChild(run, ['a:t']);
    if (textNode) run.insertBefore(property, textNode);
    else run.appendChild(property);
    properties.push(property);
  }
  return properties;
}

function createTypographyModifier(theme, slotName) {
  if (!theme) return null;
  const selectedTheme = typeof theme === 'string' ? resolveTypographyTheme(theme) : theme;
  const roleName = TYPOGRAPHY_ROLES.includes(slotName) ? slotName : inferTypographyRole(slotName);
  const style = selectedTheme.roles[roleName];
  return (element) => {
    const runStyle = {
      fontFamily: style.latinFontFamily || style.fontFamily,
      size: Math.round(style.fontSize * 100),
      isBold: style.fontWeight >= 600,
      isItalics: Boolean(style.italic)
    };
    for (const property of textRunProperties(element)) {
      ModifyTextHelper.style(runStyle)(property);
      setTypeface(property, 'a:ea', style.eastAsiaFontFamily || style.fontFamily);
      setTypeface(property, 'a:cs', style.eastAsiaFontFamily || style.fontFamily);
    }
  };
}

module.exports = {
  TYPOGRAPHY_ROLES,
  BUILT_IN_THEMES,
  createTypographyTheme,
  listTypographyThemes,
  resolveTypographyTheme,
  inferTypographyRole,
  createTypographyModifier
};
