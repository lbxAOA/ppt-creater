const path = require('node:path');
const pptxgen = require('pptxgenjs');
const { createColorTheme } = require('./color-system');
const { addNavigation } = require('./navigation');

const OUT = path.resolve(__dirname, '..', 'output', 'navigation-style-gallery.pptx');
const pptx = new pptxgen();
pptx.defineLayout({ name: 'CUSTOM_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'CUSTOM_WIDE';
pptx.author = 'Hermes / ppt-creater';
pptx.title = '功能导航栏样式库';
pptx.lang = 'zh-CN';

const themes = [
  createColorTheme({ name: 'teal', colors: { accent: '00A6A6', background: 'F5F8FC', text: '172033' } }),
  createColorTheme({ name: 'violet', colors: { accent: '7C3AED', background: 'F7F5FF', text: '211A33' } }),
  createColorTheme({ name: 'coral', colors: { accent: 'E4573D', background: 'FFF8F6', text: '30201C' } }),
  createColorTheme({ name: 'blue', colors: { accent: '2563EB', background: 'F6F9FF', text: '172554' } })
];
const specs = [
  { style: 'number-static', title: '数字 · 静态', detail: '统一页码；当前页及页数清晰可见。', page: 3, total: 10 },
  { style: 'number-dynamic', title: '数字 · 章节状态', detail: '按最终分页识别章节，并突出当前章节。', page: 6, total: 10 },
  { style: 'rail-static', title: '滑轨 · 静态', detail: '整份内容的连续进度；适合长报告。', page: 7, total: 10 },
  { style: 'rail-dynamic', title: '滑轨 · 章节状态', detail: '章节锚点、完成状态与当前位置同时呈现。', page: 6, total: 10 }
];
const sections = [
  { title: '背景', start: 1, end: 2 },
  { title: '市场', start: 3, end: 5 },
  { title: '技术', start: 6, end: 8 },
  { title: '结论', start: 9, end: 10 }
];

function addText(slide, text, options) {
  slide.addText(text, { margin: 0, breakLine: false, fit: 'shrink', fontFace: 'Microsoft YaHei', ...options });
}

function addContent(slide, theme, spec) {
  slide.background = { color: theme.colors.background };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.18, h: 7.5, fill: { color: theme.colors.accent }, line: { color: theme.colors.accent } });
  addText(slide, 'PPT-CREATER · NAVIGATION SYSTEM', { x: 0.7, y: 0.62, w: 5.2, h: 0.2, fontSize: 9, bold: true, charSpace: 1.3, color: theme.colors.accent });
  addText(slide, spec.title, { x: 0.7, y: 1.08, w: 8, h: 0.55, fontSize: 28, bold: true, color: theme.colors.text });
  addText(slide, spec.detail, { x: 0.7, y: 1.78, w: 8.5, h: 0.25, fontSize: 13, color: theme.colors.muted });
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.7, y: 2.65, w: 11.8, h: 2.55, rectRadius: 0.08, fill: { color: theme.colors.surface }, line: { color: theme.colors.line, pt: 0.8 } });
  addText(slide, '内容由 AI 完成大纲与分页后，再统一注入导航状态。', { x: 1.06, y: 3.12, w: 8.9, h: 0.3, fontSize: 16, bold: true, color: theme.colors.text });
  addText(slide, '颜色不在组件内硬编码：活动、完成、弱化、轨道和文字均来自同一语义颜色系统。\n更换主题后，正文、图表和导航栏会一起保持同色系。', { x: 1.06, y: 3.75, w: 8.9, h: 0.65, fontSize: 12.5, color: theme.colors.muted, breakLine: true });
  slide.addShape(pptx.ShapeType.roundRect, { x: 10.52, y: 3.05, w: 1.42, h: 1.42, rectRadius: 0.1, fill: { color: theme.colors.accentWeak }, line: { color: theme.colors.accentWeak } });
  addText(slide, String(spec.page).padStart(2, '0'), { x: 10.52, y: 3.42, w: 1.42, h: 0.26, fontSize: 20, bold: true, color: theme.colors.accent, align: 'center' });
  addText(slide, `主题：${theme.name}`, { x: 0.7, y: 5.62, w: 2.5, h: 0.18, fontSize: 8.5, bold: true, color: theme.colors.accent });
  addNavigation(slide, { ...spec, sections, placement: 'bottom' }, theme);
}

specs.forEach((spec, index) => addContent(pptx.addSlide(), themes[index], spec));
pptx.writeFile({ fileName: OUT });
