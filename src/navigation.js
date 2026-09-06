const VALID_STYLES = new Set(['number-static', 'number-dynamic', 'rail-static', 'rail-dynamic']);

function normalizeNavigation(input = {}) {
  const style = input.style || 'rail-static';
  if (!VALID_STYLES.has(style)) throw new Error(`Unsupported navigation style: ${style}`);
  const page = input.page;
  const total = input.total;
  if (!Number.isInteger(total) || total < 1) throw new Error('total must be a positive integer.');
  if (!Number.isInteger(page) || page < 1 || page > total) {
    throw new Error('page must be an integer from 1 to total.');
  }
  const placement = input.placement || 'bottom';
  if (!['top', 'bottom'].includes(placement)) throw new Error('placement must be top or bottom.');
  const sections = (input.sections || []).map((section) => ({ ...section }));
  for (const section of sections) {
    if (!section.title || !Number.isInteger(section.start) || !Number.isInteger(section.end)
      || section.start < 1 || section.end < section.start || section.end > total) {
      throw new Error('Each section needs title, start, and end within the final page range.');
    }
  }
  return { style, page, total, placement, sections };
}

function addShape(slide, type, options) {
  slide.addShape(type, options);
}

function addText(slide, text, options) {
  slide.addText(text, { margin: 0, breakLine: false, fit: 'shrink', ...options });
}

function position(navigation) {
  return navigation.placement === 'bottom'
    ? { y: 6.82, labelY: 6.95 }
    : { y: 0.26, labelY: 0.42 };
}

function activeSectionIndex(navigation) {
  return navigation.sections.findIndex((section) => navigation.page >= section.start && navigation.page <= section.end);
}

function addNumberNavigation(slide, navigation, theme) {
  const { y } = position(navigation);
  const colors = theme.colors.nav;
  const dynamic = navigation.style === 'number-dynamic' && navigation.sections.length > 0;
  if (!dynamic) {
    addShape(slide, 'line', { x: 0.62, y, w: 11.1, h: 0, line: { color: colors.track, pt: 0.65 } });
    addShape(slide, 'ellipse', { x: 11.92, y: y - 0.12, w: 0.46, h: 0.28, fill: { color: colors.active }, line: { color: colors.active } });
    addText(slide, `${String(navigation.page).padStart(2, '0')} / ${String(navigation.total).padStart(2, '0')}`, {
      x: 11.25, y: y - 0.06, w: 1.05, h: 0.14, fontFace: theme.fontFace || 'Microsoft YaHei',
      fontSize: 8, bold: true, color: colors.active, align: 'left'
    });
    return;
  }

  const active = activeSectionIndex(navigation);
  const startX = 0.68;
  const usableWidth = 11.9;
  const cell = usableWidth / navigation.sections.length;
  navigation.sections.forEach((section, index) => {
    const isActive = index === active;
    const isCompleted = active >= 0 && index < active;
    const color = isActive ? colors.active : isCompleted ? colors.completed : colors.inactive;
    const x = startX + index * cell;
    addShape(slide, 'ellipse', { x, y: y - 0.11, w: 0.26, h: 0.16, fill: { color }, line: { color } });
    if (isCompleted) {
      addText(slide, '✓', { x: x + 0.035, y: y - 0.105, w: 0.19, h: 0.12, fontFace: theme.fontFace || 'Microsoft YaHei', fontSize: 6.2, bold: true, color: colors.onActive, align: 'center' });
    }
    if (index < navigation.sections.length - 1) {
      addShape(slide, 'line', { x: x + 0.3, y: y - 0.03, w: cell - 0.48, h: 0, line: { color: index < active ? colors.completed : colors.track, pt: 0.75 } });
    }
    addText(slide, `${String(index + 1).padStart(2, '0')} ${section.title}`, {
      x: x + 0.36, y: y - 0.11, w: cell - 0.42, h: 0.16, fontFace: theme.fontFace || 'Microsoft YaHei',
      fontSize: 7.4, bold: isActive, color, align: 'left'
    });
  });
}

function addRailNavigation(slide, navigation, theme) {
  const { y, labelY } = position(navigation);
  const colors = theme.colors.nav;
  const startX = 0.7;
  const width = 11.93;
  const dynamic = navigation.style === 'rail-dynamic' && navigation.sections.length > 0;
  addShape(slide, 'line', { x: startX, y, w: width, h: 0, line: { color: colors.track, pt: 2 } });

  if (!dynamic) {
    const progress = (navigation.page - 1) / Math.max(1, navigation.total - 1);
    addShape(slide, 'line', { x: startX, y, w: width * progress, h: 0, line: { color: colors.active, pt: 2 } });
    const markerX = startX + width * progress - 0.09;
    addShape(slide, 'ellipse', { x: markerX, y: y - 0.1, w: 0.18, h: 0.18, fill: { color: colors.active }, line: { color: colors.active } });
    addText(slide, `${String(navigation.page).padStart(2, '0')} / ${String(navigation.total).padStart(2, '0')}`, {
      x: 11.55, y: labelY, w: 1.1, h: 0.13, fontFace: theme.fontFace || 'Microsoft YaHei',
      fontSize: 7.5, bold: true, color: colors.active, align: 'right'
    });
    return;
  }

  const active = activeSectionIndex(navigation);
  const inactiveLabel = theme.colors.text;
  navigation.sections.forEach((section, index) => {
    const x = startX + width * ((section.start - 1) / Math.max(1, navigation.total - 1));
    const isActive = index === active;
    const isCompleted = active >= 0 && index < active;
    const color = isActive ? colors.active : isCompleted ? colors.completed : colors.track;
    addShape(slide, 'ellipse', { x: x - 0.09, y: y - 0.1, w: 0.18, h: 0.18, fill: { color }, line: { color } });
    if (isCompleted) {
      addText(slide, '✓', { x: x - 0.055, y: y - 0.075, w: 0.11, h: 0.09, fontFace: theme.fontFace || 'Microsoft YaHei', fontSize: 5.6, bold: true, color: colors.onActive, align: 'center' });
    }
    addText(slide, section.title, {
      x: x - 0.5, y: labelY, w: 1, h: 0.13, fontFace: theme.fontFace || 'Microsoft YaHei',
      fontSize: 7.4, bold: isActive, color: isActive ? colors.active : inactiveLabel, align: 'center'
    });
  });
  const progress = (navigation.page - 1) / Math.max(1, navigation.total - 1);
  addShape(slide, 'line', { x: startX, y, w: width * progress, h: 0, line: { color: colors.active, pt: 2 } });
  addShape(slide, 'ellipse', { x: startX + width * progress - 0.11, y: y - 0.12, w: 0.22, h: 0.22, fill: { color: colors.active }, line: { color: colors.onActive, pt: 0.5 } });
}

function addNavigation(slide, input, theme) {
  const navigation = normalizeNavigation(input);
  if (!slide?.addShape || !slide?.addText) throw new Error('slide must support addShape and addText.');
  if (!theme?.colors?.nav) throw new Error('theme must provide semantic navigation colors.');
  if (navigation.style.startsWith('number')) addNumberNavigation(slide, navigation, theme);
  else addRailNavigation(slide, navigation, theme);
  return navigation;
}

module.exports = { addNavigation, normalizeNavigation, VALID_STYLES };
