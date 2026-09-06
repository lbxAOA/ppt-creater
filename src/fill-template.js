const fs = require('node:fs');
const path = require('node:path');
const { Automizer, modify } = require('pptx-automizer');

const { createTypographyModifier } = require('./typography-system');

function typographyModifierFor(theme, slotName) {
  return theme ? createTypographyModifier(theme, slotName) : null;
}

function ensureSafePlan(family, plan) {
  if (!family?.source_template) throw new Error('Template family must define source_template.');
  if (!fs.existsSync(family.source_template)) throw new Error(`Source template does not exist: ${family.source_template}`);
  for (const page of plan.slides || []) {
    const definition = (family.slides || []).find((slide) => slide.slide_id === page.template_slide_id)
      || (family.slides || []).find((slide) => slide.source_slide_index === page.source_slide_index);
    if (definition?.animation_pattern && definition.animation_pattern !== 'none') {
      for (const slot of definition.slots || []) {
        if (!slot.animation_locked) throw new Error(`Refusing to fill animated slot without lock: ${slot.name}`);
      }
    }
  }
}

async function fillPlan({ family, plan, outputPath }) {
  ensureSafePlan(family, plan);
  const templatePath = path.resolve(family.source_template);
  const outputDir = path.dirname(path.resolve(outputPath));
  const templateDir = path.dirname(templatePath);
  const templateName = path.basename(templatePath);
  fs.mkdirSync(outputDir, { recursive: true });

  const automizer = new Automizer({ templateDir, outputDir, removeExistingSlides: true });
  const pres = automizer.loadRoot(templateName).load(templateName, 'family');
  for (const page of plan.slides || []) {
    const sourceSlide = page.source_slide_index;
    if (!Number.isInteger(sourceSlide) || sourceSlide < 1) throw new Error(`Invalid source_slide_index: ${sourceSlide}`);
    pres.addSlide('family', sourceSlide, (slide) => {
      for (const [slot, value] of Object.entries(page.content || {})) {
        // Selecting the existing named shape is intentional: it preserves the page's
        // structure and any timing reference pointing to that shape.
        const shapeName = `slot_${slot}`;
        const typography = typographyModifierFor(page.typography_theme || plan.typography_theme, slot);
        slide.modifyElement(shapeName, [
          modify.setText(String(value)),
          ...(typography ? [typography] : [])
        ]);
      }
    });
  }
  return pres.write(path.basename(outputPath));
}

module.exports = { fillPlan, ensureSafePlan };
