const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const test = require('node:test');
const { fillPlan } = require('../src/fill-template');

test('fillPlan copies a source template slide and replaces named slot text', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-creater-fill-'));
  const templatePath = path.join(root, 'template.pptx');
  const outputPath = path.join(root, 'filled.pptx');
  const makeTemplate = path.join(root, 'make_template.py');
  fs.writeFileSync(makeTemplate, [
    'from pptx import Presentation',
    'p=Presentation()',
    's=p.slides.add_slide(p.slide_layouts[6])',
    "shape=s.shapes.add_textbox(914400,914400,7315200,914400)",
    "shape.name='slot_title'",
    "shape.text='{{title}}'",
    `p.save(r'''${templatePath.replace(/\\/g, '\\\\')}''')`
  ].join('\n'));
  execFileSync('python', [makeTemplate]);

  await fillPlan({
    family: { source_template: templatePath },
    plan: { slides: [{ source_slide_index: 1, content: { title: '已替换标题' } }] },
    outputPath
  });
  assert.equal(fs.existsSync(outputPath), true);
  const inspect = `from pptx import Presentation\np=Presentation(r'''${outputPath.replace(/\\/g, '\\\\')}''')\nprint('\\n'.join(sh.text for s in p.slides for sh in s.shapes if getattr(sh,'has_text_frame',False)))`;
  const found = execFileSync('python', ['-c', inspect], { encoding: 'utf8' });
  assert.match(found, /已替换标题/);
});
