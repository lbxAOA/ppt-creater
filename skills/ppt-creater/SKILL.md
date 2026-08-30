---
name: ppt-creater
description: Create consistent PPTX decks from a local template library.
version: 0.1.0
author: User, Hermes Agent
license: MIT
platforms: [windows]
metadata:
  hermes:
    tags: [pptx, templates, animation, design-system, mcp]
    related_skills: [ppt-master, powerpoint]
---

# PPT Creater

Use the local `ppt-creater` project to create editable presentations from a curated PPTX template library. It preserves existing animation bindings by replacing named `slot_*` shapes in copied template pages; it does not regenerate animated pages from scratch.

## When to Use

- The user asks for a PPTX based on their local template library.
- The request needs a consistent template family, native editable slides, or preserved PowerPoint animation.
- Do not use for a one-off visual mockup when editable PowerPoint output is unnecessary.

## Prerequisites

- Configure `C:/ppt-creater/config/ppt-creater.config.json` to point at the project-contained curated template root; only set a different `library_root` for a separate new-library ingestion run.
- Run the catalog indexer before searching a newly supplied external template library.
- Curate a family registry at `C:/ppt-creater/catalog/families.json`; the existing project includes candidate and production registries.
- For native animation playback validation, use Windows desktop PowerPoint. LibreOffice rendering is static-only.

## Procedure

1. **Catalog non-destructively.** Run `terminal(command="python C:/ppt-creater/src/ppt_creater/index_templates.py <LIBRARY_ROOT> --output C:/ppt-creater/catalog/templates.jsonl", workdir="C:/ppt-creater")`. Confirm every scanned template has one JSONL record; do not modify source files.
2. **Curate a template family.** Keep cover, section, content, chart, comparison, and closing pages from one visual family. Name editable shapes `slot_*`; preserve `asset_*` and `anim_*` shapes unchanged.
3. **Register motion constraints.** For each animated page, record `source_slide_index`, stable shape names, required slots, character caps, and its existing `animation_pattern`. Mark every slot on an animated sequence `animation_locked: true`.
4. **Plan before filling.** Call MCP `ppt_creater_choose_family`, then `ppt_creater_plan_deck`. Use one `family_id` per deck. The planner must reject unsupported slide types, empty slots, and motion patterns absent from the family policy.
5. **Fill safely.** Use `pptx-automizer` or PowerPoint COM to copy the registered source page and update text or picture content inside existing `slot_*` shapes. Do not delete/recreate animation targets.
6. **Verify.** Render static pages for overflow and alignment. Open the resulting deck in PowerPoint, verify the `MainSequence` targets and slideshow playback for every animated page, then export a preview video if requested.

## Motion and Style Rules

- One template family per deck; do not mix unrelated templates.
- Only use colors, fonts, spacing, and page types from the family registry.
- Do not add decorative motion. Reuse a family’s approved entrance patterns and timing.
- Keep body text at or above 16 pt; use a second slide rather than shrinking text below the family minimum.
- Keep evidence/source lines for research, technical, and investment material.

## Verification

- `npm test` in `C:/ppt-creater` passes.
- The catalog marks a selected template `has_animation_xml: true` when animation preservation is required.
- Every planned page references a registered family slide and supplies every required slot.
- A PowerPoint playback check confirms animation survives after content replacement.
