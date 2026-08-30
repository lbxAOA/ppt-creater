# PPT Template Library Import Report

Generated from `C:\PPT模板` for `C:\ppt-creater`.

## Processing policy

- Original template assets were indexed before curation.
- Selected production candidates are represented by **NTFS hard links** under `curated/template-families`; they do not duplicate the underlying template bytes.
- Legacy `.ppt` files were converted to `.pptx` copies under `legacy-converted`; original legacy files remain in the source library.
- Exact byte-for-byte duplicates were deleted from the source library only after retaining the first indexed copy in each duplicate group.
- No subjective/visual-quality deletion was performed. Scores are a routing aid, not a defensible replacement for a visual review.

## Catalogs

| Artifact | Purpose |
|---|---|
| `catalog/library-manifest.jsonl` | Original all-file manifest |
| `catalog/templates-enriched.jsonl` | Original PPTX structural + text/theme metadata |
| `catalog/legacy-conversion.jsonl` | Legacy PPT conversion audit trail |
| `catalog/legacy-templates-enriched.jsonl` | Converted legacy deck metadata |
| `catalog/all-templates-enriched.jsonl` | Unified searchable catalog |
| `catalog/families.json` | Candidate template-family registry |
| `curated/selection-manifest.json` | Curated template selection manifest |

## What the automatic review reads

For each readable PPTX, the system extracts:

- Page count, layouts, page size, shapes, native charts and tables;
- Theme color and font tokens;
- Text excerpts and page-pattern signals: agenda, chart, table, closing, minimal, content, dense composition;
- PowerPoint animation timing and transition XML presence;
- Exact SHA-256 duplicate groups.

## Important limitation

The library contains no pre-named `slot_*` shapes. Before a candidate becomes an automated production family, open it in PowerPoint and name the reusable target shapes through the Selection Pane:

```text
slot_title
slot_subtitle
slot_insight
slot_source
slot_chart_caption
asset_logo
anim_feedback_arrow
```

For animated pages, every filled `slot_*` must be marked `animation_locked: true` in the family registry. The template copier replaces only existing named shapes, preserving timing references; it never deletes/rebuilds animated shapes.

## Candidate selection

The automatic registry ranks candidates by editable PPTX status, page range, layout density, native chart/table coverage, animation/transition XML, and source-category coverage. It is intentionally marked `needs_slot_annotation`; it has not claimed that any candidate is visually approved.

## Next review loop

1. Review the 40 hard-linked candidates in `curated/template-families`.
2. Select a first production family for AI/technology research.
3. Name the reusable shapes in PowerPoint.
4. Add registered slide definitions and animation locks to `catalog/families.json`.
5. Generate a test deck with `src/fill-template.js`.
6. Play it in PowerPoint and verify animation, then mark the family `production_ready`.
