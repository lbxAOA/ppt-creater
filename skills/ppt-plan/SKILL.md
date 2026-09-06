---
name: ppt-plan
description: Use when planning a native ppt-creater deck from a brief.
version: 0.1.0
author: rdft1, Hermes Agent
license: MIT
platforms: [windows]
metadata:
  hermes:
    tags: [powerpoint, pptx, planning, templates, native]
    related_skills: []
---

# PPT Plan Command

Turn a presentation brief into a validated native-template deck plan. This is
the design phase: it selects one coherent family, maps pages and native slots,
and makes content and QA constraints explicit without exporting the final file.

## User Input

The text after `/ppt-plan` is an optional refinement. First locate the latest
`.ppt-creater/specs/*/spec.md`; if none exists, use the supplied request as a
brief and state that no durable spec was found.

## Procedure

1. Call `ppt_creater_status`. Stop if the relevant catalogs are unavailable.
2. Search or choose a family with `ppt_creater_search_templates` or
   `ppt_creater_choose_family`. Select exactly one coherent family; do not mix
   unrelated template families.
3. Call `ppt_creater_list_typography_themes` when a typography theme is
   requested. Preserve native typography by default.
4. Call `ppt_creater_list_composition_pages` for the selected target and map
   each planned slide to a listed page index and registered slot/object kind.
5. Call `ppt_creater_plan_deck` for a standard deck, or
   `ppt_creater_plan_composition` when native pages and object fills are needed.
6. Save the resulting plan as `.ppt-creater/specs/<slug>/plan.md` or
   `.plan.json` when a spec directory exists. Include family, pages, slot
   mapping, content limits, animation locks, color/typography decisions,
   external-license gates, and QA acceptance criteria.
7. Do not call `ppt_creater_compose_deck` in this phase.

## Invocation Examples

```text
/ppt-plan
/ppt-plan 采用 tech-precision，增加一页数据图表页
```

## Done When

The plan names one family, uses only registered native pages/slots, records all
assumptions and gates, and is ready for `/ppt-compose`.
