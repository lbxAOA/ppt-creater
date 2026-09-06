---
name: ppt-search
description: Use when searching approved ppt-creater template families.
version: 0.1.0
author: rdft1, Hermes Agent
license: MIT
platforms: [windows]
metadata:
  hermes:
    tags: [powerpoint, pptx, templates, search]
    related_skills: []
---

# PPT Search Command

Search the approved, non-destructive ppt-creater template catalog and return
families that fit the user's presentation purpose. Never select a template only
from a filename or generic visual label.

## User Input

The text after `/ppt-search` is the search request. Extract the use case,
visual requirements, animation requirement, slot requirement, and minimum slide
count when provided.

## Procedure

1. Call `ppt_creater_status` and stop if the catalog is not ready.
2. Call `ppt_creater_search_templates` with the relevant filters.
3. Present the strongest candidates with family ID, route, slide count,
   animation/slot evidence, licensing gate, and why each fits.
4. Keep `external-*` families visibly marked as requiring external-license
   approval. Do not imply that technical validation grants redistribution rights.
5. Do not compose a deck. Offer `/ppt-plan` as the next step.

## Invocation Examples

```text
/ppt-search AI 行业分析，要求保留动画并支持数据图表
/ppt-search investor pitch require_animation=true
```

## Done When

Return a deduplicated candidate list grounded in the catalog and clearly state
which candidate is recommended and what must be verified next.
