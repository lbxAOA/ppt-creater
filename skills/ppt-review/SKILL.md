---
name: ppt-review
description: Use when reviewing a ppt-creater plan or generated deck.
version: 0.1.0
author: rdft1, Hermes Agent
license: MIT
platforms: [windows]
metadata:
  hermes:
    tags: [powerpoint, pptx, review, qa, templates]
    related_skills: []
---

# PPT Review Command

Audit a ppt-creater plan or compose result against native-template, content,
licensing, editability, and QA requirements. Review is evidence-based and must
separate verified facts from recommendations.

## User Input

The text after `/ppt-review` identifies a plan, PPTX, or output directory. If it
is omitted, locate the latest `.ppt-creater/specs/` plan and recent output.

## Procedure

1. Read the relevant spec, plan, and returned QA JSON/report paths.
2. Check one-family consistency, registered page/slot usage, content limits,
   animation-locked objects, native editability, external-license gates, and
   required source attribution.
3. Call `ppt_creater_status` and, when the review concerns PowerPoint playback,
   call `ppt_creater_powerpoint_com_status`.
4. Classify each finding as `verified`, `review_required`, or `not_checked`.
5. Do not repair or regenerate the deck unless the user explicitly asks; return
   a prioritized correction list and whether `/ppt-compose` can be rerun.

## Invocation Examples

```text
/ppt-review 检查刚生成的 AI 行业报告
/ppt-review .ppt-creater/specs/001-ai-report/plan.md
```

## Done When

Return a compact audit with evidence paths, blocking findings, non-blocking
findings, and a clear release recommendation.
