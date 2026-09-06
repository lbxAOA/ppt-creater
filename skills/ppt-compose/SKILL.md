---
name: ppt-compose
description: Use when composing and QA-verifying a native PPTX deck.
version: 0.1.0
author: rdft1, Hermes Agent
license: MIT
platforms: [windows]
metadata:
  hermes:
    tags: [powerpoint, pptx, compose, qa, animation]
    related_skills: []
---

# PPT Compose Command

Execute a reviewed ppt-creater plan and return an editable native PowerPoint
deck. This is the production phase, not a mockup generator. The strict MCP
gate must complete preview, PowerPoint opening, normalization, animation
playback, and QA before success is reported.

## User Input

The text after `/ppt-compose` may specify the plan, output filename, or final
content corrections. Locate the latest plan when they are not supplied.

## Procedure

1. Read the selected `.ppt-creater/specs/<slug>/spec.md` and plan artifact.
2. Confirm the output filename ends in `.pptx` and is not an input source file.
3. Confirm any external family has explicit `allowExternal: true` approval from
   the user before proceeding.
4. Call `ppt_creater_powerpoint_com_status` and stop if the active COM backend
   is unavailable; do not silently fall back to a non-native export.
5. Call `ppt_creater_compose_deck` with the approved native composition request
   and explicit output name.
6. Treat the operation as successful only if the tool returns a successful QA
   status, an existing editable output path, zero unexpected new shapes, and
   valid playback/timeline evidence. If any gate fails, report `review_required`
   and do not claim completion.
7. Report output path, slide count, family/targets, object-fill result, preview
   and QA report paths, new-shape count, playback result, and timeline match.

## Invocation Examples

```text
/ppt-compose output_name=ai-industry-report.pptx
/ppt-compose 使用刚才的计划生成最终 PPTX，输出为 investor-brief.pptx
```

## Done When

The editable PPTX and returned QA reports exist, and every reported success
field is backed by the MCP result. Otherwise return the exact blocker and the
next corrective action.
