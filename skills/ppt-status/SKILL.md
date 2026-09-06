---
name: ppt-status
description: Use when checking ppt-creater readiness and backend status.
version: 0.1.0
author: rdft1, Hermes Agent
license: MIT
platforms: [windows]
metadata:
  hermes:
    tags: [powerpoint, pptx, status, diagnostics]
    related_skills: []
---

# PPT Status Command

Check whether the ppt-creater catalog, composition catalog, and PowerPoint COM
backend are ready. This command is diagnostic only; it does not create or
modify a presentation.

## Procedure

1. Call the ppt-creater status MCP tool.
2. Report the project root, catalog readiness, family count, composition
   catalog readiness, configured backend, and any missing prerequisites.
3. If the user asks whether desktop PowerPoint is actually reachable, also call
   the active PowerPoint COM status MCP tool. Distinguish configured state from
   an active probe.
4. Do not report generation readiness if a required catalog or QA backend is
   unavailable.

## Invocation

```text
/ppt-status
```

Use the discovered MCP names `ppt_creater_status` and
`ppt_creater_powerpoint_com_status` (Hermes may display an `mcp_ppt_creater_`
prefix).

## Done When

Return a concise readiness report with explicit `ready`, `blocked`, or
`partial` status and the reason for each blocked capability.
