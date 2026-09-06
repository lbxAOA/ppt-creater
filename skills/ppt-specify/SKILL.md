---
name: ppt-specify
description: Use when defining a structured PPT presentation brief.
version: 0.1.0
author: rdft1, Hermes Agent
license: MIT
platforms: [windows]
metadata:
  hermes:
    tags: [powerpoint, pptx, specification, requirements]
    related_skills: []
---

# PPT Specify Command

Create a source-grounded presentation specification before template selection or
composition. This is the ppt-creater equivalent of a Spec Kit specify phase:
define what the deck must accomplish, not how it will be implemented.

## User Input

The text after `/ppt-specify` is the presentation request. Do not ask the user
to repeat it. Infer reasonable defaults and ask only one high-impact
clarification at a time when scope, audience, or output expectations cannot be
resolved safely.

## Procedure

1. Extract audience, decision or learning goal, topic, language, approximate
   slide count, required sections, sources, deadline, visual constraints,
   animation requirements, and output expectations.
2. Produce a concise brief containing: objective, audience, narrative, slide
   requirements, evidence rules, template constraints, acceptance criteria,
   assumptions, and unresolved questions.
3. If the project needs a durable artifact, write it under
   `.ppt-creater/specs/<slug>/spec.md` and create the directory first. Do not
   overwrite an existing specification without identifying it to the user.
4. Do not choose a template or generate a PPTX during this phase.
5. End with the next command: `/ppt-plan`.

## Invocation Examples

```text
/ppt-specify 面向投资人的 AI 行业分析报告，中文，12 页
/ppt-specify 为研究组制作一份可保留原生动画的技术汇报
```

## Done When

The specification has a bounded scope, testable acceptance criteria, explicit
assumptions, and at most three unresolved questions. Report the spec path when
one was written.
