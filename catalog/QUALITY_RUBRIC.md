# Template Library Quality Rubric

Each candidate receives a 1–5 score after visual review.

| Criterion | What 5 means | Production consequence |
|---|---|---|
| Visual maturity | Contemporary, restrained, intentional system | eligible as a primary family |
| Information hierarchy | conclusion/title/body/data clearly ordered | usable for executive or research decks |
| Chinese readability | projection-safe text zones and high contrast | eligible for Chinese content generation |
| Data capability | charts, matrices, tables, metrics have coherent layouts | eligible for industry / board reporting |
| Editability | native editable shapes, not rasterized page art | eligible for automated filling |
| Motion quality | animation reveals logic; no decorative carnival | eligible for animation preservation |

## Target routing

- `academic-clean`: thesis, paper reading, research defense, education.
- `ai-industry`: AI/technology research, product strategy, deep-tech briefing.
- `investor-pitch`: startup fundraising, product story, venture update.
- `data-boardroom`: KPI, operating review, executive/consulting report.
- `marketing-event`: campaign, brand launch, event presentation.
- `visual-reference`: attractive but not safe for automatic content filling.

## Gates

A `production_ready` family needs: visual maturity ≥4, readability ≥4, editability ≥4; every registered reusable page must have named `slot_*` shapes. An animated page additionally needs one successful PowerPoint playback verification after content replacement.
