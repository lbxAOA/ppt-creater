# ppt-creater

Native PowerPoint template routing for Hermes.

## Production-ready template library

The canonical registry is `catalog/unified-production-template-library.json`.
It contains **14 validated template families** across all supported production
routes. Every routed family has an editable native master, a filled validation
copy, `new_shapes=0`, and PowerPoint slideshow verification in the local
workspace.

### Naming convention

- `local-*` — a user-owned, project-contained source template.
- `external-*` — a third-party template imported through the isolated intake
  lane. Its source `NET-XX` remains immutable provenance, not the family name.
- Family IDs are lowercase kebab-case: `<origin>-<route-or-style>-<descriptor>`.
- A generated deck uses exactly one family; route alternatives are never mixed.

### Rights gate

`local-*` families may be selected directly. `external-*` families retain
`manual_external_review_required`: before a deck uses one, the caller must pass
explicit external-license approval (`allowExternal: true`). This means the
native fill pipeline has been verified, **not** that third-party assets may be
redistributed or used commercially without checking the source terms.

The repository intentionally excludes `.ppt` and `.pptx` assets. It contains
code, registries, verification metadata, and reproducible ingestion scripts;
local and third-party PowerPoint binaries remain outside version control.

### Current routing coverage

| Production target | Default family | Alternative families |
|---|---|---|
| Academic / research | `local-modern-report` | — |
| AI / technology | `external-ai-tool-pitch` | `external-tech-product-workflow`, `external-investor-modern` |
| Investor pitch | `external-investor-modern` | `external-investor-strategy`, `external-ai-tool-pitch` |
| Data / boardroom | `external-quarterly-business-review` | `external-kpi-scorecard`, `external-financial-management-consulting` |
| Company profile | `external-company-professional` | `external-b2b-market-analysis` |
| Generic corporate | `external-company-professional` | `external-quarterly-business-review` |
| Marketing event | `external-event-culture-tourism` | — |
| Culture tourism | `external-event-culture-tourism` | — |
| Wedding / bridal | `local-wedding-album` | `external-wedding-amelia` |
| Career portfolio | `external-career-portfolio` | — |
| B2B solution | `external-b2b-market-analysis` | — |
| Technical product | `external-tech-product-workflow` | `external-ai-tool-pitch` |

The coverage audit found **no missing supported route**, so no additional
network intake is required now. Future intake should address an explicitly new
route or a missing page type, not merely duplicate an existing cover-and-bullets
family.

## Repository layout

- `template-library/local-production-sources/` — local immutable user-owned source copies (not versioned).
- `network-supplements/NET-*/` — isolated original external downloads and review assets (not versioned).
- `network-production-library/external-*/` — promoted external masters (not versioned).
- `catalog/unified-production-template-library.json` — canonical production routing.
- `catalog/network-production-template-library.json` — external intake validation evidence.

## Quick start

```bash
npm test
python -m pytest tests -q
```

To rerun the internal naming migration after an intentional intake update:

```bash
python src/ppt_creater/normalize_production_library.py
```

Do not edit a copied `source.pptx` in place. Rebuild a production master from
that single source family, then refill and verify it in Microsoft PowerPoint.
