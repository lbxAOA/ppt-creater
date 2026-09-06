# ppt-creater

Native PowerPoint template routing for Hermes.

## Public repository and template assets

This repository is public and includes only template binaries identified in the
local production-source manifest as user-owned. Third-party intake downloads,
network masters, review candidates, derived slide collections, and legacy
conversion collections remain excluded by `.gitignore`.
Validation copies and generated decks are also excluded because they are local
QA artifacts rather than reusable source assets.

The included PPTX files are distributed only under the rights the repository
owner has for those local sources. The code and metadata are MIT-licensed; a
PPTX file is not automatically covered by the code license. Users must verify
that their own use and redistribution rights cover each template asset.

For a fresh clone, install Git LFS before pulling the template binaries:

```bash
git lfs install
git lfs pull
```

The source templates are under `template-library/local-production-sources/`.
Production masters are under `production-library/` and selected native masters
under `production-template-library/` and `complete-production-library/`.

`src/powerpoint-com.js` is the built-in Windows adapter for the repository's native PowerPoint scripts. It validates a bounded script name, passes arguments without constructing shell code, applies a timeout, and reports an explicit availability result. It requires an interactive Windows session with desktop Microsoft PowerPoint; it does not bundle Office. By default `visible: false` starts PowerPoint as a real COM application in a minimized window (`Visible = -1`, `WindowState = 2`), because this PowerPoint installation rejects `Visible = 0`; set `visible: true` to restore the window.

Configuration in `config/ppt-creater.config.json`:

```json
{
  "backend": "auto",
  "powerpoint_com": {
    "visible": false,
    "timeout_ms": 120000,
    "display_alerts": false
  }
}
```

`backend` accepts `auto`, `native`, or `powerpoint-com`. `native` means portable OOXML/template filling and must not depend on Office. `powerpoint-com` enables desktop PowerPoint for native object edits and validation; the strict preview/open/playback QA pipeline always requires this backend. `auto` selects COM only after a successful `PowerPoint.Application` probe and otherwise keeps portable operations available. The normal MCP status call reports configured state without launching Office; `ppt_creater_powerpoint_com_status` performs the active probe. Forced COM returns a clear error if unavailable. `ppt_creater_compose_deck` is a production gate and intentionally does not fall back: it returns only after native copy/fill, master and output full-page previews, PowerPoint opening, normalization, and playback QA are verified; its result includes the report paths. The strict native report pipeline uses the same adapter.

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
| Marketing event | `local-marketing-event-chinese-style` | — |
| Culture tourism | `local-culture-tourism-chinese-style` | — |
| Wedding / bridal | `local-wedding-album` | `external-wedding-amelia` |
| Career portfolio | `external-career-portfolio` | — |
| B2B solution | `external-b2b-market-analysis` | — |
| Technical product | `external-tech-product-workflow` | `external-ai-tool-pitch` |

The coverage audit found **no missing supported route**, so no additional
network intake is required now. Future intake should address an explicitly new
route or a missing page type, not merely duplicate an existing cover-and-bullets
family.

## Repository layout

- `template-library/local-production-sources/` — user-owned source templates included in the public asset set.
- `production-library/` — selected user-owned native masters included in the public asset set.
- `network-supplements/NET-*/`, `network-production-library/external-*/`, `curated/`, `legacy-converted/`, and review/addition directories — local-only material excluded from public redistribution.
- `catalog/unified-production-template-library.json` — canonical production routing.
- `catalog/network-production-template-library.json` — external intake validation evidence.

## 内置字体主题

PPT-creater 提供可切换的排版主题，不只是单独替换字体。每套主题统一定义标题、 副标题、正文、章节标题、注释和数字的字体、字号、字重与行距，并为中文提供 `Noto Sans SC` / `Noto Serif SC` 回退。系统不打包字体文件：导出时优先使用目标机器已安装字体，缺失时按主题回退链显示。

当前内置主题：

| ID | 名称 | 适用场景 |
|---|---|---|
| `modern-business` | 现代商务 | 商业汇报、项目总结、公司介绍 |
| `tech-precision` | 极简科技 | AI、技术架构、数据产品 |
| `editorial-serif` | 人文杂志 | 品牌故事、内容展示 |
| `premium-finance` | 高端金融 | 投资人材料、董事会、财务报告 |
| `vibrant-creative` | 活力创意 | 营销、活动、年轻产品发布 |
| `education-clear` | 教育清晰 | 课程、培训、学术汇报 |

通过 MCP 查看主题：

```json
{}
```

调用 `ppt_creater_list_typography_themes` 获取主题清单；调用 `ppt_creater_plan_deck` 或 `ppt_creater_compose_deck` 时传入 `typography_theme_id`，例如 `tech-precision`。未传入时保持原生模板排版不变，保证旧项目兼容。

填充原生模板时，系统只修改已注册的 `slot_*` 文本对象，按槽位名称推断 `title`、`subtitle`、`heading`、`body`、`caption` 或 `number` 角色，不改变文本内容、位置或动画对象绑定。


所有可生成内容与导航栏现在通过 `src/color-system.js` 解析同一份语义主题，支持两种模式：

1. **模板模式（默认）**：从选中 template family 的 `theme.colors` 和首选字体抽取语义颜色；导航栏从该主题派生 `nav.active`、`nav.track`、`nav.completed` 等状态色。
2. **用户模式**：调用方提供完整 `background`、`text`、`muted`、`accent`、`line`、`darkBackground` 色板；同一色板同时驱动获准的模板颜色目标与导航栏。系统拒绝缺失 token 的部分覆盖，避免模板色与导航色混搭。

原生模板正文仍默认继承既有对象的原字体、字号、颜色、字重和段落样式。用户模式不能未经模板 capability manifest 白名单而重着色正文文本；首期可统一调整导航与已登记的非文本装饰/图表色。

`ppt_creater_plan_deck` 可接收：

```json
{
  "color_mode": "template",
  "color_theme_name": "optional-name"
}
```

或用户色板：

```json
{
  "color_mode": "user",
  "colors": {
    "background": "FFF7ED",
    "text": "431407",
    "muted": "9A3412",
    "accent": "EA580C",
    "line": "FED7AA",
    "darkBackground": "7C2D12"
  }
}
```

生成的 deck plan 会包含已解析的 `color_theme`，供填充器和导航器使用。旧的手工 AI 行业报告生成器也已改为通过该语义系统派生颜色，不再维护独立的基础色常量。

## Native-only generation and verification

Every public PPT-creater generation route is strict native-only: it selects one
production-ready template family, copies existing PowerPoint pages, fills
registered native objects, previews every page before and after filling, then
runs structural and PowerPoint COM QA. It does not redraw production pages with
PptxGenJS.

The standard AI report entry point is:

```bash
node src/generate-ai-industry-report.js
```

It writes the editable result to `output/2026-ai-industry-report-full-native.pptx`,
plus draft/final preview reports and QA artifacts under `output/previews/` and
`output/`.

`src/generate-navigation-style-gallery.js` is an internal diagnostic gallery,
not a public production-generation route; it may use PptxGenJS solely to
compare navigation tokens.


生成严格原生模板流程的可编辑 12 页 PPTX，并执行母版、draft、final 三阶段全页预览及 PowerPoint QA：

```bash
node src/generate-ai-industry-report.js
```

在 Windows + PowerPoint 上验证实际打开：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\verify-powerpoint-open.ps1 `
  -InputPath .\output\2026-ai-industry-report-full-native.pptx
```

运行后应得到：

- `output/2026-ai-industry-report-full-native.pptx`
- `output/2026-ai-industry-report-full-native.plan.json`
- `output/2026-ai-industry-report-full-native.qa.json`
- `output/2026-ai-industry-report-full-native.powerpoint-qa.json`
- `output/previews/2026-ai-industry-report-full-native/{master,draft,final}/slide-001.png ... slide-012.png`


To rerun the internal naming migration after an intentional intake update:

```bash
python src/ppt_creater/normalize_production_library.py
```

Do not edit a copied `source.pptx` in place. Rebuild a production master from
that single source family, then refill and verify it in Microsoft PowerPoint.
