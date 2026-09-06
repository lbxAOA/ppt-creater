# PPT-Creater 颜色风格切换与动画编辑实施方案

> **For Hermes:** Use subagent-driven-development skill to implement this plan task-by-task.

**Goal:** 在不破坏原生模板版式、文本样式或动画引用的前提下，为 `ppt-creater` 增加受控的颜色风格切换与**动画母版页复制**能力，并通过 MCP 公开为可验证工具。

**Architecture:** 将“风格”与“动画”建模为模板族的显式能力清单，而不是任意修改 PPTX。规划器只接受登记过的 palette / motion donor variant，并在生成计划阶段完成兼容性校验；执行器复制完整且已经设计、验证过的动画 donor 页，并仅替换这些页中已有的 `slot_*` 对象。每次写入都产出审计 JSON，并在 PowerPoint 幻灯片放映中验证。

**Tech Stack:** Node.js CommonJS、`@modelcontextprotocol/sdk`、Zod、Node test runner、`pptx-automizer`（既有文本槽位填充）、Windows PowerPoint COM / PowerShell（原生颜色与动画编辑和回放验证）。

---

## 当前基础与约束

- MCP 入口为 `src/mcp-server.js`，当前仅提供状态、搜索、模板族选择与 deck plan 工具。
- `src/planner.js` 已限制每个 deck 只能选择一个 family，并且对动画模式执行 allow-list 校验。
- `src/fill-template.js` 只替换既有 `slot_*` 文本对象，且对动画页要求 `animation_locked`。
- 已有的 `src/create-investor-animation-variant.ps1` 证明 PowerPoint COM 可以只对既有形状的 `TimeLine.MainSequence` 重建淡入序列，并进行一次幻灯片放映验证。
- 生产注册表 `catalog/unified-production-template-library.json` 当前规定 `strict_native_only: true`、`new_shapes: 0`；新能力必须保留这些性质。
- “颜色切换”不得通过覆盖文字字体颜色、手工创建新文本框或替换整页来实现。默认只允许修改经模板标注的既有、非动画锁定的形状色彩；若模板有完整的 theme / variant，优先使用模板原生变体。
- “动画编辑”不得静默改写原始动画。任何改动必须基于模板登记的受支持 motion preset，并清楚声明是保留、替换还是移除原动画。

---

## 对外能力边界

### 1. 颜色风格切换

支持三层，从安全性最高到最低依次尝试：

1. **原生主题/变体层**：应用模板预登记的 PowerPoint theme variant 或已验证的 theme color 映射；维持原对象的主题色绑定。
2. **语义 palette 层**：模板为既有形状登记 `semantic_color_slots`（例如 `accent_primary`、`accent_secondary`、`surface`、`chart_series_1`），只按映射改写这些对象的填充、线条或图表系列色。
3. **不可用则拒绝**：若目标 family 未登记某 palette，或修改对象包含未锁定/未审批的文本字体色、SmartArt、图片、渐变、图表或动画绑定形状，返回明确错误，不退化成“猜测颜色”。

首期范围：纯色填充、纯色线条、图表系列的已标注颜色；文字颜色仅允许对象明确登记 `allow_text_color_change: true` 时修改。渐变、图片调色、艺术字、SmartArt 与未解析的主题引用列为后续能力。

### 2. 动画设计：以“动画母版页复制”为主，而非重建

应将一个经过人工设计和播放验证的 PPTX 作为**动画母版（motion donor）**。生成新 PPT 时，复制其中完整的动画页，再只替换这些页中既有 `slot_*` 对象的内容。这样可以原样保留 PowerPoint 内部的时间线、并行/延迟关系、触发器、路径动画、音频、切换和对象 ID 引用，而不是用 COM 重新猜测动画。

**首选工作流：**

1. 为每个模板族建立不可变的 `motion-master.pptx`，其中每种可复用页面类型各有一张或多张已设计好的动画页。
2. 每页的可替换文本、图表数据或图片占位对象必须使用稳定的 `slot_*` 名称；动画所引用的形状必须就是这些页面中的原有形状。
3. 生成时通过复制完整 donor slide 到新的输出 deck，再填充 slot，而不是从一张普通静态页向另一张页面“移植”动画。
4. 一个新 deck 仍只用一个 family；可以选用同一 family 中不同 `motion_variant_id` 的 donor 页，但不能把不同 family 的动画页混用。
5. 填充后必须比较动画 XML 与对象关系，并在 PowerPoint 中实际放映，确认点击次数、显现顺序、路径和触发器仍按 donor 工作。

**为什么不将动画直接复制到任意新页面：**动画效果引用 slide 内形状的内部关系和顺序，复杂动画还可能依赖 timing node、触发器、媒体关系、路径、组对象及页面切换。即使视觉上两页结构相似，跨页拷贝时间线也容易失效。因此，“复制完整设计页 + 原位替换内容”是生产级可靠路径；对于没有对应 donor 页的页面类型，应拒绝动画生成，或回退为同 family 的静态原生页，绝不虚构动画。

COM 方式保留为**受限补充**：只用于修复已经登记、低复杂度的页内动画（例如对原有对象应用一个已验证的淡入序列），不能替代 donor-copy 路线，也不用于路径动画、复杂触发器、音视频同步或动画窗格中的组合效果。

首期 motion 语义改为：

- `copy_verified_donor`：复制完整动画 donor 页并填充 slot；这是默认且优先的方式。
- `preserve_original`：复制原生页面且不触碰时间线；适用于 donor 已带动画但不需改节奏。
- `static_native`：复制同 family 静态原生页，不添加或重建动画。
- `com_fade_sequence`：仅已人工验证的低复杂度例外，使用 COM 在原有形状上重建淡入序列。

调用方只能选择 family 已登记的 `motion_variant_id` / `donor_slide_id`，不能提交任意 shape index、任意效果常量或任意外部 PPTX 路径。

---

## 建议的数据模型

在 family 或 production template 元数据中增加版本化的 `style_capabilities` 与 `motion_capabilities`；不改变现有 `theme` 字段的兼容语义。

```json
{
  "capabilities_version": 1,
  "style_capabilities": {
    "default_palette_id": "native",
    "palettes": [
      {
        "palette_id": "native",
        "label": "模板原生",
        "kind": "native",
        "verified": true,
        "mapping": {}
      },
      {
        "palette_id": "teal-night",
        "label": "深青科技",
        "kind": "semantic",
        "verified": true,
        "mapping": {
          "accent_primary": "18B6B2",
          "accent_secondary": "315476",
          "surface": "F6FAFB",
          "chart_series_1": "18B6B2"
        }
      }
    ],
    "semantic_color_slots": [
      {
        "slide_id": "architecture",
        "shape_name": "slot_accent_bar",
        "property": "fill",
        "token": "accent_primary",
        "animation_locked": false
      }
    ]
  },
  "motion_capabilities": {
    "default_variant_id": "native-original",
    "donor_deck_path": "C:\\ppt-creater\\production-library\\local-modern-report\\motion-master.pptx",
    "variants": [
      {
        "motion_variant_id": "architecture-sequence-v1",
        "mode": "copy_verified_donor",
        "donor_slide_id": "architecture-sequence-v1",
        "source_slide_index": 6,
        "slot_names": ["title", "step_1", "step_2", "step_3"],
        "verified": true,
        "playback_evidence_path": "catalog/qa/local-modern-report/architecture-sequence-v1.json"
      }
    ],
    "slide_permissions": [
      {
        "slide_id": "architecture",
        "allowed_variant_ids": ["native-original", "architecture-sequence-v1"],
        "requires_playback_verification": true
      }
    ]
  }
}
```

生产 registry 应只保存已验证的能力摘要、母版路径、验证证据路径和版本；细粒度 slide/shape 标注放在一个新的 family capability manifest 中，例如 `catalog/capabilities/<family_id>.json`，避免让 15KB+ 生产路由注册表承担人工编辑的形状清单。

---

## 实施任务

### Task 1：定义能力 manifest、加载与验证模块

**Objective:** 引入稳定且可审计的 style/motion capability schema，并拒绝不安全或自相矛盾的登记。

**Files:**
- Create: `src/capabilities.js`
- Create: `catalog/capabilities/example-tech-navy.json`
- Modify: `src/catalog.js:32-62`
- Test: `tests/capabilities.test.js`

**Step 1: Write failing tests**

覆盖：
- 缺少 `palette_id`、重复 palette id、非法 RGB（非 6 位十六进制）会失败；
- `semantic_color_slots` 的 `property` 不是 `fill` / `line` / `text` / `chart_series` 会失败；
- `text` 属性未显式 `allow_text_color_change: true` 会失败；
- motion preset 引用了未登记的 shape 名称、或 slide 允许不存在的 preset，会失败；
- `animation_locked: true` 的 shape 不能进入颜色修改目标；
- 合法 manifest 能被加载，并保留版本与验证状态。

**Step 2: Run the new test and verify it fails**

Run: `npm test -- --test-name-pattern="capabilit"`

Expected: FAIL，因为模块尚不存在。

**Step 3: Implement `src/capabilities.js`**

导出：
- `loadCapabilities(filePath)`：读取 UTF-8 JSON、去 BOM、返回对象；
- `validateCapabilities(manifest)`：执行上述静态校验；
- `getPalette(manifest, paletteId)`：默认选择 `default_palette_id`；
- `getSlideMotionPermission(manifest, slideId, presetId)`：验证 slide 和 preset 兼容；
- `validateColorEditTargets(manifest, paletteId, slideIds?)`：只返回已验证且未动画锁定的目标。

不要把 COM 细节、文件写入或 PowerPoint 路径放进本模块。

**Step 4: Run targeted tests**

Run: `npm test -- --test-name-pattern="capabilit"`

Expected: PASS。

**Step 5: Run complete unit suite**

Run: `npm test`

Expected: existing tests and new capability tests all pass.

---

### Task 2：扩展 deck plan，选择 donor 动画变体但不执行写入

**Objective:** 让 plan 在生成阶段选择 palette 和同一 family 内已验证的 motion donor variant，并在写入前完成兼容性校验。

**Files:**
- Modify: `src/planner.js:25-65`
- Modify: `src/mcp-server.js:60-74`
- Modify: `tests/planner.test.js`
- Test: `tests/planner.test.js`

**Step 1: Write failing tests**

新增以下情形：
- request `palette_id: 'teal-night'` 生成的 plan 顶层含 `palette_id` 和 capability manifest version；
- 未登记 palette 被拒绝；
- 页面请求已允许 `motion_variant_id: 'architecture-sequence-v1'` 时，plan 包含 donor deck、source slide index 与 slot contract；
- 调用方选择的 donor variant 不属于该页或 family 时被拒绝；
- 调用方未指定时继承 `native-original` / `preserve_original`；
- family 不存在 capabilities 时，只允许默认 `native` + 原始页面复制，其他切换请求明确失败；
- 页面内容缺少 donor 的一个必填 `slot_*` 时被拒绝。

**Step 2: Run targeted test and verify failure**

Run: `npm test -- --test-name-pattern="buildDeckPlan|chooseFamily"`

Expected: FAIL。

**Step 3: Implement planner changes**

- `buildDeckPlan` 接收一个已加载 capabilities manifest（可通过 `family.capabilities_path` 派生加载，避免 MCP 调用端传任意文件路径）。
- 顶层加入：`plan_version: 2`、`palette_id`、`capabilities_version`、`style_change_requested`、`motion_change_requested`。
- 每页加入：`motion_variant_id`、`motion_mode`、`donor_deck_path`、`donor_source_slide_index`、`required_slots`。
- 保留现有 `animation_pattern` 字段一版以兼容现有 plan；将它映射为 `preserve_original` 的语义，不与 donor variant 冲突。
- 在 `mcp-server.js` 的 `ppt_creater_plan_deck` schema 中加入可选 `palette_id` 与 `slides[].motion_variant_id`，保留旧 `animation` 字段并标记为过渡输入。

**Step 4: Run tests**

Run: `npm test`

Expected: PASS。

---

### Task 3：制作登记与扫描工具，建立稳定的形状身份

**Objective:** 不再依赖易变的 shape index，让每个可编辑对象有可审核的名字、属性和动画锁定状态。

**Files:**
- Create: `src/inspect-template-capabilities.ps1`
- Create: `src/register-template-capabilities.ps1`
- Create: `docs/template-capability-authoring.md`
- Modify: `README.md:55-77`
- Test: manual fixture + JSON validation from `tests/capabilities.test.js`

**Step 1: Define required manual authoring workflow**

文档明确：在 PowerPoint 中只为批准的原生形状命名，例如 `pptc_color_accent_01`、`pptc_motion_card_01`；不要修改源模板，先复制为 family master。每个命名映射须记录 source slide、shape name、shape index、原有填充/线色、文字是否许可、动画锁定、预期 preset。

**Step 2: Implement inspection script**

`inspect-template-capabilities.ps1` 使用 PowerPoint COM 输出每页：
- slide index；
- shape index、name、type、是否有文本、填充/线条是否可读取；
- 所有 MainSequence effect 的 shape target、effect 类型、trigger；
- 是否发现重复或空 shape 名。

输出 JSON，不修改输入 PPTX。

**Step 3: Implement registration script**

`register-template-capabilities.ps1` 接收已审核的 mapping JSON，调用 `validateCapabilities`，将 capability manifest 写入 `catalog/capabilities/<family>.json`。脚本不得自动“猜测”语义 token、形状归属或动画组。

**Step 4: Verify against one existing local production master**

对 `local-modern-report` 或另一个用户拥有且可访问的 family master 跑 inspection，人工选择一个纯色装饰形状和一个纯色线条作为样本，生成 capability manifest；不改动 master。

Expected: JSON 中所有颜色目标均有稳定 `shape_name`，且 motion target 可映射到既有形状。

---

### Task 4：实现原生颜色应用执行器

**Objective:** 在既有、已批准形状上应用语义 palette，并生成变更审计，不改变形状数量、位置、尺寸或文本样式。

**Files:**
- Create: `src/apply-native-palette.ps1`
- Create: `src/style-executor.js`
- Modify: `src/fill-template.js:19-43`
- Test: `tests/style-executor.test.js`

**Step 1: Write tests for execution request construction and audit validation**

Node 测试不要求本机 PowerPoint；覆盖：
- executor 只接受 v2 plan、已验证 palette 与 matching capability manifest；
- `palette_id: native` 是 no-op，不调用 PowerShell；
- 颜色编辑目标若为动画锁定或未验证，执行前拒绝；
- audit JSON 若 `new_shapes !== 0`、`position_changes !== 0`、`text_style_changes !== 0` 或 palette id 不匹配，判定失败；
- 输出路径必须位于调用方指定的 output root。

**Step 2: Implement `style-executor.js`**

- 从 plan/family 推导 master 与 capability manifest，不接受任意 PPTX 输入路径；
- 复制 master 到目标输出；
- 以非交互 `powershell.exe -File` 启动执行脚本；
- 将 palette、slots 和审计输出路径经临时 JSON 文件传入，不拼接为 shell 字符串；
- 读取审计 JSON，执行 Node 侧二次校验。

**Step 3: Implement `apply-native-palette.ps1`**

- 打开复制后的 PPTX；
- 对每一个清单目标根据 `shape_name` 定位 shape，若找不到或找到了多个则停止；
- 仅允许：Shape.Fill.ForeColor.RGB、Shape.Line.ForeColor.RGB，和首期明确支持的图表 series fill；
- `text` 仅在 manifest 明确允许时处理，并在 audit 中记为 `text_style_changes`；首期示范 capability 不应包含 text；
- 记录处理前后颜色、slide/shape identity；
- 保存、重新打开并记录 slide count、shape count、每个 shape 的位置尺寸哈希；
- 输出 `{new_shapes: 0, position_changes: 0, text_style_changes, palette_id, applied_targets, skipped_targets, verification_status}`。

**Step 4: Verify with the sample native master**

- 先复制并对 `native` 跑 no-op；审计应显示 zero changes。
- 对样本 semantic palette 跑一次；人工打开输出 PPTX，检查原生文本字体、字号、颜色和布局不变，只有登记的装饰对象改色。
- 用 COM 对比输入/输出的每页形状数和 geometry hash。

**Step 5: Run regression tests**

Run: `npm test`

Expected: PASS。

---

### Task 5：实现 donor 动画页复制、填充与播放验证

**Objective:** 从同一模板族的已验证 `motion-master.pptx` 复制完整设计页，原位填充 slot，并证明时间线与复杂动画关系得到保留。

**Files:**
- Create: `src/motion-donor-executor.js`
- Create: `src/inspect-motion-donor.ps1`
- Create: `src/verify-motion-donor-output.ps1`
- Modify: `src/fill-template.js:19-43`
- Test: `tests/motion-donor-executor.test.js`

**Step 1: Write failing tests**

覆盖：
- `copy_verified_donor` 只接受 production registry 关联的 donor deck，而不接受调用方指定任意 PPTX 路径；
- 变体所选 `source_slide_index`、`donor_slide_id` 和 `slot_names` 与 capability manifest 不符时拒绝；
- 输出 deck 页来自 donor 的完整 slide 复制，随后仅调用现有 `slot_*` 文本替换；
- 缺少必填 slot 或填充试图触及未登记对象时拒绝；
- donor 与输出的动画结构签名（per-slide effect count、effect target name/order、timing XML canonical hash）不一致时失败；
- 放映验证失败时输出状态为 `review_required`，MCP 不得报告成功；
- 没有 donor 的页面只能走同 family `static_native`，不能调用 COM 猜测复杂动画。

**Step 2: Implement donor inspection and registration evidence**

`inspect-motion-donor.ps1` 对 `motion-master.pptx` 输出每个候选 donor 页的：
- source slide index、slide ID / 类型、全部 `slot_*` 形状名称；
- MainSequence effect count、每个 effect 的 target shape name、effect 类型、trigger、delay / duration；
- slide transition；
- timeline XML 与相关 relationship 的规范化 SHA-256；
- 是否含路径动画、触发器、媒体、音频或视频关系。

人工审核后将上述签名写入 capability manifest；脚本只读取，不修改 donor。

**Step 3: Implement `motion-donor-executor.js`**

- 从 v2 plan 和 family capability manifest 推导唯一 donor deck / donor source slide；不接收任意路径。
- 使用支持跨 deck 复制原页面包及关系的机制复制 donor slide；优先扩展 `pptx-automizer` 的既有 `pres.addSlide('family', sourceSlide, ...)` 工作流，并以 donor deck 作为 `family` source，而非通过 COM 在新静态页重建效果。
- 复制完成后只调用既有 `modify.setText` 替换登记的 `slot_*`；不得新建形状、删除形状或改动页面尺寸。
- 在每页填充后调用 PowerShell verifier；将审计 JSON 附在 output metadata 旁。

**Step 4: Implement `verify-motion-donor-output.ps1`**

- 重新打开 donor 与输出，按 stable `shape_name` 比较动画 effect target、effect type、trigger、order、duration 和 delay；
- 比较 donor / 输出的 shape count、形状名称集合、geometry hash；允许已批准 slot 的 text content 变化，但文本样式（字体、字号、颜色、粗细、段落样式）必须相同；
- 运行 `SlideShowSettings.Run()`，对每张 donor-copy 页按 donor effect 数逐步点击并记录是否能完整结束；路径、触发器或媒体存在时额外记录对象关系未丢失；
- 输出 `new_shapes`、`deleted_shapes`、`geometry_changes`、`timeline_signature_match`、`slot_style_match`、`powerpoint_playback`、`status`。

**Step 5: Verify one controlled real animation master**

建立一份用户拥有、结构清晰的 `motion-master.pptx`，至少包含封面、流程图、卡片序列与结尾四种 donor 页面：
1. 在 PowerPoint 中人工设计并确认动画；
2. 为每个可替换对象命名 `slot_*`；
3. 生成 capability manifest 和 donor 签名；
4. 以新内容生成 PPTX；
5. 在 PowerPoint 中与 donor 对照放映，确认点击次数、显现顺序、路径 / 同时动画关系完全相同；
6. audit 必须显示 `new_shapes: 0`、geometry unchanged、timeline signature match、slot style match 和 playback true。

**Step 6: Keep COM reconstruction as an explicit later exception**

仅当某页无法通过 donor 复制、且动画仅为简单淡入时，另建 `com_fade_sequence` 实验路线。它不得进入 production MCP，直至拥有独立的 donor 对照、回放证据和人工批准。

**Step 7: Run all tests**

Run: `npm test`

Expected: PASS。

---

### Task 6：新增单一受控的 MCP 生成工具

**Objective:** 将经过 plan 校验的 style/motion 执行能力暴露给 Hermes，同时保持执行需要明确确认的安全边界。

**Files:**
- Modify: `src/mcp-server.js:25-83`
- Create: `src/generate-native-deck.js`
- Modify: `tests/mcp-server.test.js`（如现有测试框架缺失则新建）
- Modify: `README.md`

**Step 1: Define MCP tool contract**

新增：`ppt_creater_generate_native_deck`。

输入：
- `target` 或 `family_id`（二选一）；
- `title`、`slides`；
- `palette_id`（可选，默认 `native`）；
- `allow_motion_variants: boolean`（默认 false；仅 `copy_verified_donor` 需要显式开启）；
- `allowExternal: boolean`（沿用现有版权闸门）；
- `output_name`（仅安全文件名，不接受路径）。

输出：
- output PPTX path；
- plan；
- `palette_audit`；
- `motion_audit`；
- `strict_native_only`、`new_shapes`、`powerpoint_playback`；
- 若任何强制验证未通过，返回 error 而不是半成品成功信息。

**Step 2: Implement safety gates**

- 默认 `allow_motion_variants: false`；保留 donor 原动画不需要开启，选择 `copy_verified_donor` 变体时必须为 true。
- 对 external family 仍必须 `allowExternal: true`。
- 只允许 production-ready master、registry 关联的 capability manifest 和其中登记的 donor deck。
- 颜色切换或动画 donor 变体前，要求匹配的 manifest `verified: true`。
- `fillPlan`、style executor、motion donor executor 中任一失败时删除或隔离临时输出，不覆盖任何 master。

**Step 3: Add MCP-level tests**

以 stub executor 验证：schema 默认值、颜色/动画闸门、external license 闸门、拒绝路径、输出序列和 audit 汇总。不得在 Node 单元测试中启动实际 PowerPoint。

**Step 4: Perform an end-to-end manual verification**

对一个用户拥有的已生产就绪 family：
1. 生成仅替换 slot 文本的原生 deck；
2. 生成同内容的 palette 变体；
3. 若该 family 已完成 motion capability 标注，生成一个动画预设变体；
4. 用 PowerPoint 手工检查文本继承、布局、动画节奏；
5. 核对所有 audit 为 `verified`、`new_shapes: 0`、geometry unchanged、playback true。

---

### Task 7：注册表升级、文档与回归矩阵

**Objective:** 把 capability 状态纳入生产可用性判断，防止未验证的 style/motion 能力被路由或执行。

**Files:**
- Modify: `catalog/unified-production-template-library.json`
- Modify: `src/production-library.js:8-44`
- Modify: `src/target-router.js:7-28`
- Modify: `README.md`
- Create: `catalog/QUALITY_RUBRIC.md` 中的 capability 增补章节，或新建 `catalog/CAPABILITY_QUALITY_RUBRIC.md`
- Test: `tests/production-library.test.js`, `tests/target-router.test.js`

**Step 1: Add optional capability summary fields**

对每个 template 增加：

```json
{
  "capabilities_path": "C:\\ppt-creater\\catalog\\capabilities\\local-modern-report.json",
  "palette_status": "native_only | verified",
  "motion_status": "preserve_only | verified_presets",
  "capabilities_verified_at": "ISO-8601 timestamp",
  "capability_validation_pptx": "..."
}
```

不要把所有 families 一次性标记为 verified；默认均为 `native_only` / `preserve_only`。仅完成 Task 3–5 实证的 family 可升级。

**Step 2: Tighten routing and production assertions**

- 所有 family 可继续作为普通原生填充的 route；
- 只有 `palette_status: verified` 时允许 non-native palette；
- 只有 `motion_status: verified_donor_variants` 时允许 `copy_verified_donor`；
- `strict_native_only`、`new_shapes: 0`、timeline signature match 与 PowerPoint playback 仍是生成成功的必要条件。

**Step 3: Build regression matrix**

将以下案例写入文档并尽可能自动化：

| Case | Expected |
|---|---|
| 原生 palette + preserve original | 允许；无颜色/动画改动 |
| 已验证 semantic palette | 允许；只登记形状改色 |
| 未验证 palette | 拒绝 |
| 已验证 fade sequence | 允许；播放验证通过 |
| 未登记 slide motion preset | 拒绝 |
| 动画锁定对象要求改色 | 拒绝 |
| 未通过外部授权的 external family | 拒绝 |
| 执行后 shape count/geometry 不一致 | 失败并隔离输出 |

**Step 4: Run final verification**

Run:

```bash
npm test
python -m pytest tests -q
```

Expected: Node suite passes; Python suite result must be reported exactly as actually returned. Then run one manually inspected PowerPoint output for each已启用的 capability family.

---

## 关键验收标准

1. MCP 可以列出某 family 支持的 palette 与 motion preset，且不会把未验证能力报成可用。
2. 用户在创建 plan 时选择颜色与动画预设，系统会在执行前拒绝所有未登记、更改动画锁定对象、或跨 family 的请求。
3. 非 native palette 只能改 capability manifest 中的既有对象；输出仍满足 `new_shapes = 0`、形状数量相同、geometry 不变。
4. 动画预设只影响指定页中的指定既有对象；未请求编辑的页时间线不变。
5. 每个动画编辑输出均有 PowerPoint 放映验证；验证失败不应报告生成成功。
6. 新增/替换文字始终通过既有 `slot_*` 对象完成，继承模板原生字体、字号、颜色、字重与段落样式。
7. 外部模板的版权确认与当前 production-ready 校验继续生效。

---

## 风险、取舍与待确认事项

- **PowerPoint COM 依赖 Windows 桌面 PowerPoint。** CI 无法可靠跑放映验证；COM 行为应由本机人工/自动 smoke test 作为发布闸门，Node 测试负责纯逻辑和审计验证。
- **形状名称稳定性是核心前提。** 未命名或名称重复的形状不能进入 capability manifest；应先复制为可编辑 production master，再标注，绝不直接改源文件。
- **动画清除是破坏性操作。** 首期应只有少量经用户审阅的 slide 明确允许 `none`，并默认禁用 motion edits。
- **主题变体支持存在 PowerPoint 文件格式差异。** 首期优先语义 palette；“完整 theme variant 切换”应只在已被 COM 实测的 family 上开启，避免破坏 theme-color 引用。
- **文本颜色切换会违背原生文字样式继承的默认原则。** 建议第一版禁用；仅当用户明确要做品牌重色且手工审查后，在个别 shape 上显式开白名单。
- **需要你决定产品语义：** 颜色切换是否只面向装饰/图表色，还是未来也允许标题/正文的品牌重色。方案默认前者，以保护你已明确要求的原模板文字继承。
- **需要你决定动画编辑的体验：** 建议 MCP 提供“预设选择”，而不是暴露任意动画参数。这样能保证每个 preset 都可回放验证、可复现且符合模板节奏。
