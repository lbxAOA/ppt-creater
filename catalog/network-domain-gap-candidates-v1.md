# 网络模板覆盖缺口补齐：第 1 轮候选清单

**状态：仅候选（`candidate_only`）**。以下链接均已在来源页确认存在 PowerPoint/PPTX 导出入口，但本机直连 Google 导出端点超时，尚未成功下载、渲染或验证。不得用于生产路由，也不得混入现有模板族。

## 当前缺口与候选

| ID | 覆盖缺口（新路线） | 目标能力 / 页面 | 来源页 | PPTX 合法导出链接 | 来源声明 | 当前技术状态 |
|---|---|---|---|---|---|---|
| NET-14 | 医疗健康 / 药企 | 疾病机制、患者风险、临床/药物研究、治疗方案 | [Medical](https://slidesmania.com/medical-free-presentation-template) | [PPTX](https://docs.google.com/presentation/d/1cKg48BXunJHNgSvaqzSdbUgL--4Jp7wKstFhlSvxIis/export/pptx) | 医疗、心血管、药物治疗、药企演示 | 下载端点超时 |
| NET-15 | 教育培训 / 教务 | 课程计划、月/周安排、学生信息、教学仪表盘 | [Teacher Dashboard](https://slidesmania.com/free-teacher-dashboard-august-to-july-version/) | [PPTX](https://docs.google.com/presentation/d/1kwiLbkHo8igqWKbzzhN4CSw-NKkYH4GVNoc87jkjqzg/export/pptx) | 教师数字仪表盘，含月/周视图、笔记、联系人、教案 | 下载端点超时 |
| NET-16 | 金融教育 / 普惠金融 | 预算、信用、储蓄、金融知识科普 | [Personal Finance](https://slidesmania.com/personal-finance-free-presentation-template/) | [PPTX](https://docs.google.com/presentation/d/1ogTknpwbhW55zcpJ5YNUl4LiJwpjkXzB7m0qHILRu6M/export/pptx) | 金钱管理、信用卡、储蓄等主题 | 下载端点超时 |
| NET-17 | 法律 / 合规 / 政策 | 律所介绍、合规治理、案件/政策说明、正式演讲 | [Spelman](https://slidesmania.com/spelman-free-template-for-google-slides-or-powerpoint-presentations/) | [PPTX](https://docs.google.com/presentation/d/1acEVAjhqEx8F0fK6JzbHJy3Q257AUMdcA72tjoXbIso/export/pptx) | 来源明确提及 legal firm 和 law | 下载端点超时 |
| NET-18 | 工程建设 / 制造业 | 工艺流程、项目进度、施工现场、安全与工业流程 | [Caterpillar](https://slidesmania.com/free-industrial-template-for-google-slides-or-powerpoint-presentations) | [PPTX](https://docs.google.com/presentation/d/1d7WbFZJU87NuD9-C2MdpIhHBjiJYhIaXclMJ93sW8a0/export/pptx) | 来源明确覆盖 construction、engineering、industrial processes | 下载端点超时 |
| NET-19 | 能源环保 / ESG | 循环经济、减碳、绿色能源、环境倡议 | [Think GREEN](https://slidesmania.com/think-green-free-presentation-template/) | [PPTX](https://docs.google.com/presentation/d/19pTLR4m7ZgMhaDI4RyTjzeRZtozrZbsVSPtQO6U2VC0/export/pptx) | 回收、绿色能源、可持续发展 | 下载端点超时 |
| NET-20 | 零售消费 / 电商 / 增长 | 市场研究、SWOT、竞品对比、图表和时间线 | [Marketing Analysis](https://slidesmania.com/marketing-analysis-free-presentation-template/) | [PPTX](https://docs.google.com/presentation/d/1OKV7IFBk4VNfQoSx06re9iY9IOvAqtOaeym-YdUM6ks/export/pptx) | 来源明确具备 timelines、charts、SWOT、competitor comparisons | 下载端点超时 |

## 暂不独立建路线的缺口

- **供应链 / 物流**：现有 `technical-product`、`data-boardroom` 与 `b2b-solution` 可承担多数流程、KPI、风险及经营分析表达；若需要仓储/运输网络等行业原生图形，下一轮专门补充。
- **地产 / 建筑**：NET-18 已覆盖工程建设语境；地产销售、楼盘或空间展示则需单独候选。
- **政府 / 政务**：可暂用学术、法律合规及数据经营路由；若需要正式政务视觉、区域地图与公共服务图形，单独建立政务路线。
- **媒体娱乐 / 游戏**：已有营销活动与作品集可应急，但缺少原生影视/游戏叙事页，优先级低于上表 7 类。

## 准入顺序

1. 网络可达后，只下载到 `network-supplements/NET-14` 至 `NET-20`，保留原文件名、哈希、来源页和导出地址。
2. 检查 PPTX 包结构、可编辑对象、动画/切换、页数与完整页面系统。
3. 导出全部页面，生成候选联系表 PDF；此时仍仅为候选。
4. 经人工风格与用途审核后，才可为各模板建立独立 native master、命名 `slot_*`、填充验证并在 PowerPoint 中回放验收。
5. 外部模板的 `license_status` 一律保持 `manual_external_review_required`，直至针对实际使用完成条款审查。
