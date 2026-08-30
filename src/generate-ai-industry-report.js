const pptxgen = require('pptxgenjs');
const path = require('node:path');

const OUT = path.resolve(__dirname, '..', 'output', '2026-ai-industry-report-cn.pptx');
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Hermes / ppt-creater';
pptx.subject = 'AI 行业汇报';
pptx.title = '2026 AI 行业：从模型能力竞争走向系统价值兑现';
pptx.company = 'ppt-creater';
pptx.lang = 'zh-CN';
pptx.theme = {
  headFontFace: 'Microsoft YaHei', bodyFontFace: 'Microsoft YaHei', lang: 'zh-CN'
};
pptx.defineLayout({ name: 'CUSTOM_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'CUSTOM_WIDE';

const C = { bg:'F5F8FC', navy:'0B1F3A', navy2:'123A63', ink:'172033', muted:'64748B', accent:'00A6A6', accentLight:'DFF6F4', orange:'F59E0B', line:'DCE5EF', white:'FFFFFF', red:'D9485F' };
const F = 'Microsoft YaHei';
const W=13.333, H=7.5, M=.58;
function addText(slide, text, opt={}) { slide.addText(text, { fontFace:F, color:C.ink, margin:0, breakLine:false, fit:'shrink', ...opt }); }
function bg(slide, dark=false) { slide.background={color:dark?C.navy:C.bg}; }
function footer(slide, n, source='资料：公开公司材料、行业报告与分析框架；仅作研究讨论', dark=false) {
  slide.addShape(pptx.ShapeType.line,{x:M,y:7.06,w:12.17,h:0,line:{color:dark? '315476':C.line,pt:.7}});
  addText(slide,source,{x:M,y:7.14,w:10.8,h:.16,fontSize:7.5,color:dark?'A9C2D7':C.muted});
  slide.addShape(pptx.ShapeType.ellipse,{x:12.25,y:7.09,w:.34,h:.23,fill:{color:dark?C.accent:C.navy},line:{color:dark?C.accent:C.navy}});
  addText(slide,String(n).padStart(2,'0'),{x:12.25,y:7.12,w:.34,h:.12,fontSize:7.5,bold:true,align:'center',color:C.white});
}
function title(slide, heading, kicker='', n, dark=false) {
  if(kicker) addText(slide,kicker.toUpperCase(),{x:M,y:.42,w:3.2,h:.22,fontSize:8.5,bold:true,color:dark?'6EE7E0':C.accent,charSpace:1.2});
  addText(slide,heading,{x:M,y:.72,w:11.8,h:.62,fontSize:dark?27:25,bold:true,color:dark?C.white:C.navy,breakLine:false});
  footer(slide,n,undefined,dark);
}
function card(slide,x,y,w,h,head,body,accent=C.accent) {
  slide.addShape(pptx.ShapeType.roundRect,{x,y,w,h,rectRadius:.08,fill:{color:C.white},line:{color:C.line,pt:.7}});
  slide.addShape(pptx.ShapeType.rect,{x,y,w:.07,h,fill:{color:accent},line:{color:accent}});
  addText(slide,head,{x:x+.22,y:y+.18,w:w-.36,h:.28,fontSize:14,bold:true,color:C.navy});
  addText(slide,body,{x:x+.22,y:y+.64,w:w-.38,h:h-.8,fontSize:11.5,color:C.muted,breakLine:true,breakLine:false,fit:'shrink'});
}
function pill(slide,x,y,label,color=C.accent) { slide.addShape(pptx.ShapeType.roundRect,{x,y,w:1.25,h:.34,rectRadius:.08,fill:{color},line:{color}}); addText(slide,label,{x,y:y+.09,w:1.25,h:.12,fontSize:8,bold:true,color:C.white,align:'center'}); }

// 1 cover
{ const s=pptx.addSlide(); bg(s,true);
  s.addShape(pptx.ShapeType.arc,{x:8.3,y:-1.7,w:6,h:6,adjustPoint:.25,line:{color:'1B5677',pt:1.2,transparency:25},adjustPoint:.4});
  s.addShape(pptx.ShapeType.arc,{x:9.1,y:-.8,w:4.6,h:4.6,adjustPoint:.25,line:{color:C.accent,pt:2,transparency:15},adjustPoint:.4});
  addText(s,'AI INDUSTRY BRIEFING',{x:M,y:.76,w:4,h:.25,fontSize:10,bold:true,color:'6EE7E0',charSpace:1.5});
  addText(s,'2026 AI 行业：\n从模型能力竞争走向系统价值兑现',{x:M,y:1.55,w:8.5,h:1.5,fontSize:35,bold:true,color:C.white,breakLine:true,fit:'shrink'});
  addText(s,'面向技术、产业与投资决策者的十页研究简报',{x:M,y:3.45,w:5.5,h:.3,fontSize:15,color:'BED1E0'});
  s.addShape(pptx.ShapeType.line,{x:M,y:4.28,w:1.55,h:0,line:{color:C.accent,pt:2.5}});
  addText(s,'ppt-creater · AI 行业研究模板库 · 2026',{x:M,y:6.6,w:6,h:.22,fontSize:9.5,color:'A9C2D7'});
}
// 2 executive summary
{ const s=pptx.addSlide(); bg(s); title(s,'行业主线已从“更大模型”转向“更可交付的 AI 系统”','EXECUTIVE SUMMARY',2);
  card(s,.62,1.75,3.82,3.86,'01｜价值重心迁移','模型能力仍在推进，但采购、部署与组织改造决定了价值能否兑现。',C.accent);
  card(s,4.75,1.75,3.82,3.86,'02｜竞争单元升级','竞争不再是单模型对比，而是“模型 × 数据 × 工作流 × 分发 × 可信治理”。',C.orange);
  card(s,8.88,1.75,3.82,3.86,'03｜投资判断变化','关注可验证 ROI、专有数据闭环、落地摩擦与单位经济，而非单点参数叙事。',C.navy2);
  addText(s,'结论：未来 12–24 个月，能够把 AI 嵌入关键业务闭环的公司，才可能将技术势能转化为持续收入。',{x:.7,y:6.05,w:11.7,h:.35,fontSize:14,bold:true,color:C.navy,align:'center'});
}
// 3 market map
{ const s=pptx.addSlide(); bg(s); title(s,'AI 产业已形成四层价值栈：从算力供给到行业交付','MARKET MAP',3);
  const layers=[['应用与工作流','行业 Copilot、智能体、自动化系统','00A6A6'],['模型与平台','基础模型、多模态、推理、评测与部署','247BA0'],['数据与工具链','数据治理、RAG、向量检索、观测与安全','F59E0B'],['基础设施','芯片、云、网络、存储与能耗优化','0B1F3A']];
  layers.forEach((r,i)=>{const y=1.6+i*1.12; s.addShape(pptx.ShapeType.roundRect,{x:1.15,y,w:10.95,h:.82,rectRadius:.08,fill:{color:r[2]},line:{color:r[2]}}); addText(s,r[0],{x:1.48,y:y+.23,w:2.0,h:.25,fontSize:16,bold:true,color:C.white}); addText(s,r[1],{x:4.1,y:y+.26,w:6.7,h:.18,fontSize:12,color:C.white});});
  addText(s,'观察重点：应用层商业化速度最快；模型与基础设施层的利润结构更依赖规模、供给与资本效率。',{x:1.15,y:6.26,w:11,h:.28,fontSize:12.5,color:C.muted,align:'center'});
}
// 4 technology stack
{ const s=pptx.addSlide(); bg(s,true); title(s,'技术栈的关键分水岭：推理、工具调用与反馈闭环','TECHNOLOGY STACK',4,true);
  const boxes=[['感知与上下文','多模态输入\n企业知识与状态'],['推理与世界建模','理解任务\n预测行动后果'],['规划与工具调用','分解目标\n调用系统与协作'],['执行与反馈','写入业务系统\n监控、纠错与学习']];
  boxes.forEach((b,i)=>{let x=.75+i*3.08; s.addShape(pptx.ShapeType.roundRect,{x,y:2.05,w:2.52,h:2.42,rectRadius:.08,fill:{color:'113A5E'},line:{color:i===1?C.accent:'315476',pt:i===1?1.8:.7}}); addText(s,String(i+1).padStart(2,'0'),{x:x+.23,y:2.3,w:.4,h:.16,fontSize:9,bold:true,color:'6EE7E0'}); addText(s,b[0],{x:x+.23,y:2.7,w:2.0,h:.38,fontSize:15,bold:true,color:C.white}); addText(s,b[1],{x:x+.23,y:3.45,w:1.95,h:.54,fontSize:11,color:'B9D0E2',breakLine:true}); if(i<3) s.addShape(pptx.ShapeType.chevron,{x:x+2.57,y:3.0,w:.42,h:.5,fill:{color:C.accent},line:{color:C.accent}}); });
  addText(s,'“模型”正在成为系统能力的一部分；真正的差异化来自连接真实数据、工具与持续反馈的闭环。',{x:1.0,y:5.5,w:11.2,h:.38,fontSize:15,bold:true,color:'DFF6F4',align:'center'});
}
// 5 value chain
{ const s=pptx.addSlide(); bg(s); title(s,'价值链重构：高价值环节向“最后一公里交付”集中','VALUE CHAIN',5);
  const items=[['模型供给','能力商品化加速','模型能力趋同，价格与推理效率成为关键'],['平台集成','连接企业系统','数据、权限、工作流与可观测性决定部署成本'],['行业应用','直接承担业务结果','围绕高频、刚需、可量化 ROI 的场景更易付费'],['服务与治理','保证可靠运行','安全、合规、评测、变更管理成为长期需求']];
  items.forEach((a,i)=>{const x=.75+i*3.13; s.addShape(pptx.ShapeType.rect,{x,y:1.95,w:2.55,h:3.25,fill:{color:i===2?'E8FAF8':C.white},line:{color:i===2?C.accent:C.line,pt:i===2?1.6:.8}}); addText(s,String(i+1),{x:x+.22,y:2.18,w:.3,h:.2,fontSize:11,bold:true,color:i===2?C.accent:C.muted}); addText(s,a[0],{x:x+.22,y:2.6,w:2.05,h:.3,fontSize:15,bold:true,color:C.navy}); addText(s,a[1],{x:x+.22,y:3.22,w:2.0,h:.25,fontSize:12,bold:true,color:i===2?C.accent:C.orange}); addText(s,a[2],{x:x+.22,y:3.78,w:2.02,h:.8,fontSize:10.7,color:C.muted,breakLine:true}); if(i<3)s.addShape(pptx.ShapeType.rightArrow,{x:x+2.58,y:3.23,w:.34,h:.35,fill:{color:C.line},line:{color:C.line}})});
  addText(s,'最有价值的公司不是“有 AI”，而是能将 AI 与客户的核心流程、数据资产和结果指标深度绑定。',{x:.95,y:5.8,w:11.3,h:.3,fontSize:13.5,bold:true,color:C.navy,align:'center'});
}
// 6 metrics
{ const s=pptx.addSlide(); bg(s); title(s,'投资评估应从“能力叙事”升级为“可验证的单位经济”','INVESTMENT LOGIC',6);
  const metrics=[['Time-to-Value','首个可量化业务成果','< 90 天'],['Adoption Depth','嵌入工作流的深度','日常使用'],['Gross Margin','推理与交付后的毛利','可扩张'],['Data Flywheel','专有反馈带来的提升','可复用']];
  metrics.forEach((m,i)=>{let x=.72+i*3.13; s.addShape(pptx.ShapeType.roundRect,{x,y:1.85,w:2.55,h:2.62,rectRadius:.08,fill:{color:C.white},line:{color:C.line,pt:.8}}); addText(s,m[0],{x:x+.22,y:2.14,w:2.08,h:.22,fontSize:10.5,bold:true,color:C.accent,align:'center'}); addText(s,m[2],{x:x+.22,y:2.66,w:2.1,h:.46,fontSize:25,bold:true,color:C.navy,align:'center'}); addText(s,m[1],{x:x+.25,y:3.55,w:2.0,h:.35,fontSize:10.5,color:C.muted,align:'center'});});
  s.addShape(pptx.ShapeType.roundRect,{x:1.2,y:5.35,w:10.9,h:.7,rectRadius:.08,fill:{color:C.navy},line:{color:C.navy}}); addText(s,'建议：以“业务结果—使用深度—交付成本—数据闭环”四项证据替代泛化的参数、估值与客户数量叙事。',{x:1.48,y:5.61,w:10.35,h:.2,fontSize:13,bold:true,color:C.white,align:'center'});
}
// 7 risk matrix
{ const s=pptx.addSlide(); bg(s); title(s,'风险不是单一技术问题，而是可靠性、成本与组织变革的组合','RISK MATRIX',7);
  s.addShape(pptx.ShapeType.rect,{x:2.1,y:1.65,w:8.8,h:4.35,fill:{color:C.white},line:{color:C.line,pt:.9}});
  s.addShape(pptx.ShapeType.line,{x:6.5,y:1.65,w:0,h:4.35,line:{color:C.line,pt:.8}}); s.addShape(pptx.ShapeType.line,{x:2.1,y:3.82,w:8.8,h:0,line:{color:C.line,pt:.8}});
  addText(s,'影响高',{x:1.25,y:1.7,w:.6,h:.2,fontSize:10,color:C.muted,rotate:270}); addText(s,'影响低',{x:1.25,y:5.32,w:.6,h:.2,fontSize:10,color:C.muted,rotate:270}); addText(s,'发生概率低',{x:2.15,y:6.15,w:1.2,h:.2,fontSize:10,color:C.muted}); addText(s,'发生概率高',{x:9.65,y:6.15,w:1.2,h:.2,fontSize:10,color:C.muted});
  const risks=[['模型幻觉与可靠性',7.2,2.15,C.red],['推理成本失控',7.7,4.35,C.orange],['数据/权限边界',3.0,2.45,C.red],['组织采用阻力',3.25,4.8,C.orange]];
  risks.forEach(r=>{s.addShape(pptx.ShapeType.ellipse,{x:r[1],y:r[2],w:1.35,h:.48,fill:{color:r[3],transparency:8},line:{color:r[3]}}); addText(s,r[0],{x:r[1]+.06,y:r[2]+.16,w:1.23,h:.12,fontSize:7.8,bold:true,color:C.white,align:'center'});});
  addText(s,'管理动作：建立高风险场景准入、离线/在线评测、人工兜底、成本观测和分阶段扩展机制。',{x:1.2,y:6.55,w:10.9,h:.22,fontSize:12,bold:true,color:C.navy,align:'center'});
}
// 8 roadmap
{ const s=pptx.addSlide(); bg(s,true); title(s,'未来 24 个月：从助手、智能体到跨系统自主协作','ROADMAP',8,true);
  s.addShape(pptx.ShapeType.line,{x:1.08,y:3.57,w:11.15,h:0,line:{color:'4F7694',pt:2}});
  const phases=[['0–6 个月','Copilot 普及','单点任务辅助\n知识检索与内容生成'],['6–12 个月','工作流智能体','跨工具编排\n可观测、可评测'],['12–24 个月','结果导向系统','围绕 KPI 自主执行\n人机协作与治理']];
  phases.forEach((p,i)=>{let x=1.15+i*3.75; s.addShape(pptx.ShapeType.ellipse,{x,y:3.3,w:.5,h:.5,fill:{color:C.accent},line:{color:C.accent}}); addText(s,p[0],{x:x-.1,y:2.1,w:1.4,h:.22,fontSize:12,bold:true,color:'6EE7E0'}); addText(s,p[1],{x:x-.1,y:2.55,w:2.9,h:.32,fontSize:18,bold:true,color:C.white}); addText(s,p[2],{x:x-.1,y:4.15,w:2.8,h:.55,fontSize:11,color:'B9D0E2',breakLine:true});});
  addText(s,'核心判断：AI 的边界会从“生成内容”持续扩展到“承担可审计的业务结果”。',{x:1.2,y:5.9,w:10.9,h:.28,fontSize:14,bold:true,color:C.white,align:'center'});
}
// 9 recommendations
{ const s=pptx.addSlide(); bg(s); title(s,'行动建议：以业务闭环为单位，而非以模型能力为单位推进 AI','RECOMMENDATION',9);
  const acts=[['01','锁定高价值任务','选择高频、强约束、结果可衡量的任务，而非泛化试点。'],['02','建立可观测闭环','将质量、时延、成本、人工介入与业务结果纳入同一仪表盘。'],['03','沉淀专有数据资产','把交互、反馈和失败案例转化为可复用的领域数据飞轮。'],['04','设计治理与扩张机制','先定义权限、责任、评测和回滚，再推进跨团队复制。']];
  acts.forEach((a,i)=>{let y=1.55+i*1.12; addText(s,a[0],{x:.88,y:y+.15,w:.5,h:.22,fontSize:16,bold:true,color:C.accent}); s.addShape(pptx.ShapeType.line,{x:1.55,y:y+.29,w:.5,h:0,line:{color:C.line,pt:1}}); addText(s,a[1],{x:2.3,y:y+.05,w:2.3,h:.28,fontSize:15,bold:true,color:C.navy}); addText(s,a[2],{x:4.85,y:y+.07,w:6.8,h:.27,fontSize:11.5,color:C.muted});});
  s.addShape(pptx.ShapeType.roundRect,{x:.82,y:6.15,w:11.7,h:.5,rectRadius:.07,fill:{color:C.accentLight},line:{color:C.accentLight}}); addText(s,'一句话：把 AI 当作“可持续学习的业务系统”来建设，而不是一次性部署的软件功能。',{x:1.1,y:6.33,w:11.1,h:.14,fontSize:11.5,bold:true,color:C.navy,align:'center'});
}
// 10 sources
{ const s=pptx.addSlide(); bg(s,true); title(s,'结论与资料边界','CLOSING',10,true);
  addText(s,'AI 行业的核心机会在于：\n把模型能力嵌入真实数据、工作流与结果责任。',{x:1.1,y:1.7,w:10.9,h:.9,fontSize:27,bold:true,color:C.white,align:'center',breakLine:true});
  s.addShape(pptx.ShapeType.line,{x:4.85,y:3.05,w:3.65,h:0,line:{color:C.accent,pt:2.3}});
  const srcs=['资料类型：公开公司材料、模型发布说明、行业研究与政策文件','方法：趋势框架与产业逻辑整理，不构成投资建议','使用建议：在正式外部发布前，替换为可追溯的具体数据、来源与日期'];
  srcs.forEach((t,i)=>{s.addShape(pptx.ShapeType.ellipse,{x:2.1,y:4.05+i*.52,w:.14,h:.14,fill:{color:C.accent},line:{color:C.accent}}); addText(s,t,{x:2.45,y:4.02+i*.52,w:8.5,h:.18,fontSize:11,color:'BED1E0'});});
  addText(s,'ppt-creater · ai-industry-navy-v1',{x:4.35,y:6.37,w:4.7,h:.22,fontSize:9,bold:true,color:'6EE7E0',align:'center'});
}

pptx.writeFile({ fileName: OUT });
