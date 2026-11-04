# 活动策划天团（campus-event-planner）

多角色协同的活动策划专家团：输入活动基本信息（名称、类型、规模、日期、场地），并行产出策划方案、预算、风险预案与宣传文案，最终汇编为一份完整的 HTML 策划方案，并附预算 Excel、风险预案 Excel、宣传 PPTX 三份附件。

## 类型

Team 型（多角色协作团队）· 分类：09-OperationsHR（运营人力）

## 团队成员

| Agent ID | 名字 | 职业 | 负责内容 |
|---|---|---|---|
| campus-event-planner-team-lead | 周全成 | 策划总监 | 编排调度、交叉核对、整合输出 |
| proposal-writer | 陈述之 | 方案主笔 | 一~六 段 + 八总纲（含物料/场地/分工） |
| budget-analyst | 钱算清 | 预算分析师 | 七、预算（知识库取价）+ 预算 Excel |
| risk-manager | 卫平安 | 风险管控师 | 八、风险预案（四类全覆盖）+ 风险预案 Excel |
| promo-copywriter | 宣达远 | 宣传文案师 | 附录 A 宣传文案 + PPTX |

## 工作流程（SOP）

- **Phase 1（并行）**：主理人确认活动六要素（名称/类型/级别/规模/日期/场地）→ TeamCreate → 同时调度四位团员。
- **Phase 2（串行）**：主理人回收四份产出，核对金额、时间、地点一致性，汇编为八段式 HTML 方案。
- **性能优化**：路径传递不传全文、成员完成即待命、文案数据与 PPTX 渲染解耦。预期总耗时约 4-5 分钟。

## 内置技能与知识库

`skills/event-planning-kit` — 四套格式模板 + 四份角色专属知识库 + 三个生成脚本：

| 知识库 | 使用者 | 内容 |
|--------|--------|------|
| `zhishiku/活动统筹全流程标准.md` | 方案主笔 | 分级、场地、审批、里程碑、物料基线 |
| `zhishiku/预算标准.md` | 预算分析师 | 分级上限、单价参考、计算规则、省钱策略 |
| `zhishiku/风险与应急预案标准.md` | 风险管控师 | 四类风险源、分级标准、应急联络 |
| `zhishiku/宣传文案标准.md` | 宣传文案师 | 语气分档、标题公式、文案规范 |

| 脚本 | 产出 |
|------|------|
| `scripts/gen_budget_excel.py` | 预算 Excel（明细含活公式 + 分类汇总） |
| `scripts/gen_risk_excel.py` | 风险预案 Excel（等级排序着色 + 应急联络） |
| `scripts/gen_pptx.py` | 宣传 PPTX（4 页推广版） |

> 脚本依赖 `openpyxl` 与 `python-pptx`，缺失时 `pip install openpyxl python-pptx`。

> 标 ⚠️ 的条目（预算上限、应急电话）需按本校实际替换。改知识库不需要动代码。

## 输出目录

```
output/
├── 预算/            预算表_<活动名>.xlsx
├── 风险预案/        风险预案_<活动名>.xlsx
├── 宣传物料/        宣传_<活动名>.pptx
└── 策划书_<活动名>_<YYYYMMDD>.html
```

## 使用示例

- 帮我策划一场社团迎新破冰活动，规模 80 人，下周五晚 7 点在大学生活动中心多功能厅，要完整策划书、预算、风险预案和宣传文案。
- 给一场讲座活动做一份详细预算表，列出场地、物料、餐饮、奖品等支出项目与单价合计。
- 为我的活动写一版面向大一新生的宣传推送文案和海报文案，风格活泼有网感。

## 头像

头像位于 `avatars/`，当前为本地生成的品牌占位图（深蓝灰底 + 姓名/职能）。可直接替换为真人插画风格，要求：

- 格式：PNG（推荐）或 JPG
- 尺寸：512×512 px
- 大小：单张不超过 500KB
- 建议风格：统一漫画/插画风，灰蓝（navy slate-blue）背景

## 安装

专家包目录已注册，路径：

```
C:\Users\lenovo\.workbuddy-ai\plugins\marketplaces\my-experts\plugins\campus-event-planner
```

如需在其他机器安装，将包放入同级目录后执行：

```bash
python3 scripts/register_expert.py <expert-dir>
```

## 打包分享

```bash
python3 scripts/package_expert.py <expert-dir> <output-dir>
```
