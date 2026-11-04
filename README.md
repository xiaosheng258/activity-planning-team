# 活动策划天团（campus-event-planner）

多角色协同的活动策划专家团：输入活动基本信息（名称、类型、规模、日期、场地），并行产出策划方案、预算、风险预案与宣传文案，最终汇编为一份完整的 HTML 策划方案，并附预算 Excel、风险预案 Excel、宣传 PPTX 三份附件。

## 快速开始

```bash
git clone https://github.com/xiaosheng258/activity-planning-team.git campus-event-planner
```

放入 `~/.workbuddy-ai/plugins/marketplaces/my-experts/plugins/` 后注册一次即可使用，详见下方[安装到自己的 WorkBuddy](#安装到自己的-workbuddy)。

## 类型

Team 型（多角色协作团队）· 分类：09-OperationsHR（运营人力）· 许可：MIT

## 团队成员

| Agent ID | 名字 | 职业 | 负责段落 |
|---|---|---|---|
| campus-event-planner-team-lead | 周全成 | 策划总监 | 编排调度、交叉核对、整合输出 |
| proposal-writer | 陈述之 | 方案主笔 | 一主题 / 二背景意义 / 三时间 / 四参与人员 / 五内容与流程 / 八总纲 |
| logistics-coordinator | 彭备齐 | 后勤筹备官 | 六活动准备（物料、场地、分工与联系方式） |
| budget-analyst | 钱算清 | 预算分析师 | 七活动预算（项目/数量/单价/总价/合计）+ 预算 Excel |
| risk-manager | 卫平安 | 风险管控师 | 八注意事项深化（风险分级与应急）+ 风险预案 Excel |
| promo-copywriter | 宣达远 | 宣传文案师 | 宣传文案 + 宣传 PPTX |

## 工作流程（SOP）

- **Phase 1（并行）**：主理人确认活动五要素 → TeamCreate → 同时调度五位团员各自独立产出。
- **Phase 2（串行）**：主理人回收五份产出，核对金额、时间、地点一致性，汇编为八段式 HTML 策划方案并附三份附件。

## 内置技能

`skills/event-planning-kit` — 固化「策划书八段式」「预算表」「风险预案」「宣传文案」四套模板，统一全团产出格式。

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

## 安装到自己的 WorkBuddy

### 1. 下载

```bash
git clone https://github.com/xiaosheng258/activity-planning-team.git campus-event-planner
```

注意末尾的 `campus-event-planner`：**目录名必须和专家包内部标识一致**，否则无法被识别。

不想用 Git 的话，在仓库页面点 **Code → Download ZIP**，解压后把文件夹重命名为 `campus-event-planner`。

### 2. 放入专家目录

把整个文件夹放到 `my-experts/plugins/` 下：

| 系统 | 路径 |
|------|------|
| Windows | `C:\Users\<你的用户名>\.workbuddy-ai\plugins\marketplaces\my-experts\plugins\` |
| macOS / Linux | `~/.workbuddy-ai/plugins/marketplaces/my-experts/plugins/` |

放好后的结构应为：

```
my-experts/plugins/campus-event-planner/
├─ .codebuddy-plugin/plugin.json
├─ agents/
├─ skills/
├─ avatars/
├─ README.md
└─ settings.json
```

### 3. 注册

仅放入文件不会自动出现在面板里，需要注册一次。在 WorkBuddy 中直接说：

> 帮我把 campus-event-planner 这个专家包注册到我的专家列表

或者如果你本机有专家管理脚本，执行：

```bash
python3 <expert-manager>/scripts/register_expert.py <上述 campus-event-planner 路径>
```

> `register_expert.py` 属于 WorkBuddy 内置的 expert-manager 技能，不在本仓库内。仓库只包含专家包本体。

### 4. 使用

注册成功后，在左侧「专家」→「我的专家」中找到「活动策划天团」，发送活动基本信息即可。


## 自定义

想改造这套专家团，主要改这几处：

| 想改什么 | 改哪个文件 |
|---------|-----------|
| 团队成员的名字/职业/职责 | `.codebuddy-plugin/plugin.json` 的 `members` + 对应 `agents/*.md` 的 frontmatter |
| 某个角色产出什么 | `agents/<角色>.md` 的「输出规范」章节 |
| 策划书段落结构 | `skills/event-planning-kit/references/策划书八段式模板.md` |
| 团队协作流程 | `agents/campus-event-planner-team-lead.md` 的「标准工作流程」 |
| 头像 | 直接替换 `avatars/` 下同名 PNG（512×512，≤500KB） |

改完 `plugin.json` 后需要重新注册一次才会生效。

## 许可

[MIT](LICENSE) © 2026 胡一鸣
