# 商业收入及回款预测

## 项目简介

按用户指定的**基准日**与**门店**，从观远 BI 拉取门店租赁合同数据，生成收入/回款预测流程所需的**中间临时表**（Excel）：

| 临时表 | 说明 | 列数 |
|---|---|---|
| 在执行合同临时表 | 起租日 ≤ 基准日 ≤ 到期日 的存量合同（模式 A） | 26 |
| 未执行合同临时表 | 起租日 > 基准日 且 起租日 < 到期日 的已签约未起租合同（模式 B） | 26 |
| 执行+未执行合同清单 | 两表按（门店, 铺位号）匹配续签关系后合并，并附加收款周期信息 | 32 |
| 合同全周期应收明细 | 按合同号关联应收账单明细（含不含税金额计算） | 19 |
| 续签合同应收明细 | 预测年度内到期未续签合同假定续签，滚动生成应收预测 | 19 |

**数据源**（观远 BI，每日约 02:00 更新快照）：

- `ADS-租费分析-租赁合同台账明细`（dsId `k3a8a272ee4143d694a2ecd`）——合同清单主表，一行一合同；
- `ADS-租费分析-合同全周期应收明细`（dsId `ud1f309f6e04d4285b7364c4`）——收入/回款预测金额主表，一行一条应收账单明细。

各表均为收入/回款预测流程的**中间基础表，不是最终输出表**；完整口径、执行流程与红线约定见 `SKILL.md`，项目背景、已确认决策与变更日志见 `PROJECT_MEMORY.md`。

## 目录结构

```
商业收入及回款预测/
├── README.md                  # 本文件
├── SKILL.md                   # 技能完整执行流程与口径（权威文档）
├── PROJECT_MEMORY.md          # 项目记忆：背景、约定、已确认决策、变更日志
├── references/
│   └── dataset-and-rules.md   # 数据集字段、口径验证记录、门店别名映射
├── scripts/
│   ├── make_report.py         # 在执行/未执行合同临时表生成器（--mode active/future）
│   ├── merge_contracts.py     # 执行+未执行合同清单合并器（续签标记，28 列基础表）
│   ├── receivable_detail.py   # 应收明细临时表生成 + 合并清单收款周期升级（32 列）
│   └── renewal_receivable.py  # 续签合同应收明细生成器（假定续签滚动预测）
└── output/                    # 中间临时表 xlsx 输出目录（不入库）
```

## 环境要求

- **Node.js 20+**（推荐 22+）：运行 guancli（观远 CLI 取数工具）
- **guancli**：已认证并能访问上述两个数据集（需向观远管理员申请数据集权限）
- **Python 3** + **openpyxl**（脚本唯一第三方依赖）
  - CodeBuddy 环境优先使用 `~/.workbuddy/binaries/python/envs/default/bin/python`；该路径不存在时用本机 Python 3

## 安装步骤

### 1. 克隆仓库

```bash
git clone https://github.com/HowardTao/rsun-skill.git
```

### 2. 安装 guancli 并登录

参考官方文档 <https://www.guandata.com/guandata-cli-installation-guide.md>：

```bash
npm install -g --foreground-scripts @guandata/guanskill
guanskill install-skill
guancli auth login    # 配置观远访问凭证（PAT），登录 Domain 默认 guanbi
guancli auth status   # 确认 PAT 有效
```

更新方式：重新执行安装命令即可。

### 3. 安装技能到本地客户端

将技能目录复制到本地技能目录，使 AI 客户端可以加载本技能：

```bash
# WorkBuddy 环境
cp -r rsun-skill/商业收入及回款预测 ~/.workbuddy/skills/
# CodeBuddy 环境
cp -r rsun-skill/商业收入及回款预测 ~/.codebuddy/skills/
```

### 4. 准备 Python 环境

```bash
pip install openpyxl
```

## 使用说明

### 通过 AI 客户端使用（推荐）

将技能目录安装到 `~/.workbuddy/skills/`（或 `~/.codebuddy/skills/`）后，在 CodeBuddy / WorkBuddy 中直接对话触发，例如：

> 帮我生成弘阳家居南京江北店基准日 2026-08-31 的在执行合同临时表
> 生成江北店 2026 年收入回款预测的基础表

AI 会按 `SKILL.md` 的流程执行：确认基准日 → 确认预估年度（收入/回款预测必问）→ 确认门店与铺位类型 → guancli 取数 → 校验 → 生成 Excel 到 `output/`。

**交互约定（红线）**：

- **基准日**与**门店**必须由用户提供并确认，AI 不得自行假设默认值；
- 执行收入/回款预测时必须先确认**预测年度**（默认当前自然年，可自定义）；
- 在执行/未执行口径为日期窗口判定（`合同状态` 字段只展示、绝不作为筛选条件）；
- 铺位类型默认包含正铺 + 多经，仅用户明确指定时才筛选。

### 手动执行脚本

先按 `SKILL.md` 第 4 步用 guancli 取数到 `/tmp/*.json`，再运行脚本。命令模板（`<...>` 为占位符）：

```bash
PY=~/.workbuddy/binaries/python/envs/default/bin/python
SKILL=<技能目录>

# 1. 在执行 / 未执行合同临时表（26 列）
$PY $SKILL/scripts/make_report.py \
  --mode active \
  --input /tmp/active_contracts_<基准日>.json \
  --base-date <基准日> --store <门店全称> \
  --output $SKILL/output/<门店简称>在执行合同_<基准日>.xlsx \
  --snapshot-time "<数据集更新时间>"

# 2. 合并清单（28 列，含续签标记）
$PY $SKILL/scripts/merge_contracts.py \
  --active /tmp/active_contracts_<基准日>.json \
  --future /tmp/future_contracts_<基准日>.json \
  --base-date <基准日> --store <门店全称> \
  --output $SKILL/output/执行+未执行合同清单_<基准日>.xlsx \
  --snapshot-time "<数据集更新时间>"

# 3. 应收明细临时表（19 列）+ 合并清单升级（32 列，含最后账期/支付周期/提前收款天数）
$PY $SKILL/scripts/receivable_detail.py \
  --active /tmp/active_contracts_<基准日>.json \
  --future /tmp/future_contracts_<基准日>.json \
  --receivable "/tmp/receivable_batches_<基准日>/b_*.json" \
  --base-date <基准日> --store <门店全称> \
  --output-detail $SKILL/output/<门店简称>合同全周期应收明细_<基准日>.xlsx \
  --output-merged $SKILL/output/执行+未执行合同清单_<基准日>.xlsx \
  --snapshot-time "<数据集更新时间>"

# 4. 续签合同应收明细（19 列，假定续签滚动预测）
$PY $SKILL/scripts/renewal_receivable.py \
  --merged $SKILL/output/执行+未执行合同清单_<基准日>.xlsx \
  --forecast-year <预测年度> --base-date <基准日> --store <门店全称> \
  --output $SKILL/output/<门店简称>续签合同应收明细_<预测年度>年_<基准日>.xlsx \
  --snapshot-time "<数据集更新时间>"
```

**取数约定**：

- 一律用 `guancli ds preview --filter`，不用 `ds execute-sql`（该环境报 Spark `PARSE_SYNTAX_ERROR`）；
- 日期过滤条件的时间部分必须写 `00:00:00`（字段为零点时间戳）；
- 合同清单取数 `--limit 2000`、应收明细按合同号 `IN` 每 50 个一批 `--limit 10000`；行数达到 limit 时需分批处理；
- 所有中间临时表 xlsx 统一输出到技能目录 `output/`，命名 `<门店简称><表名>_<基准日>.xlsx`。

各脚本的参数与口径详见脚本头部注释及 `SKILL.md`。

## 贡献指南

1. **口径红线不可绕过**：在执行 = 起租日 ≤ 基准日 ≤ 到期日；未执行 = 起租日 > 基准日；`合同状态` 只展示不筛选；铺位类型默认正铺 + 多经。任何口径调整须先与项目负责人确认，并记入 `PROJECT_MEMORY.md` 的「已确认的决策」。
2. **文档同步**：修改脚本功能、参数、输出列或新增数据源时，必须同步更新 `SKILL.md`、相关脚本头部注释，以及本 README 中对应的功能描述、使用方式和配置说明，确保文档与实际代码一致。
3. **变更登记**：每次功能性变更须在 `PROJECT_MEMORY.md` 末尾的「变更日志」追加一行（日期 + 变更摘要）。
4. **代码规范**：脚本保持"参数必填、无硬编码业务值"（基准日/预测年度/门店等均通过命令行参数传入）；新增依赖需在 `SKILL.md` 运行环境一节与本 README「环境要求」同步声明。
5. **提交约定**：提交信息用中文简要描述变更内容；`output/` 下的临时产物与 `__pycache__/`、`*.pyc` 不入库。
