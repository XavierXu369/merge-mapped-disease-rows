3. Merge Mapped Disease Rows

在疾病 Mapping 完成后，将同一来源记录中映射至同一 Disease 的拆分行进行结构性收敛，并对不同来源记录中的业务等价候选进行显式2. Map Indications to TA Diseases

将已拆分的逐条适应症，映射到某一治疗领域冻结的 Disease CN 白名单，并生成带 Mapping ID、标准化疾病实体、ICD-10 逻辑、状态和 Rationale 的可审计 Mapping 结果。

核心价值

本 Skill 解决“来源适应症如何进入 TA Disease 词池”的问题。

它通过：

冻结本次 TA Disease 白名单；

为每条拆分记录建立唯一 Mapping ID；

将适应症标准化为疾病实体；

必要时使用 ICD-10 作为中间桥梁；

区分成功匹配、待人工判断、其他 TA 和无效信息；

为每一行形成可说服、可复核的 Rationale；

只对成功 Mapped 的记录回填 对应疾病。

它不会修改 TA 白名单，也不会为了提高命中率强行匹配。

适用场景

适用于：

已经完成“一行一个适应症”拆分；

需要将来源适应症映射到指定 TA 的固定 Disease 清单；

需要保留未匹配、跨 TA 和无效信息；

需要形成可复核的 Mapping 结果和完整来源回填。

不适用于：

仍含多个未拆分适应症的来源行；

需要修改或扩充 TA Disease 白名单；

需要合并同 Disease 记录；

需要合并已上市库和临床库；

需要判断资产保留、删除或商业价值。

Mapping 决策单位

判断单位是：

一个 Mapping ID × 一个拆分适应症

Mapping ID 是 Mapping 结果与完整来源底稿之间唯一允许的回填主键。

若来源中没有 Mapping ID，本 Skill按来源顺序创建 1..N；一旦进入决策阶段，不得重新编号。

输入

需要提供：

已拆分的 .xlsx 或 .xlsm 工作簿；

完整拆分后来源 Sheet，而不是三列摘录；

来源 ID、资产名称、拆分适应症字段；

指定 TA；

TA Disease List 及对应 Sheet；

TA、Disease (EN)、Disease (CN) 字段；

经确认的 ICD 标准、版本、来源和证据日期；

必要且经批准的辅助上下文字段；

新的候选、决策和最终输出路径。

默认核心字段为：

序号

药品

适应症拆分结果

TA Disease 白名单规则

每次运行必须重新确认 TA 和白名单版本；

最终 Disease CN 只能来自本次冻结白名单；

Disease CN 必须非空且唯一；

仅做空格层面的等值标准化，不修改原词池；

外部白名单应在结果中保留所选 TA 的快照；

其他 TA 已完成项目中的疾病名称和结果不能直接复用。

五种固定状态

Mapping Status 只能使用以下五种值，且必须使用英文半角连字符：

Mapping Status

Disease CN

含义

Mapped

一个白名单值

已可靠匹配

Manual Review Required - No TA Match

空

是可识别疾病，但当前 TA 白名单没有可接受项

Manual Review Required - Multiple Candidates

空

有两个或以上白名单候选，现有证据无法唯一判断

Unmapped - Other TA

空

人工确认属于其他 TA 或通常由其他科室处理

Unmapped - Invalid Information

空

空值、治疗方式、风险终点、产品用途或其他无效疾病信息

不使用：

泛化的 Unmapped；

Unmapped - No TA Match；

自定义或近似状态名称。

核心判断思路

第一轮：受控 Mapping

优先使用：

拆分适应症；

冻结的 TA Disease 白名单及其层级关系。

第二轮：未决项补充判断

只对第一轮无法关闭的记录使用经批准的辅助信息，例如：

完整获批适应症；

权威 ICD 来源；

监管机构说明；

公司正式产品资料；

其他有针对性的权威证据。

不默认使用靶点、机制、企业、剂型或商业字段推断疾病，除非使用者明确扩大判断边界。

ICD-10 的角色

ICD-10 是推荐的中间桥梁，但不是强制门槛：

来源适应症
→ 标准化疾病实体
→ ICD-10 疾病
→ TA Disease CN

若无法获得可靠的 ICD 代码和名称，可以留空。只有在一个白名单 Disease 临床上明确无歧义时，才可直接 Mapping，并在 Rationale 中解释直接映射逻辑。

Rationale 要求

每一行都必须有完整 Rationale。

Mapped

说明：

原始适应症；

标准化疾病实体；

ICD 到 Disease 的逻辑，或直接 TA 映射的依据。

Manual Review Required - No TA Match

说明：

可识别的疾病实体；

可靠 ICD 逻辑，如有；

为什么现有白名单不能接受；

与现有 TA 词池是完全无关、模糊相关还是上位/下位关系；

尚待人工解决的具体问题。

该状态的 Rationale 在最终结果中应标红。

Manual Review Required - Multiple Candidates

说明：

至少两个准确的白名单候选；

每个候选的依据；

为什么当前证据无法区分；

缺少什么证据；

建议如何处理。

Unmapped - Other TA

说明：

标准化疾病实体；

主要所属 TA 或临床科室；

为什么不应进入当前 TA 池。

Unmapped - Invalid Information

明确写出信息类型，例如：

来源适应症为空；

治疗方式；

风险程度或研究终点；

产品用途；

检测或诊断方式；

宽泛症状；

无法解释的属性信息。

无效信息不得虚构标准化疾病实体或 ICD 信息。

标准流程

只读检查完整拆分底稿、ID、表头、有效行和多适应症残留。

确认 TA、白名单、ICD 版本和 Mapping 边界。

创建或验证 Mapping ID。

冻结本次 TA Disease 白名单。

完成第一轮受控 Mapping。

对未决行完成第二轮轻调研或上下文复核。

生成五状态初步结果。

人工关闭 No TA Match 和 Multiple Candidates。

导入结构化决策。

只读检查状态、Rationale 和回填逻辑。

对保留的人工复核项取得明确批准。

生成最终工作簿并重新打开验证。

输出

Mapping结果

固定包含以下十列，顺序不可改变：

序号

Mapping ID

药品

适应症拆分结果

标准化疾病实体

ICD-10代码

ICD-10疾病名称

Disease CN

Mapping Status

Rationale

完整来源 Sheet 回填

在完整来源 Sheet 中：

增加或保留 Mapping ID；

增加 对应疾病；

Mapped 行回填 Disease CN；

其他四种状态的 对应疾病 必须为空。

最终工作簿保留原输入 Sheet；若 Disease List 来自外部文件，还应保留所选 TA 的白名单快照。

质量门槛与停止条件

必须确认：

来源与 Mapping 结果行数一致；

Mapping ID 非空、唯一、集合与顺序一致；

来源 ID、资产和拆分适应症按 Mapping ID 一致；

每行只有一个固定状态；

每行都有 Rationale；

Mapped 只有一个有效白名单 Disease；

非 Mapped 行的 Disease CN 和回填均为空；

ICD 代码与名称同时有值或同时为空；

Multiple Candidates 的 Rationale 至少包含两个准确候选；

所有保留的人工复核项已明确批准；

No TA Match 的 Rationale 标红规则有效；

原始字段、公式和顺序未被破坏。

遇到白名单、ICD 版本、输入边界、人工复核项或文件指纹不明确时必须停止。

与其他 Skills 的关系

Split Indications and Backfill
→ Map Indications to TA Diseases
→ Merge Mapped Disease Rows

上游负责机械拆分；

本 Skill 负责医学 Mapping；

下游负责在不改变医学结论的前提下合并同 Disease 记录。

本 Skill 不进行跨库合并，也不计算 BI 自有、MOA、VBP 或 TA 领先企业。

跨 TA 复用原则

可以复用：

Mapping ID 机制；

五状态体系；

Rationale 结构；

白名单冻结、人工确认和回填原则；

质量检查与停止条件。

每个 TA 必须重新确认：

Disease 白名单；

ICD 来源和版本；

疾病层级和临床科室边界；

具体适应症结论；

Mapped、Unmapped 和人工复核数量。

文档层级

README 用于说明业务价值和使用方式。状态、字段、命令、配置和验证规则以 SKILL.md、references/input-output-contract.md 及配置示例为正式执行依据。

判断，形成可用于后续跨库合并的静态完整底稿。

核心价值

适应症拆分和 Mapping 会扩大行数。同一来源记录的多个适应症可能最终映射到同一个 Disease；不同注册或来源记录也可能代表同一业务资产。

本 Skill通过两层合并：

收拢同一来源记录、同一 Disease 的拆分行；

对不同来源记录中的潜在业务等价项进行候选识别和人工关闭。

它只改变记录粒度，不改变任何医学 Mapping 结论。

使用前提

输入必须已经完成：

适应症拆分；

Mapping ID 创建；

疾病 Mapping；

对应疾病 回填；

Mapping 结果的人工复核与关闭。

只处理一个来源池。已上市库与临床库应分别完成本步骤，之后再进行两库合并。

决策单位

第一层：结构性合并

来源记录主键 + Disease CN

用于重建同一原始来源记录中因适应症拆分产生的多行。

第二层：业务等价候选

不同来源记录之间的批准业务身份字段组合

常见身份字段包括：

药品或资产名称；

通用名；

剂型；

商品名；

合并后的适应症；

Disease；

其他经本次任务批准的资产定义字段。

字段组合必须按每次运行冻结，不能直接套用历史 TA。

输入

需要提供：

Mapping 完成后的完整工作簿；

完整来源 Sheet；

固定十列表头的 Mapping结果 Sheet；

来源记录主键；

Mapping ID；

实体、拆分适应症和 对应疾病 字段；

第一层允许变化的字段；

第二层身份字段；

第二层允许汇总的注册或历史字段；

输出 Sheet 名称；

第二层候选决策文件，如存在候选。

Mapping 前置契约

只有同时满足以下条件的记录可以进入合并：

Mapping Status = Mapped；

Disease CN 非空；

来源 Sheet 的 对应疾病 与 Mapping 结果一致。

以下状态永不参与合并，并按原顺序逐行保留：

Manual Review Required - No TA Match

Manual Review Required - Multiple Candidates

Unmapped - Other TA

Unmapped - Invalid Information

本 Skill 不转换旧状态，也不重做医学判断。

两层合并逻辑

第一层：同来源、同 Disease 的结构性合并

对同一个来源主键和 Disease：

按来源顺序拼接 Mapping ID；

按来源顺序拼接拆分适应症；

来源主键和 Disease 保留一次；

其他字段原则上必须一致；

只有明确批准的拆分变化字段允许拼接。

实体名称不放入第一层主键，而作为一致性验证字段。若同一来源主键出现不同实体名称，应暴露冲突并停止，不能通过把实体名称加入主键来掩盖问题。

第二层：跨来源记录的业务等价判断

第一层完成后，再比较不同来源记录。

每个候选必须明确决定：

MERGE

KEEP SEPARATE

并填写有实质内容的 Rationale。

以下差异天然可汇总：

来源主键；

Mapping ID。

其他差异只有被列入本次批准的允许差异字段时，才可以合并，例如：

持证商；

生产企业；

批准文号；

规格；

包装；

批准日期；

其他注册历史字段。

若存在未批准差异，即使人工选择 MERGE，系统也应阻止生成。

不同 Disease 永不进入同一个第二层候选。

标准流程

只读检查来源表和 Mapping 结果。

验证固定 Mapping 结构和五状态体系。

验证 Mapping ID 的集合、唯一性、顺序和回填一致性。

冻结第一层主键、允许变化字段和拼接符。

冻结第二层身份字段和允许差异字段。

生成只读 Preview。

检查第一层所有合并组和字段冲突。

检查第二层候选及其全部差异。

对每个第二层候选关闭 MERGE 或 KEEP SEPARATE。

绑定输入、字段契约、候选集合和决策文件的运行指纹。

获得正式执行确认。

生成新工作簿并重新打开完成 QC。

输出

输出为新的工作簿：

保留所有原始 Sheet 和原始审计链；

新增一个完整字段的合并结果 Sheet；

结果 Sheet 使用与完整来源相同的表头和顺序；

结果值为静态值，不保留公式；

保留可追溯到原始 Mapping结果 的 Mapping ID；

非 Mapped 记录完整保留；

输入文件保持不变。

质量门槛与停止条件

以下情况必须停止：

来源表或 Mapping 结果不完整；

Mapping 结果不是固定十列结构；

出现非批准状态；

Mapping ID 为空、重复、缺失或顺序不一致；

来源回填 Disease 与 Mapping 结果不一致；

Rationale 为空；

Mapped 行缺少来源主键、实体或拆分适应症；

第一层存在未批准的字段差异；

第二层候选没有关闭；

MERGE 候选存在禁止差异；

决策指纹与当前候选不一致；

外部公式未获批准；

输出 Sheet 已存在；

输出将覆盖输入或已有文件；

原始 Sheet、行数核算或重新打开验证失败。

与其他 Skills 的关系

Map Indications to TA Diseases
→ Merge Mapped Disease Rows
→ Merge Listed and Clinical Pools

上游 Mapping 决定 Disease；

本 Skill 只收拢同一来源池内部的记录；

下游负责合并中国已上市池和中国临床池。

本 Skill 不生成 Full/Clean 双视图，也不执行 BI 自有、MOA、VBP 或 TA 领先企业标记。

跨 TA 复用原则

稳定复用：

Mapping ID 唯一连接；

两层合并结构；

第一层结构合并、第二层显式决策；

不跨 Disease 合并；

非 Mapped 原样保留；

输入指纹和候选指纹约束。

每个 TA 或来源池重新确认：

来源主键；

业务身份字段；

剂型、盐型、复方和释放形式等身份边界；

允许聚合的注册历史字段；

第二层候选和最终决定；

行数、候选数和收敛数量。

文档层级

README 用于业务理解和快速接入。正式字段契约、决策结构、脚本用法和验证条件以 SKILL.md 与 references/ 为准。
