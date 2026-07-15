Merge Mapped Disease Rows

将已完成疾病 Mapping 的完整 Excel 底稿，按最终研究与计费单位进行两层合并，同时保留完整审计链和未匹配记录。

What it does

合并同一来源记录中已映射至同一 Disease 的多个适应症；

在核心研究信息一致时，收敛仅因持证商、规格或批文不同而重复的注册记录；

保留所有来源序号、Mapping ID、适应症、Rationale 和注册差异；

输出可独立使用的静态完整底稿。

Output

输出工作簿包含：

原始 Mapping 后底稿：保持不变，供追溯和核验；

合并版拆分底稿：完整字段、静态值，可用于后续分析与计费。

Key safeguards

仅处理 Mapped 且 Disease CN 非空的记录；

不跨不同 Disease 合并；

不重做适应症拆分、医学判断或疾病 Mapping；

Manual Review Required、Unmapped 等非 Mapped 记录均原样保留；

核心字段冲突时暂停并请求人工确认；

不修改或覆盖原始输入文件。

Typical use case

适用于已完成“适应症拆分—疾病 Mapping”的药品或分子池，需要将重复适应症和业务等价的跨注册记录收敛为最终研究、分析或计费单位的场景。
