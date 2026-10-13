# 语义布局库（16 类）

所有布局均为 16:9 原生 PPTX。统一坐标系和页眉/页脚（封面、章节、结束页例外）；严禁以装饰形状冒充数据证据。

| type | 学术任务 | 关键字段 | 常见误用 |
|---|---|---|---|
| cover | 题名与署名 | title, subtitle, author | 大段背景堆在封面 |
| section | 论证转场 | number, title, lead | 每张内容页前插扉页 |
| question | 明确研究问题 | context, question, significance | 把政策口号当研究问题 |
| tension | 经验悖论 | expectation, observation, puzzle | 用假对立制造问题 |
| literature | 区分已有解释与缺口 | groups[{heading,claim,authors}], gap | 杜撰作者观点 |
| argument | 集中表达核心论点 | claim, reasons[{title,body}], implication | 论据与结论不分 |
| mechanism | 机制和关系可视化 | nodes[{id,label,detail,layer?}],edges[{source,target,label?,kind}] | 将相关边误写成因果；变量数量不匹配 |
| process | 过程追踪、时序 | steps[{time,title,detail,actor?}], takeaway | 将时间先后自动视作因果 |
| design | 案例/数据/方法/边界 | case, materials, strategy, limitations | 只列方法名称不解释可行性 |
| chart | 数量趋势/对比 | chart_type,categories,series[{name,values}],source | 造数据；百分比与绝对量混用 |
| estimate | 置信区间与参照线 | items[{name,estimate,low,high}],reference,source | 没有标识区间水平；把示意数字当实证 |
| table | 结果对照 | headers,rows,source,note | 一页塞入过多列/行 |
| evidence | 引文或关键材料 | excerpt,source,interpretation | 不注明截取范围或出处 |
| image | 研究地图/田野图片 | image,caption,source | 未经授权或未注明图片性质 |
| conclusion | 回答问题、贡献及局限 | findings,contribution,limitations | 把展望替代研究结论 |
| closing | 研究结语 | takeaway,thanks,contact | 加入与研究无关的口号 |

## 机制图的边语义

`causal` 因果，实线紫；`temporal` 时间，虚线浅紫；`association` 相关，虚线灰；`organizational` 组织/权责，实线深紫。仅针对明确给出的研究判断作图，边不是装饰。自动拓扑分层：最多 5 层、每层 4 个节点；若有循环关系，应补充全部节点的 `layer`，且人工检查阅读方向。节点上限 12 个，复杂时拆成两页而不是减小字号。

## 图表原则

- `chart` 的数据为原生 PowerPoint chart series；柱、条、线、面积和饼可选。不能把设计元素当作真实图表。
- `estimate` 为原生可编辑矢量形状，但不是 Excel 联动图；更改数据要重跑 JSON 构建。指定置信区间的类型/水平到标题、注释或来源中。
- `table` 为原生 PowerPoint 表格，不是截图；超出 7 列或 10 行建议拆页。
- 正文、图表、图片的数字和引文必须人工溯源。演示假数据需要显著注明“合成示例”。

## 生成约束

- 选择一页承载一个主要认知任务，并非每页只能有 3 个元素。材料多时按论证关系拆页，不牺牲完整意义。
- 严肃汇报以白底、深字、单主色为基础，字号通过层级控制，不允许某页异常缩小。
- 视觉 Q/A：图片清晰、细节不与标题抢注意力、结论与证据的位置合理、映射关系可以被听众解释。
- 任何与模板一致性冲突的情形，优先保存研究实质，重新选择或开发布局。
