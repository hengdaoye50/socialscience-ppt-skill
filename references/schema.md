# JSON 输入协议 v0.1

`examples/demo.json` 为完整 16 类可运行样例。根对象如下：

```json
{
  "theme": "nankai",
  "deck": {"title": "课题", "author": "作者", "affiliation": "单位", "date": "2026-10"},
  "slides": [
    {"type": "cover", "title": "研究报告题目", "subtitle": "副标题"},
    {"type": "question", "title": "研究问题", "context": "现象描述", "question": "要解释的具体问题"}
  ]
}
```

`theme` 接收 `nankai/slate/wine` 三个预设名或对象；对象可覆写 `primary/deep/soft/accent/ink/body/muted/line/white/font_cn/font_en`，颜色采用 6 位 HEX。`image` 相对路径以 JSON 文件所在目录为基准。`source` 为可选但强烈建议的页面资料来源；在 chart、estimate、table、image 和 evidence 中应当提供；audit 对遗漏给警告。`deck` 的元信息进入封面/页脚。

机制示例：

```json
{
  "type":"mechanism", "title":"组织行动机制",
  "nodes":[{"id":"a","label":"数据监测","layer":0},{"id":"b","label":"组织研判","layer":1},{"id":"c","label":"协同行动","layer":2}],
  "edges":[{"source":"a","target":"b","kind":"temporal","label":"输入"},{"source":"b","target":"c","kind":"organizational","label":"部署"}],
  "note":"模型示意：请结合材料审查真实关系"
}
```

原生图表更新 JSON 为数组，例如：

```json
[{"slide":1,"shape_id":3,"categories":["A","B"],"series":[{"name":"人数","values":[11,15]}]}]
```

需要先 inspect 原始 PPTX 获取 shape_id（仅原生图表支持），且更新时不能擅自改变已有的 series 数量。
