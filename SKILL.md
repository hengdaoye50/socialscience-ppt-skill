---
name: academic-ppt-skill
description: 面向中文社会科学、学术答辩与研究比赛的语义驱动型 PowerPoint 生成与编辑 Skill。先理解问题意识、文献、研究设计、证据与机制，再选择原生可编辑版式；支持原创生成、原有 PPT 基础上追加/插入、真实图表数据更新、逐页渲染与排版质检。
---

# Academic PPT Skill v0.1.0

## 原则

1. **论证优先于装饰**：先还原研究问题、理论预期、证据、识别策略及限制，不因版式裁剪改变研究者原意。不编造来源、估计值、引用或因果关系。
2. **视觉规则稳定、论证结构可变**：统一颜色、网格和字体层级；根据实质内容调整可编辑图表、表格、机制节点及图形关系。不要把相关关系误画成因果关系。
3. **从 PPTX 到渲染双重交付验收**：生成后执行 inspect、render、audit，并人工逐页复核实际预览。所有素材应有来源与授权信息。

## 执行入口

在含 `academic_ppt/` 目录的仓库根目录运行：

```bash
python -m academic_ppt build examples/demo.json out/demo.pptx
python -m academic_ppt render out/demo.pptx out/render
python -m academic_ppt audit out/demo.pptx --pdf out/render/demo.pdf --spec examples/demo.json --json out/audit.json
```

若对用户既有 `.pptx` 修改：**不要覆盖源文件**。先 inspect + render，明确旧页面的视觉语汇和页面角色。如仅追加本系统生成的页面：`python -m academic_ppt build spec.json new.pptx --base original.pptx --after 3`（在原第 3 页之后插入；默认在末尾追加）。源稿保留另存，输出后的整稿必须逐页渲染复核。

## Agent 工作流

1. **解析意图与约束**：研究主题、受众、讲述时间、现有大纲、不可改页面/内容、引用习惯、16:9、主色。需求已足够时直接做，不反复询问。
2. **先写论证大纲**：每页明确「这一页要让听众理解什么」，而非预先按模板强行分割。区分已知事实、理论判断、假设、证据、示意。
3. **按语义选页**：从 `references/layouts.md` 选择 16 类布局；机制图必须显式指定每条边的 `kind`。需要真实数据图表时提供数列、单位、数据来源；无法确定的数据不得伪造。
4. **写 JSON**：按 `references/schema.md` 和 `examples/demo.json`；保留完整学术措辞，过长时拆页或换布局，拒绝截句、省略号挤框、任意缩标题字号。
5. **构建**：`python -m academic_ppt build ...`；原始 PPTX 另存、不覆盖。插页前对现有文件留备份，并验证新旧同名 `slideN.xml` 内容是否改变。
6. **质量关卡**：inspect → render → audit → 自己逐张看 PNG；同时核对中文字体、注释、图表数据、机制连线、表格、引用与原稿未改页面。`audit=pass` 只是机械检查通过，**不是**视觉和研究质量的充分保证。
7. **交付**：PPTX、预览 PDF（可选）、审计 JSON、源 JSON。明确尚不能实现的操作，而非假称全部自动化。

## 模式边界

- A 原创生成：16 种学术页，全部原生文本/图形/表格/统计区间/PPT 图表可编辑；图片以用户提供文件嵌入。
- B 已有 PPT：在源文件内追加或按页码插入**新生成**的页面；尽量不改变原页面，但**不会自动学习原稿全部复杂样式，也不会对任意既有页面进行无损的全要素重设计**。风格相近需人工核对/调整 `theme`。
- C 原生图表更新：按 `(slide, shape_id)` 定位并替换数列，保持图型和系列数量。图片图表、装饰性色块图不伪装成可更新图表。
- 当前不做：可靠自动识别任意上传稿的完整设计 token、复杂 SmartArt/动画迁移、SVG 全语义编辑、自动文献事实核验、任意图表跨类型原地变更。

## 安全与许可证

原创代码 MIT，不打包、不搬运 GordenPPTSkill 第三方受限 PPT 模板。参考其结构化思路，未复制其具体模板或脚本。研究图片、字体及用户素材遵循原权利人许可。不要默默运行网络更新脚本；升级需审核来源和差异。
