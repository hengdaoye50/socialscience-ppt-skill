# Academic PPT Skill

**面向中文社会科学的语义驱动式学术 PPTX 创作与编辑工具（v0.1.0）**

以原创、可复用的设计语言代替受限的第三方成套模板。先用规范 JSON 描述论证任务，再生成原生可编辑 PowerPoint。支持将新设计页插入已有 16:9 演示文稿、更新既有 PPT 原生图表，以及逐页渲染审计。

## 当前交付内容

- 16 类原创布局：封面、章节、研究问题、经验张力、文献、论点、机制图、过程追踪、研究设计、图表、区间估计、表格、证据、图片、结论、结束
- 三套配色：南开紫（默认 RGB 126/12/110）、蓝灰、酒红；可改自定义颜色
- 原生 PowerPoint 文本、可编辑形状与连线、图表及表格；提供 PDF 和 PNG 页面预览
- JSON 输入协议、静态结构质检、渲染后的 PDF 文本边界检查
- 插入现有 PPT 的新增页面（不会覆盖用户提供的源文件）

## 安装

Python 3.10+，在仓库根目录安装：

```bash
python -m pip install -r requirements.txt
```

生成 PDF / PNG 还需要系统安装 LibreOffice 命令 `soffice` 和可用的中文字体。建议 Linux 安装 Noto Sans CJK SC；PowerPoint 在目标机器中将按宿主机字体重新布局，请交付前以目标环境检查。

## 一分钟开始

```bash
python -m academic_ppt build examples/demo.json out/demo.pptx
python -m academic_ppt inspect out/demo.pptx --json out/inspect.json
python -m academic_ppt render out/demo.pptx out/render --dpi 130
python -m academic_ppt audit out/demo.pptx --pdf out/render/demo.pdf --spec examples/demo.json --json out/audit.json
```

查看 `out/render/slide-01.png` 等实际图片和渲染 PDF。示范数据与图片都是**合成演示内容**，绝非真实研究结论。

## 把新页面插入已有 PPT

```bash
python -m academic_ppt inspect original.pptx --json out/original.json
python -m academic_ppt render original.pptx out/original-preview
python -m academic_ppt build my-slides.json out/combined.pptx --base original.pptx --after 3
python -m academic_ppt render out/combined.pptx out/combined-preview
```

在第 3 页后插入 JSON 中的新页面；不写 `--after` 则追加到末尾。不会覆盖 `original.pptx`。**注意：插入模式沿用原稿画布和母版环境，新生成页面采用所选 `theme`，并不自动从原稿复杂视觉风格学习全套 token**；成稿必须复核。仅支持标准 16:9、13.333 × 7.5 英寸画布。

## 更新原生图表数据

先 `inspect` 确认含图表的页码和 `shape_id`，并确认原图表 series 数。然后：

```bash
python -m academic_ppt chart-update original.pptx chart_edits.json out/chart-updated.pptx
```

此命令只针对真正的原生 PowerPoint 图表，不支持截图、SmartArt、装饰性色块冒充的伪图表；原有图表类型及 series 数保持不变。

## AI Agent 使用

AI 入口位于 [SKILL.md](SKILL.md)，布局说明位于 [references/layouts.md](references/layouts.md)，协议见 [references/schema.md](references/schema.md)。用 `examples/demo.json` 作为输入参照。阅读研究原文时，先产生“逐页论证任务”，再决定页型，不把论文强行套进固定开题章节。保持数据、引文和关系语义可追溯。

## 验收及诚实边界

`audit` 检查越界、可能的占位符、极小字号、PDF 文本越界、渲染页数、部分出处缺失。**它不能保证没有局部遮挡、没有内容错误或不存在语义失真，必须看每页渲染图并核对学术材料。** 目前没有实现任意旧 PPT 自动风格学习、复杂动画保持迁移、自动文献事实核验、原稿内容页全要素无损重设计；这份 v0.1.0 是可运行的第一版内核，而不是宣称全能的商业成品。

## 独立原创和许可

本项目的原创代码、示例布局及文档采用 MIT。受 [GordenPPTSkill](https://github.com/GordenSun/GordenPPTSkill) 的结构化模板组织、语义槽位与渲染质检理念启发，未复制其第三方受限 PPTX 模板或具体脚本。项目输入的用户图片/字体/材料仍须分别遵守各自许可证。

