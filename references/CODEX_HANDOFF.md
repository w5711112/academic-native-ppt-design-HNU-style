# Codex 78页学术 PPT 全量构建与重构总纲 (CODEX_HANDOFF v5.0 全量学术终极版)

> **当前基准时间**：**2026年9月23日**  
> **目标输出路径**：`D:/path/to/share\整理skill\docs\AI科研流程-HNU规划-20260921\output\科研Agent与自定义Skills-HNU-77页.pptx` (全篇精准收敛至 **78 幻灯片**：77 页正文/结构页 + 1 页官方结语，允许在 75~85 页动态微调)  
> **验证基线**：已通过 41 项门禁自动化审计 (**0 Critical / 0 High**)  
> **底层规范版本**：湖南大学专属学术风格 PPT 设计系统（`academic-native-ppt-design-HNU-style` v5.0）与简体中文人性化改写规范（`renhua` route-b-v3）  
> **核心战略基石**：**全面继承 100页底座（`build_reconstructed.py`）的分体式卡片架构，深度落地 Rules 34~41 体系：学术零人称白描、要点加粗方法、痛点动机先行、零 Markdown 字符泄漏、卡片物理边界强约束、防假性降级同页行距一致、同类 Skill 绝对连续、色彩调和、充实分点、交替强对比有机宽方差字阶梯队、行高齐平与大模型演进全景！**

---

## 零、全篇排版与语言质量铁律体系 (Rules 34~41 / QG-34~QG-41)

针对学术组会高标准汇报要求，全篇文稿在排版几何与语言表达上必须同时严格遵守以下 8 项核心铁律：

### 1. 语义色彩调和律、2~4 色配比与湖大红纯白高对比度铁律 (Rule 34 / QG-34)
- 彻底废除按列模运算循环（`[BLUE, GREEN, ROSE, WARM]`）的“彩虹马戏团”配色。
- **单页主题色严格 $\le 4$ 种**（不含纯白背景与深炭灰正文），推荐 2~3 种语义协调色。
- **严格语义映射**：
  - **底色湖大红 (`#A6232B` / `#F8EAEB`)**：各级标题、流程粗箭头、章节高亮、核心痛点对比；
  - **学术深蓝/板岩灰 (`#1D4999` / `#EAF1F6`)**：技术底座、通用流水线步骤、基础组件平台；
  - **鼠尾草灰绿 (`#0F766E` / `#EEF3E7`)**：工程规则、通过验证的最终成效、沉淀复用的 Skill 资产；
  - **暖赭石/琥珀黄 (`#B45309` / `#FEF3C7`)**：**严格专用于负面/警示语义**（算力成本超标、踩坑记录、版本漂移风险），严禁滥用于中性模型或选型卡片。
- **湖大红深色背景文字高对比度纯白铁律 (High Contrast Pure White Text on HNU Red)**：
  - 凡以**湖大红 (`#A6232B` / `#8B1D23` / 深红)**作为底托/填充背景的文本框、卡片或强调色块，**其内部承载的标题、正文、分点及标注文字必须统一强制使用纯白色 (`#FFFFFF` / `RGBColor(255, 255, 255)` / `schemeClr val="bg1"`)**；
  - **严禁在湖大红深色背景上继续沿用深炭灰 (`#242424`)、黑色或暗红文字**（如 Slide 18 曾出现黑字配红底的反面教训），彻底杜绝文字隐形与低可读性，保障高反差对比与清晰度。

### 2. 核心分点自适应展开与长句主动拆解律 (Rule 35 / QG-35)
- **弹性分点机制 (2 ~ 5 条)**：彻底打破“固定 2 条分点”的机械教条！分点数量完全由卡片可用尺寸与信息深度动态决定：常规精炼卡片 2~3 条，高密技术卡片 3~5 条；
- **长难句主动拆解**：严禁在单个分点内堆砌长难复合句；长句必须拆解为多个独立分点，每点聚焦一个核心机制或参数维度；
- **客观事实扩充赋权**：在描述通用技术、开源大模型架构与科研流水线等**客观事实**时，明确允许且倡导 Codex 结合领域常识进行 2~4 条充实的客观技术事实分点扩充。

### 3. 有机饱满度与真实交替宽方差律 (Rule 36 / QG-36)
- **垂直饱满度判定区间：严格 [60% ~ 95%]**（字数覆盖率 60%~95%，底线 60%，消除空洞死白；上限 95%，防止文字溢出；物理溢出阈值 108%）。
- **真实交替宽方差 (Alternating High-Contrast Variance)**：坚决杜绝“全篇卡片字数完全相同、齐刷刷停在同一列左侧”的机械僵硬感！
  - **强对比与微对比有机交替**：并排卡片相邻极差不强制要求处处巨大，但整体必须“时大时小”——约一半比例相邻极差 $\ge 25\%$（如 94% 紧邻 63%），另一半比例保持 $5\%\sim 25\%$ 的平滑微差（如 84% 紧邻 75%），交替呈现自然起伏的人工编写与排版呼吸感；
  - **单页两极起伏**：单页内各卡片填充率极差 $\ge 20\%$，向两极倾斜分布。
- **行间距自适应与防假性误降级铁律 (Anti-False-Degradation & Line Spacing Rules)**：
  - **默认黄金基准区间**：默认正文行间距必须保持在 **18 ~ 20 pt**（基准推荐 **18.5 pt**）；
  - **父级卡片净高防误降级算法 (Strict Anti-False-Degradation)**：饱满度与文字密度测算必须以**底层背景卡片的可用正文净高（$H_{\text{avail}} = H_{\text{card}} - 41.0\text{ pt}$）**为基准，绝对禁止以紧缩包裹后的临时文本框高度计算密度！
    - *典型案例*：Slide 34 顶部卡片高达 $h=241$ pt，可用正文净高 $200$ pt，9 行文字占高 $166.5$ pt（实际密度仅 $0.833 \le 0.95$），此前因临时文本框包裹为 $188$ pt 导致误判为 $185/188 = 0.984 > 0.95$ 从而被错误降级为 17.0 pt；现已彻底规正，必须稳固锁定 18.5 pt 标准行距；
  - **真实极端情况弹性微调**：仅当且仅当卡片在父级净高下密度确实 $> 0.95$ 且内容无法进一步精炼时，才允许单页统一微调为 **17.0 pt**（严禁低于 16.9 pt 触发门禁）；若卡片框宽松开阔，允许微调至 **19.5 ~ 21.0 pt**；
  - **同页严格一致性铁律 (Same-Page Line Spacing Variance = 0.0 pt)**：无论该页最终采用 17.0 pt、18.5 pt 还是 19.5 pt，**同一页内所有正文文本框的行间距必须保持完全一致**（方差为 0.0 pt），彻底杜绝同页文本框忽紧忽松。

### 4. 混排多栏高度齐平与卡片物理边界强约束律 (Rule 37 / QG-37, QG-20/QG-30)
- **同行卡片齐平**：同一行并排卡片高度必须严格物理齐平（$\Delta h = 0$）。
- **底层卡片物理边界强约束 (Card Boundary Physical Confinement Invariant, QG-20/QG-30)**：
  - 正文文本框的物理下边缘必须严格受限于底层卡片下边界（`shape.top + shape.height <= parent_bg.top + parent_bg.height - 6.0 pt`），严禁超出卡片底边穿透至下方卡片（如 Slide 6 曾出现顶部里程碑文字下渗 26 pt 穿透到底部卡片的重大缺陷，已被门禁严密封堵）；
- **140 pt 里程碑卡片文字容量预算 (Milestone Card Budget)**：
  - 高度 $h=140$ pt 的里程碑卡片（如 Slide 5/6 顶部卡片），可用正文净高仅 95 pt，文字容量预算上限严格为 **4 行**（2 个分点，每点折行后严格 $\le 2$ 行，中文字符建议 $\le 22$ 字/点），严禁产生 3 行长分点，确保总渲染高度 $\le 82$ pt，留出 $\ge 13$ pt 充裕安全边距。
- **非对称水印协调**：官方水印（如逸夫楼手绘 `图片 9`）自然坐落在左下角开阔处，右侧卡片统一高度（如 $h=240\text{ pt}$），消除压在水印头顶的丑陋小残块。

### 5. 跨周期演进全景展开与选型防线 (Rule 38 / QG-38)
- 彻底根治模型演进缺失痛点，不再将 2024~2026 庞大的模型演进硬塞在单页。拆分为 **3 页连贯纵深架构**：
  - **Slide 5：自主与开源大模型演进**（涵盖 DeepSeek V2/V2.5、V3/R1、V4 Flash/Pro、V4.1 Flash 全系，以及智谱 GLM-4/AllTools 至 GLM-5/5.1/5.2/5.3/5.3 Flash 全系，底部配备私有化部署与长链推理实测支撑卡）；
  - **Slide 6：前沿闭源与多模态演进**（涵盖 OpenAI GPT-4/4.1 至 GPT-5/5.1~5.5、5.6 Sol/Terra/Luna 及 GPT-6 Astra，Google Gemini 1.5/2.5 至 3.1 Pro、3.6/3.7/3.8 Flash，小米 MiMo v2 至 v2.5 Pro/Flash 及 2026年9月发布的 v2.6 Pro，底部配备商业多模态理解与多模型梯队调度支撑卡）；
  - **Slide 7：科研场景多模型选型防线与调度策略**（5卡矩阵：版本口径锁定、脱离合成跑分陷阱、慢思考推理边界、多模态图表防线、分级混合调度）。

### 6. renhua 人性化语言风格、学术零人称与要点加粗铁律 (Rule 39 / QG-39 - 重中之重)
- **学术叙述零人称规范 (Zero Pronoun Invariant)**：
  - 学术汇报与技术文稿中**严禁出现第一/第二人称主语（如“我”、“你”、“他”、“我们”、“大家”）**（如严禁写“我希望减少反复拖动文本框的时间”）；
  - 一律改用客观、平实、目的导向的工程白描表达（如“**为了减少反复拖动文本框的时间**：……”、“**针对文献漏检实际问题**：……”、“**为打通三端联动流程**：……”）；
- **分点开头核心论点/动机强制加粗 (Bullet-Lead Bold Rule)**：
  - **听众方法原则**：分点开头必须将根本原因、技术痛点或关键动作显式加粗（如 `**根本动机**：...`、`**手工排版瓶颈**：...`、`**断点回溯机制**：...`）；杜绝没有视觉重心的平淡平铺叙述，让观众一眼抓住核心信息！
- **大厂黑话与伪架构套话清洗对照表 (Buzzwords Purge Invariant)**：
  - 必须能读懂！杜绝夸大描述与晦涩词汇堆砌，严格执行如下替换：
    | 禁用词 / 伪架构套话 | 严禁原因 | 规范学术与工程白描替换 |
    | :--- | :--- | :--- |
    | **壁垒** | 大厂黑话、夸大叙事 | **困难 / 阻碍 / 技术门槛 / 实际限制** |
    | **(固化为)可复用资产** | 大厂黑话、虚词堆砌 | **具体工程规则 / 操作规范 / 可复用规则与方法** |
    | **原子化** | 互联网黑话、抽象难懂 | **单项基础操作 / 细分步骤 / 独立功能模块** |
    | **沉淀为(工程规则/资产)** | 大厂黑话、空泛套话 | **整理成规范 / 写成规则 / 固化为稳定方法** |
    | **高频科研方法** | 伪架构词、不符语境 | **常用的科研方法** |
    | **标准化通信与状态感知** | 晦涩空洞套话 | **把各类外部工具连接进来，让程序随时调用并了解运行状态** |
    | **自愈闭环** | 伪架构套话、空泛口号 | **排查并自动修正问题 / 完整执行与验证流程** |
    | **旨在打造 / 飞速发展** | 官话套话 | **直接删除，陈述具体技术机制与实验事实** |
    | **具有划时代/里程碑意义** | 浮夸修饰 | **直接删除，用具体评测指标与性能对比说话** |
    | **不仅……而且……** | 空洞连词 | **拆分为两个独立客观的技术事实分句** |
    | **显著提升 / 显著优势** | 虚饰主观词 | **用具体数字、收敛轮数或真实实验对比呈现** |
    | **极大地帮助** | AI客服腔 | **支持 / 辅助 / 帮助 / 加速** |
- **客观事实与术语守恒 (Strict Factual Grounding)**：
  - 严禁为了“去 AI 味”而删减真实技术细节或篡改事实；
  - 领域专业术语（如 DeepSeek MLA、MoE、MTP、R1 CoT、GLM-4 AllTools、MiMo 2.6 Pro、MCP、Skill 等）必须精准保留；
- **标点克制与排版洁净 (Clean Punctuation & Layout)**：
  - 严禁破折号（——），严禁分号（；），冒号克制，不用人工字符拼凑项目符号（如 `"• "` 或 `"- "`），统一依托原生 XML 悬挂缩进；中英文与数字之间保留标准空格。

### 7. Skill 叙事逻辑三部曲、痛点先行与大纲聚合铁律 (Rule 40 / QG-40)
- **痛点与动机先行原则 (Pain Point & Motivation First)**：
  - 学术逻辑因果必须严密，**严禁一上来就悬空讲解 Skill 的代码或构建过程**！
  - 每个要讲述到的 Skill，必须拥有专门的一段、一页或独立的显式区域，先讲透**“为什么做”、“解决什么痛点”、“原始问题是什么”**，再讲“如何构建与运作”，最后讲“沉淀效果”；
- **相同 Skill 页面绝对连续律 (Contiguous Skill Pages)**：
  - 同一 Skill 的多张页面必须在结构上**紧密相连、一气呵成**，严禁跨 Skill 穿插割裂（如严禁在 `read-paper-analysis` 页面中间穿插 `obsidian-note-style` 页面，必须将 read-paper 全部连续讲完后，再进入 obsidian-note-style 体系！）；
- **需求驱动核心哲学 (Core Grounding Philosophy)**：
  - 必须在文稿显要位置向听众明确：**无论哪一个 Skill，其核心都是根据研究者的真实需求建立的；只有需求描述得足够清楚，AI 才能发挥最大功效；AI 是强力辅助，但绝无法替代研究者的学术判断与主导地位**；
- **篇幅弹性与动态伸缩律 (Dynamic Page Elasticity)**：
  - 演示文稿总页数根据内容逻辑饱满度动态自适应（基准在 75 ~ 85 页区间内弹性浮动），严禁为了死凑某个固定数字而导致页面排版过挤、过窄或大面积留白。

### 8. 零 Markdown 字符泄漏与原生 OpenXML 富文本渲染契约 (Rule 41 / QG-41 - 核心红线)
- **零 Markdown 符号泄漏 (Zero-Markdown-Leak Invariant)**：
  - 彻底杜绝在 PPTX 字符串或分点中直接拼接输出 Markdown 语法标记字符（如 `**加粗**`、`__强调__`、`## 标题`、`` `代码` ``）；
  - PPTX 不是浏览器或 Markdown 解析器，DrawingML 的 `<a:t>` 文本标签会将 `**` 视为普通可见字符原样打印在幻灯片上（如 Slide 9 曾出现 `”**外部工具连接**“` 的重大缺陷，严重损害学术汇报的严肃性与专业度，属于严厉禁止的大忌）；
- **原生 OpenXML Run 样式帮助 (Native DrawingML Run Bolding)**：
  - 所有分点要点加粗，必须在 Python-pptx / OpenXML 中通过独立的 text run 与 `run.font.bold = True` 原生渲染；
  - 文本流必须在写入 PPT 前经过强制清洗（如 `re.sub(r'\*\*([^*]+)\*\*', r'\1', text).replace('**', '').replace('__', '')`），确保流入 `<a:t>` 的字符 100% 洁净。

---

## 一、全系核心 Skill 构建根源与真实动机对齐表 (Rule 40 执行基准)

在编写各 Skill 章节时，Codex 必须将以下真实动机作为各页面的第一逻辑入口：

| 模块 / Skill 名称 | 核心定位 | 原始痛点与“为什么构建”(Motivation First) | 固化规则与沉淀成效 |
| :--- | :--- | :--- | :--- |
| **`academic-native-ppt-design-HNU-style`** | 湖大学术原生 PPT 设计系统 | **针对组会与学术汇报美观度不足的实际痛点**：官方原生模板版式不够统一、边框线条生硬不够美观，且每次手工排版极度耗时、成果质量不稳定。 | 提炼出固定美观逻辑与零损坏规范，以官方母版为模具，实现 DrawingML 原生排版、字阶梯队与字框解耦，一键生成免微调的高质量汇报胶片。 |
| **`draw-style`** | 学术论文级绘图风格规约 | **针对 AI 绘图“一眼假”、塑料感重、不学术且改图效率低下的痛点**：每次出图提示词重新构思浪费大量时间，生成的图片风格漂移且无法直接放入顶会论文。 | 沉淀顶会论文图纸级配色体系（学术蓝、陶粉红、鼠尾草绿）与几何拓扑规约，实现高质量、固定形式的一键出图与跨面板风格一致。 |
| **`renhua`** | 中文人性化白描与去 AI 味 | **针对现有 GitHub 高星降 AI 味工具生硬套用、不符合中文自然科研表达习惯的痛点**：机器改写后语句怪异，AI 腔套话依然连篇。 | 自建工科白描求真规范，消除机械套话与公文黑话，坚持学术零人称与客观陈述，让表达回归务实求真。 |
| **`drone-literature-scout`** | 前沿文献定向追踪与探索 | **针对选题初期不知道研究方向、海量文献不知从何下手的迷茫痛点**：面对海量前沿文献无从切入，人肉检索耗时费力。 | 借助 AI 定时自动化爬取、过滤与调研前沿方向，作为首个自建探索型 Skill，建立候选论文池并输出结构化可核验清单。 |
| **`zotero-obsidian-paper-import`** | 文献自动入库与双向链接 | **针对人肉逐篇搜索下载论文极度缓慢、且极易漏检的痛点**：文献与阅读笔记割裂，无法追溯依据。机器在限定边界下批量检索速度远快于人工。 | 打通 Codex-Zotero-Obsidian 三端联动，在来源准确前提下自动入库 PDF 原文，并在笔记知识点中直链跳转 PDF 原文对应页面。 |
| **`read-paper-analysis-highlight`** | 论文深度精读与原生高亮 | **针对长篇论文精读效率低下、关键机理容易遗忘的痛点**：普通摘要无法深入数学推导与实验细节。 | 快速提炼核心创新与数学机理，三轮精读分解输入输出，并按照个人习惯对论文 PDF 施加原生高亮与批注，便于快速复盘。 |
| **`obsidian-note-style`** | 体系化科研笔记智能构建 | **针对不想再手写/手敲笔记（速度太慢）的痛点**：人工整理笔记耗时漫长，知识点孤立无法网状互联。 | 让 AI 扫描已有大量笔记素材，深度学习研究者个人风格与排版逻辑，自动生成包含折叠块、知识双链与彩色标注的体系化笔记。 |
| **治理类 Skills** (`contract-lock`, `governor`, `collect-bug`, `workspace-hygiene`) | 多 Skill 协作与工程环境自愈 | **针对多 Skill 协作下规则冲突打架、依赖漂移、Bug 重复踩坑与工程环境污染的痛点**。 | 建立契约锁定防止篡改，环境卫生定期巡检，错误按问题族分类复用沉淀，形成自愈闭环。 |

> [!IMPORTANT]
> **全篇核心哲学基石 (Core Grounding Philosophy)**：
> 无论哪一个 Skill，其核心都是**根据研究者的真实需求建立的**；**只有需求描述得足够清楚，AI 才能发挥最大功效**；AI 是强力辅助，但**绝无法替代研究者的学术判断与主导地位**！

---

## 二、Codex 必备资产与文件清单

Codex 执行时聚焦以下核心文件：

| 序号 | 资产类别 | 绝对路径 | 核心作用说明 |
| :---: | :--- | :--- | :--- |
| **1** | **执行总纲** (必读) | `D:/path/to/share\整理skill\docs\AI科研流程-HNU规划-20260921\CODEX_HANDOFF.md` | **本指南 (v5.0 全量学术终极版)**，定义全域规范、痛点动机、版型代码、大纲映射与验证指令。 |
| **2** | **全量生成代码** (唯一工作台) | `D:/path/to/share\整理skill\docs\AI科研流程-HNU规划-20260921\.skill-contract\rebuild-v4\build_77page_deck.py` | **包含前 12 页高精校准与 78 页流水线的生成脚本**，包含 `fixed_card`、`merged_matrix`、`open_source_model_page` 等成熟函数。 |
| **3** | **100页成熟母本** (完整参考) | `D:/path/to/share\整理skill\docs\AI科研流程-HNU规划-20260921\.skill-contract\rebuild-v4\build_reconstructed.py` | 100 页成熟版本的完整内容、分体架构与原始文字底本。 |
| **4** | **内容交接文档** (语义依据) | `D:/path/to/share\整理skill\docs\AI科研流程-HNU规划-20260921\当前PPT构建思路与完整内容交接.md` | 全书 5 大篇章递进脉络、原始分点与真实技术论据。 |
| **5** | **模板底板** (母版) | `C:/path/to/skills/academic-native-ppt-design-HNU-style\templates\hnu_base_template.pptx` | 官方母版，尤其参考第 21 页。 |
| **6** | **门禁审计器** (验收) | `C:/path/to/skills/academic-native-ppt-design-HNU-style\scripts\audit_hnu_deck.py` | 运行检测 40 项门禁（包含 QG-34~QG-40 自动化检测）。 |
| **7** | **真实素材图库** (装配) | • `D:/path/to/slide-assets\` (35张主素材 + Benchmark + `展示自动维护整个体系.png`)<br>• `C:/path/to/paper-reference-figure\` (3张学术基准图) | **39 张真实图源**，100% 完整收录，严禁裁切，单张复用不超过 2 次。<br>★ **特别提醒**：`D:/path/to/slide-assets\展示自动维护整个体系.png` (1132x850) 必须装配在 **Slide 71: 【skill-ecosystem-governor：维护来源与链接】**，采用大图展台与底部双卡片协同！ |

---

## 三、卡片容量、字数标定与防溢出工程参考表

为了彻底消除文字溢出（如 `Body need > body_h` 报错）并满足 QG-36（利用率 60%~95% 强对比交替）、QG-39（学术零人称白描 + 要点加粗）与 QG-40（动机先行），Codex 必须按以下标定参数编写各卡片分点：

| 版型卡片类型 | 卡片外框高 $h$ | 正文高度 `body_h` | 正文字号 | 推荐分点数 | 换行行数上限 | 推荐字符数区间 | 估算利用率 |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **矩阵高密卡片 (上方)** | 190 pt | 135 pt | 14.0 pt | 2~3 条 | 6 ~ 7 行 | 85 ~ 100 字符 | 88% ~ 95% (高密) |
| **矩阵轻量卡片 (上方)** | 190 pt | 135 pt | 14.0 pt | 2 条 | 4 ~ 5 行 | 60 ~ 72 字符 | 62% ~ 72% (透气) |
| **矩阵卡片 (下方)** | 185 pt | 130 pt | 14.0 pt | 2 条 | 5 ~ 6 行 | 75 ~ 88 字符 | 75% ~ 85% (稳态) |
| **流程/里程碑卡片** | 140~150 pt | 95 pt | 13.0 pt | 2 条 | 恰好 4 行 | 45 ~ 55 字符 | 75% ~ 83% |
| **底部大支撑卡片** | 220~235 pt | 190 pt | 13.5 pt | 3~4 条 | 8 ~ 9 行 | 130 ~ 160 字符| 82% ~ 92% |
| **流程底部支撑卡片** | 131 pt | 90 pt | 13.5 pt | 2 条 | 恰好 4 行 | 60 ~ 75 字符 | 78% ~ 85% |

### 真实交替宽方差编排范式 (Alternating High-Contrast Pattern)
在 5 卡矩阵页中，严禁所有卡片字数整齐划一！推荐采用如下真实强对比交替序列：
- **卡片 1 (左上·核心痛点)**：高密详述型，3 分点，95% 饱满度（字数 95 字）；
- **卡片 2 (中上·直接动机)**：精炼透气型，2 分点，63% 饱满度（字数 62 字）；
- **卡片 3 (右上·工程方法)**：中密结构型，2 分点，84% 饱满度（字数 82 字）；
- **卡片 4 (左下·验证指标)**：精炼透气型，2 分点，70% 饱满度（字数 68 字）；
- **卡片 5 (右下·沉淀成效)**：高密总结型，3 分点，92% 饱满度（字数 92 字）。
> **整页饱满度极差达 32%**，相邻卡片极差达 32% 与 21%，视觉节奏自然起伏，彻底摆脱机械死板感！

---

## 四、核心版型标准实现代码 (Copy-Paste Ready)

Codex 在扩展全部 78 页时，请严格复用以下标准函数：

### 1. 分体式卡片几何结构与要点加粗解析 (`fixed_card` & `fixed_tx`)
```python
def fixed_tx(slide, x, y, w, h, value, size=14.0, bold=False, color=None, name="Body", bullet=None):
    values = value if isinstance(value, list) else [value]
    bullet = isinstance(value, list) if bullet is None else bullet
    color = base.b.INK if color is None else color
    if not bold:
        for candidate in [size, size - 0.5, size - 1.0, 13.0]:
            if candidate >= 13 and base.height(values, w, candidate, bullet) <= h:
                size = candidate
                break
    
    # 过滤掉 markdown 加粗标记后计算换行与行数
    clean_values = [re.sub(r'\*\*(.*?)\*\*', r'\1', text) for text in values]
    lines = [base.b.wrap(text, w - 20 if bullet else w - 2, size) for text in clean_values]
    need = sum(len(line) * size * 1.25 + (10 if bullet else 0) for line in lines)
    if need > h + 1:
        raise ValueError(f"Overflow {slide.page_no} {name} {need:.1f}>{h}, {values}")
    
    shape = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    shape.name = name
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = False
    frame.auto_size = MSO_AUTO_SIZE.NONE
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = Pt(0)
    
    for index, (raw_val, wrapped) in enumerate(zip(values, lines)):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.line_spacing = Pt(max(18.0, size * 1.25))  # Rule 36: 行间距 18~20 pt
        if bullet:
            base.apply_rich_bullet(paragraph, " ".join(wrapped), size, base.b.rgb(color), True)
        
        # Rule 39: 分点开头核心论点/动机强制加粗 (Bullet-Lead Bold)
        p_raw = "\v".join(wrapped)
        m_bold = re.match(r'^\*\*(.*?)\*\*[:：](.*)$', raw_val.strip(), re.DOTALL)
        if m_bold and bullet and not bold:
            lead_txt = m_bold.group(1).strip()
            rest_txt = m_bold.group(2).strip()
            # 重新根据换行切分
            lead_with_colon = f"{lead_txt}："
            paragraph.text = ""  # 清空后分 run 注入
            run1 = paragraph.add_run()
            run1.text = lead_with_colon
            base.apply_yahei(run1, size, True, base.b.rgb(base.b.RED if color == base.b.RED else base.b.INK))
            run2 = paragraph.add_run()
            # 将 rest_txt 格式化为 wrapped 中剩余的部分
            clean_raw_wrapped = "\v".join(wrapped)
            if clean_raw_wrapped.startswith(lead_with_colon):
                run2.text = clean_raw_wrapped[len(lead_with_colon):]
            else:
                run2.text = rest_txt
            base.apply_yahei(run2, size, False, base.b.rgb(color))
        else:
            paragraph.text = "\v".join(wrapped)
            for run in paragraph.runs:
                base.apply_yahei(run, size, bold, base.b.rgb(color))
                
    frame.vertical_anchor = MSO_ANCHOR.TOP
    base.metrics.append({"page": slide.page_no, "name": name, "x": x, "y": y, "w": w, "h": h, "font": size, "text": values, "estimated_height": need})
    return shape


def fixed_card(slide, x, y, w, h, title, body, size=14.0, fill=None, title_size=18.0, body_h=None):
    """
    分体式卡片解耦架构：
    1. 底托几何矩形 (base.b.rect)：锁死卡片物理外框与行高，消除形状与文字边距耦合；
    2. 独立标题框 (fixed_tx)：置顶 y+8，固定高度 28pt，加粗红字（红底时为纯白），绝对不与正文争抢行距；
    3. 独立正文字框 (fixed_tx)：置于 y+35，高 h-41 (或显式指定 body_h)，自适应字号阶梯 (14.0 -> 13.5 -> 13.0 pt)，防止文字溢出；
    4. 湖大红等深色底文字纯白铁律：当底色为湖大红 (RED / 深红) 时，标题与正文颜色自动切换为纯白 (WHITE)！
    """
    fill = base.b.WARM if fill is None else fill
    base.b.rect(slide, x, y, w, h, fill)
    is_dark = (fill in (base.b.RED, getattr(base.b, 'DARK_RED', None))) or (hasattr(fill, 'rgb') and str(fill.rgb).upper() in ('A6232B', '8B1D23', 'C00000'))
    title_clr = base.b.WHITE if is_dark else base.b.RED
    body_clr = base.b.WHITE if is_dark else base.b.INK
    fixed_tx(slide, x + 10, y + 8, w - 20, 28, title, title_size, True, title_clr, "Section", False)
    fixed_tx(slide, x + 10, y + 35, w - 20, h - 41 if body_h is None else body_h, body, size, False, body_clr, "Body", True)
```

### 2. 五框合并卡片矩阵 (`merged_matrix`)
用于将原本松散的 2~3 页纯文字内容合并为高密、饱满的 1 页（如 Slide 3、Slide 7、Slide 8、Slide 11、Slide 12 等）：
```python
def merged_matrix(slide, item: dict) -> None:
    for index, card_data in enumerate(item["cards"]):
        if index < 3:
            fixed_card(
                slide,
                43 + index * 294,
                104,
                282,
                190,
                card_data["title"],
                card_data["points"],
                14.0,
                [base.b.BLUE, base.b.ROSE, base.b.ROSE][index],
                body_h=135,
            )
        else:
            fixed_card(
                slide,
                296 + (index - 3) * 317,
                309,
                305,
                185,
                card_data["title"],
                card_data["points"],
                14.0,
                [base.b.GREEN, base.b.BLUE][index - 3],
                body_h=130,
            )
```

### 3. 横向流程拓扑页 (`agent_flow_page` / `mcp_flow_page`)
用于展示多阶段流水线、从问答到执行等流程页（如 Slide 4、Slide 9）：
```python
def agent_flow_page(slide, item: dict) -> None:
    base.b.rect(slide, 43, 104, 875, 70, base.b.BLUE)
    base.tx(slide, 55, 112, 180, 22, "从问答到执行", 16, bold=True, color=base.b.INK, name="Section", bullet=False)
    base.tx(slide, 232, 112, 672, 45, "单次问答只提供孤立建议，完整科研涵盖环境感知、工具调度、多轮校验与产物交付。Agent 的价值在于把分散动作连成具有长期状态记忆、可恢复的持续推进过程。", 13.5, name="Body", bullet=False)
    # 5 步里程碑
    stages = [...]
    xs = [43, 221, 399, 577, 755]
    for index, ((title, points), x) in enumerate(zip(stages, xs)):
        fixed_card(slide, x, 189, 163, 150, title, points, 13, fills[index], 16, body_h=95)
        if index < len(stages) - 1:
            base.engine._add_flowchart_arrow(slide, Pt(x + 160), Pt(254), Pt(20), Pt(14), color=base.b.rgb(base.b.RED))
    # 底部 2 卡片 (高度严格齐平，Delta h = 0！)
    fixed_card(slide, 296, 363, 305, 131, "完成证据与溯源", [...], 13.5, base.b.GREEN, body_h=90)
    fixed_card(slide, 613, 363, 305, 131, "研究者责任边界", [...], 13.5, base.b.BLUE, body_h=90)
```

### 4. 单图解耦展台与水印避让 (`image_page` 特化)
```python
def image_page(slide, item: dict) -> None:
    first = item["images"][0]
    if first == "I02" or item.get("old_page") in (13, 14):
        # 解耦左侧卡片与右侧大图，左下角 y=355 留空避让逸夫楼水印
        fixed_card(slide, 43, 104, 240, 240, "能力分层与模块封装", [...], 13.5, base.b.WARM, 16.0, body_h=190)
        base.pic(slide, first, 296, 100, 622, 390, crop=None) # 零裁切等比居中
        base.tx(slide, 296, 493, 622, 17, "课题组科研 Agent 与自定义 Skill 总体组织架构图", 13, color=base.b.MUTED, name="Caption", bullet=False)
        return
```

---

## 五、全篇 78 幻灯片完整构建大纲 (1 to 78 - 相同 Skill 绝对连续架构)

> [!IMPORTANT]
> **大纲结构彻底重构优化 (Rule 40)**：
> 原版本中第 57 页与 64 页的 `obsidian-note-style` 穿插在 `read-paper-analysis-highlight` 内部，造成严重叙事割裂。
> 本大纲已将第四篇进行彻底梳理：**先全量连续呈现 `read-paper-analysis-highlight` 的完整闭环（Slides 49~59），紧接着全量连续呈现 `obsidian-note-style` 的体系化构建（Slides 60~65）**，完全杜绝交叉割裂！

| 序号 | 原 100 页对应 | 章节 | 幻灯片标题 | 版型类型 | 视觉元素 / 流程 / 水印 |
| :---: | :---: | :---: | :--- | :--- | :--- |
| **1** | 1 | 封面 | 用 Agent 搭建科研流程 | Cover | 岳麓山实拍底图 + 保护底框 |
| **2** | 2 | 目录 | 汇报提纲 | Agenda | 湖大五大篇章标准提纲 |
| **3** | 3, 4 | 第一篇 | 科研主线与重复劳动的真实代价 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **4** | 5, 6 | 第一篇 | Agent 从回答问题走向持续执行 | `agent_flow` (5步+2卡) | 流程粗箭头 + 水印 |
| **5** | 7 (拆) | 第一篇 | 2024 至 2026 年自主与开源模型演进 | `open_source_model` (4节点+2卡) | DeepSeek/GLM 谱系时间线 |
| **6** | 8 (拆) | 第一篇 | 2024 至 2026 年前沿闭源与多模态演进 | `frontier_model` (4节点+2卡) | GPT/Gemini/MiMo 谱系时间线 |
| **7** | 9 (拆) | 第一篇 | 科研场景下的多模型选型防线与调度策略 | `model_selection` (5卡) | 逸夫楼水印 |
| **8** | 10, 11| 第一篇 | 把科研需求写成可执行任务 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **9** | 12, 13| 第一篇 | MCP 连接工具，Skill 保存方法 | `mcp_flow` (5步+2卡) | 流程粗箭头 + 水印 |
| **10**| 14 | 第一篇 | Plugin 把相关 Skills 组织起来 | `image_page` (大图展台) | `图片 1.png` / `I02` 零裁切全景 |
| **11**| 15, 16| 第一篇 | Codex 执行操作，研究者负责判断 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **12**| 17, 18| 第一篇 | 从个人需求到可复用工程规则 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **13**| 18 | 第一篇 | 模型会更新，需求与方法仍可复用 | `content` (基础单页) | 湖大蓝卡片 |
| **14**| 19 | 第一篇 | 按科研职责认识整套体系 | `content` (基础单页) | 湖大蓝卡片 |
| **15**| 20 | 过渡页 | 第二篇：图文表达与可复用规则 | Agenda Transition | 第二章高亮提纲 |
| **16**| 20 | 第二篇 | 【academic-native-ppt-design-HNU-style：构建动机与痛点】 | `image_page` | `图片 2.png` 零裁切全景 (动机先行) |
| **17**| 21 | 第二篇 | 【academic-native-ppt-design-HNU-style：格式迭代】 | `image_page` | `图片 3.png` 零裁切全景 |
| **18**| 22 | 第二篇 | 【academic-native-ppt-design-HNU-style：学术配色】 | `image_page` | `图片 4.png` 零裁切全景 |
| **19**| 25, 26| 第二篇 | 【academic-native-ppt-design-HNU-style：版式与字阶规则】 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **20**| 25 | 第二篇 | 【academic-native-ppt-design-HNU-style：质量核对】 | `image_page` | `图片 5.png` 零裁切全景 |
| **21**| 28, 29| 第二篇 | 【draw-style：构建动机与学术配色】 | `merged_matrix` (5卡) | 逸夫楼水印 (动机先行) |
| **22**| 30, 31| 第二篇 | 【draw-style：几何图的结构与配色】 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **23**| 32, 33| 第二篇 | 【draw-style：方法图动线与拓扑】 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **24**| 32 | 第二篇 | 【draw-style：数据图的层次】 | `image_page` | `图片 6.png` 零裁切全景 |
| **25**| 35, 36| 第二篇 | 【draw-style：图型选择与对照规范】 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **26**| 35 | 第二篇 | 【draw-style：动作与联合评价】 | `image_page` | `图片 7.png` 零裁切全景 |
| **27**| 36 | 第二篇 | 【draw-style：安全走廊与表达】 | `image_page` | `图片 8.png` 零裁切全景 |
| **28**| 37 | 第二篇 | 【draw-style：速度障碍的几何】 | `image_page` | `图片 9.png` 零裁切全景 |
| **29**| - | 第二篇 | 【draw-style：动态规划的状态复用】 | `image_page` | `图片 10.png` 零裁切全景 |
| **30**| - | 第二篇 | 【draw-style：观测到安全速度】 | `image_page` | `图片 11.png` 零裁切全景 |
| **31**| - | 第二篇 | 【draw-style：地图更新的可见过程】 | `image_page` | `图片 12.png` 零裁切全景 |
| **32**| 38 | 第二篇 | 【draw-style：从证据到图形】 | `image_page` | `图片 13.png` 零裁切全景 |
| **33**| 44, 45| 第二篇 | 【renhua：构建动机与白描求真规范】 | `merged_matrix` (5卡) | 逸夫楼水印 (动机先行) |
| **34**| 41 | 第二篇 | 【renhua：改写效果对比】 | `image_page` | `图片 14.png` 零裁切全景 |
| **35**| 47 | 过渡页 | 第三篇：文献发现与精确导入 | Agenda Transition | 第三章高亮提纲 |
| **36**| 48, 49| 第三篇 | 四端协作减少重复操作 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **37**| 50, 51| 第三篇 | 【drone-literature-scout：探索动机与资源边界】 | `merged_matrix` (5卡) | 逸夫楼水印 (动机先行) |
| **38**| 46 | 第三篇 | 【drone-literature-scout：正式来源】 | `image_page` | `图片 15.png` 零裁切全景 |
| **39**| 53, 54| 第三篇 | 【drone-literature-scout：评分只是方向判断的入口】 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **40**| 55, 56| 第三篇 | 【drone-literature-scout：周期更新与可核对清单】 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **41**| 57, 58| 第三篇 | 【searching-at-scale：规模化发现与证据分层】 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **42**| 53 | 第三篇 | 【zotero-obsidian-paper-import：三端联动动机与身份】 | `image_page` | `图片 16.png` 零裁切全景 (动机先行) |
| **43**| 54 | 第三篇 | 【zotero-obsidian-paper-import：核验全文】 | `image_page` | `图片 17.png` 零裁切全景 |
| **44**| 55 | 第三篇 | 【zotero-obsidian-paper-import：去重与附件】 | `image_page` | `图片 18.png` 零裁切全景 |
| **45**| 56 | 第三篇 | 【zotero-obsidian-paper-import：笔记回到原文】 | `image_page` | `图片 19.png` 零裁切全景 |
| **46**| 57 | 第三篇 | 【zotero-obsidian-paper-import：原文回到笔记】 | `image_page` | `图片 20.png` 零裁切全景 |
| **47**| 58 | 第三篇 | 跨平台阅读需要同步与路径配合 | `image_page` | 跨端同步示意图 |
| **48**| 63 | 过渡页 | 第四篇：核心精读与知识复用 | Agenda Transition | 第四章高亮提纲 |
| **49**| 66, 67| 第四篇 | 【read-paper-analysis-highlight：精读动机与模块解释】 | `merged_matrix` (5卡) | 逸夫楼水印 (动机先行) |
| **50**| 61 | 第四篇 | 【read-paper-analysis-highlight：公式与实现】 | `image_page` | `图片 21.png` 零裁切全景 |
| **51**| 62 | 第四篇 | 【read-paper-analysis-highlight：实验与条件】 | `image_page` | `图片 22.png` 零裁切全景 |
| **52**| 63 | 第四篇 | 【read-paper-analysis-highlight：作者信息】 | `image_page` | `图片 23.png` 零裁切全景 |
| **53**| - | 第四篇 | 【read-paper-analysis-highlight：追溯重点引用】 | `image_page` | `图片 24.png` 零裁切全景 |
| **54**| 64 | 第四篇 | 【read-paper-analysis-highlight：原生高亮】 | `image_page` | `图片 25.png` 零裁切全景 |
| **55**| 65 | 第四篇 | 【read-paper-analysis-highlight：完整检查图表】 | `image_page` | `图片 26.png` 零裁切全景 |
| **56**| 66 | 第四篇 | 【read-paper-analysis-highlight：短句帮助记忆】 | `image_page` | `图片 27.png` 零裁切全景 |
| **57**| 68 | 第四篇 | 【read-paper-analysis-highlight：补齐前置知识】 | `image_page` | `图片 29.png` 零裁切全景 (聚合连续) |
| **58**| 71 | 第四篇 | 【read-paper-analysis-highlight：个人理解】 | `image_page` | `图片 34.png` 零裁切全景 (聚合连续) |
| **59**| 73 | 第四篇 | 【read-paper-analysis-highlight：资源与问题延伸】 | `image_page` | 资源延伸图 (聚合连续) |
| **60**| 67 | 第四篇 | 【obsidian-note-style：笔记构建动机与折叠展开】 | `image_page` | `图片 28.png` 零裁切全景 (动机先行) |
| **61**| 70 | 第四篇 | 【obsidian-note-style：知识点互联】 | `image_page` | `图片 31.png` 零裁切全景 |
| **62**| - | 第四篇 | 【obsidian-note-style：按需打开前置概念】 | `image_page` | `图片 32.png` 零裁切全景 |
| **63**| - | 第四篇 | 【obsidian-note-style：颜色标出阅读重点】 | `image_page` | `图片 33.png` 零裁切全景 |
| **64**| 72 | 第四篇 | 【obsidian-note-style：目录与图像】 | `image_page` | 目录与图像混排 (聚合连续) |
| **65**| 69 | 第四篇 | 【draw-style：在笔记中的分层图纸呈现】 | `image_page` | `图片 30.png` 零裁切全景 |
| **66**| 74 | 第四篇 | 让工具产物进入自己的研究判断 | `image_page` | 人工对账示意图 (哲学收束) |
| **67**| 85 (新)| 第四篇 | 模型 Benchmark，分数之外还要看成本与边界 | `benchmark_page` | `benchmack的对比.jpg` 零裁切展台 |
| **68**| 80 | 过渡页 | 第五篇：体系治理与未来展望 | Agenda Transition | 第五章高亮提纲 |
| **69**| 75 | 第五篇 | Skills 多起来之后需要明确分工 | `content` (基础单页) | 湖大蓝卡片 (治理痛点) |
| **70**| 76 | 第五篇 | 【skill-contract-lock：原始规则核验】 | `content` (基础单页) | 湖大蓝卡片 |
| **71**| 77 | 第五篇 | 【skill-ecosystem-governor：维护来源与链接】 | `image_page` (大图展台) | `展示自动维护整个体系.png` 零裁切全景 + 2张支撑卡片 |
| **72**| 89, 90| 第五篇 | 【collect-bug-update-accelerate：从问题族到复用方案】 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **73**| 91, 94| 第五篇 | 【workspace-hygiene：工作区与文件可恢复整理】 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **74**| 92, 93| 第五篇 | 迁移验证与公开包准备 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **75**| 95, 96| 第五篇 | 从一次任务积累可复用方法 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **76**| 97, 98| 第五篇 | 下一步，把科研流程做得更轻 | `merged_matrix` (5卡) | 逸夫楼水印 |
| **77**| 85 | 第五篇 | 下一步，把科研流程做得更轻、更顺手 | `content` (基础单页) | 湖大蓝卡片 |
| **78**| 100| 结语页 | 汇报完毕，敬请指正 | Ending | 湖大官方结束页底板 |

---

## 六、Codex 直接执行的完整 Prompt (请直接复制给 Codex)

```markdown
请直接阅读并严格遵循以下两份规范文件：
1. `D:/path/to/share\整理skill\docs\AI科研流程-HNU规划-20260921\CODEX_HANDOFF.md` (v5.0 全量学术终极版)
2. `C:/path/to/skills/academic-native-ppt-design-HNU-style\SKILL.md` (v5.0，重点遵循 Rules 34~40 / QG-34~QG-40)

【任务目标】
请以已经通过质量验证的生成脚本：
`D:/path/to/share\整理skill\docs\AI科研流程-HNU规划-20260921\.skill-contract\rebuild-v4\build_77page_deck.py`
为唯一起点，完整构建并输出全篇 **78 幻灯片**（77 页正文/结构页 + 1 结语页）学术 PPT，并保存至：
`D:/path/to/share\整理skill\docs\AI科研流程-HNU规划-20260921\output\科研Agent与自定义Skills-HNU-77页.pptx`。

【核心质量铁律清单 (必须 100% 达成)】
1. **renhua 人性化语言风格与学术零人称 (Rule 39 / QG-39 - 重中之重)**：
   - **学术叙述零人称**：正文中坚决杜绝出现第一/第二人称主语（“我”、“你”、“他”、“我们”、“大家”），一律改为客观目的句（如“为了减少反复拖动文本框的时间：……”、“针对文献漏检实际问题：……”）；
   - **大厂黑话与伪架构套话彻底清洗**：坚决禁止“壁垒”、“可复用资产”、“原子化”、“沉淀为...”、“高频科研方法”、“标准化通信与状态感知”、“自愈闭环”等晦涩夸大词汇，严格执行白描替换（详见总纲对照表）；
   - **分点开头要点强制加粗 (Bullet-Lead Bold)**：每个分点开头必须显式加粗核心论点或根本动作（格式为 `**核心动机/要点**：...`），突出听众方法；
   - **标点洁净**：严禁破折号（——）与分号（；），不留孤立断字，中英文保留空格。
2. **Skill 叙事逻辑、痛点先行与大纲聚合 (Rule 40 / QG-40)**：
   - **痛点与动机先行**：每个 Skill 必须先讲透“为什么做/解决什么痛点”（参考总纲第一节动机对齐表），绝不能一上来就悬空讲构建；
   - **相同 Skill 页面绝对连续**：严格按照总纲第五节 78 页大纲执行，read-paper-analysis-highlight 全部连续讲完后，再进入 obsidian-note-style 体系，绝无跨 Skill 穿插割裂；
   - **核心哲学呈现**：明确表达“需求驱动、AI是强力辅助但无法替代学者判断”的核心观点。
3. **分点自适应展开、真实交替宽方差与行距同页一致性 (Rule 35 / QG-35 & Rule 36 / QG-36)**：
   - 彻底打破固定 2 分点限制，根据卡片尺寸与信息深度动态配置 2~5 条分点，长难句主动拆解为独立短点；
   - **字数覆盖率严格在 [60% ~ 95%] 区间分布**，相邻卡片呈现真实起伏交替（约一半相邻极差 $\ge 25\%$，另一半保持 $5\%\sim 25\%$ 微差，单页极差 $\ge 20\%$），摆脱机械死板感；
   - **行间距自适应与同页完全一致**：默认基准 18 ~ 20 pt；紧凑框可微调至 17.0 pt，开阔框可微调至 21 ~ 23 pt；**同一页内所有正文文本框行间距必须保持完全一致 (方差 0.0 pt)**。
4. **色彩调和与行高齐平 (Rule 34 / QG-34 & Rule 37 / QG-37)**：
   - 单页显性主题色严控在 2~3 种（上限 <= 4 种），严禁轮盘彩虹色；暖赭黄仅用于警示与成本，中性事实严禁乱用；
   - **湖大红等深色底文字纯白铁律**：以湖大红 (`#A6232B` / `#8B1D23`) 为背景底色的卡片或文本框，其内部所有标题、正文与标注文字**必须强制设为纯白色 (`#FFFFFF` / `bg1`)**，严禁使用深灰或黑字导致难以辨识！
   - 同一行并排卡片高度严格物理齐平 (Delta h = 0)；左下角水印处保持开阔避让，消除压顶残块。
5. **模型演进全景纵深 (Rule 38 / QG-38)**：
   - 严格落实 Slides 5、6、7 三页纵深架构，全面覆盖 DeepSeek (V2~V4.1)、GLM (4~5.3)、OpenAI GPT (4~GPT-6 Astra)、Google Gemini (1.5~3.8 Flash)、小米 MiMo (v2~v2.6 Pro) 等主流谱系与选型防线。
6. **素材零裁切与新增图像实装 (Zero-Crop Invariant, QG-29)**：
   - 全部 39 张图源必须使用 `crop=None`，以 `contain` 完整等比呈现；
   - **特别提醒**：用户新增图片素材 `D:/path/to/slide-assets\展示自动维护整个体系.png` (1132x850) **必须在第 71 页【skill-ecosystem-governor：维护来源与链接】实装**，采用大图展台并在底部配备 2 张各 2 分点的支撑卡片，严禁遗漏！
7. **全自动化门禁验收**：
   生成完成后，运行底层 OpenXML 审计脚本：
   `python C:/path/to/skills/academic-native-ppt-design-HNU-style\scripts\audit_hnu_deck.py --input "D:/path/to/share\整理skill\docs\AI科研流程-HNU规划-20260921\output\科研Agent与自定义Skills-HNU-77页.pptx"`
   必须达成 `0 Critical / 0 High` 门禁完全通过！
```
