# academic-native-ppt-design-HNU-style

湖南大学风格、纯原生可编辑的学术与工业汇报 PPT 设计系统。

Hunan University–style academic PPT system with fully native, editable `.pptx` output.

**Skill 版本 v5.0.0（formal_release）** · 姊妹仓库：[Unified-Scholarflow-Skills](https://github.com/w5711112/Unified-Scholarflow-Skills)

---

## 解决什么问题

组会与答辩 PPT 反复出现三类麻烦：模板版式不够统一，边框和字号靠手感，每次手工排版耗时且质量不稳。本系统把可复用的视觉与排版规则写成 Skill，以湖南大学标准母版作底板，生成完全原生、可继续编辑的 `.pptx`。

适合湖南大学学术汇报、组会答辩、工程方案汇报。默认规则带有维护者的汇报密度习惯，可以按自己的需要调整字阶、分点数量和页面节奏。

## 效果（10 张）

下列为实际成品截屏。完整分辨率见 `docs/images/`。

![01](docs/images/hnu-01.webp)

![02](docs/images/hnu-02.webp)

![03](docs/images/hnu-03.webp)

![04](docs/images/hnu-04.webp)

![05](docs/images/hnu-05.webp)

![06](docs/images/hnu-06.webp)

![07](docs/images/hnu-07.webp)

![08](docs/images/hnu-08.webp)

![09](docs/images/hnu-09.webp)

![10](docs/images/hnu-10.webp)

## 为什么单独成仓

湖大母版、印章禁飞区、字阶门禁和通用学术 PPT 分属不同规则集。拆成独立仓库后，检索和安装路径更清楚。

主仓库 [Unified-Scholarflow-Skills](https://github.com/w5711112/Unified-Scholarflow-Skills) 负责科研阅读与治理。两仓互相链接：从科研链能找到高质量汇报出口，做 PPT 时也能回到文献与笔记来源。

## 几条代表规则

完整条文见 `SKILL.md` 与 `references/`（Rules 1–41、QG-01–QG-41）。

1. **事实按来源写。** 参考材料中的技术栈、指标和名词原样使用。
2. **双模态字阶（QG-03）。** 精炼页正文约 16–18 pt。高密页 13–14.5 pt 并配合多子区域拆分。大标题不超过 28 pt。部分历史参考写有“绝对 ≥14 pt”，执行以最新主 Skill 与 QG-03 为准。
3. **印章禁飞区。** 左下角印章与建筑水印区域避开正文卡片。
4. **原生可编辑。** 文本框、形状、表格保持 DrawingML 对象。
5. **自动门禁。** `scripts/audit_hnu_deck.py` 检查字阶、溢出、Markdown 字符泄漏和卡片净高。

## 仓库结构

```text
academic-native-ppt-design-HNU-style/
├─ SKILL.md              # 主规则
├─ references/           # 模板、字阶、版式、门禁等参考
├─ scripts/              # 模板引擎与审计脚本
├─ templates/            # 湖大母版
├─ assets/               # 素材
├─ docs/images/          # 上列 10 张效果图
├─ RELEASE_NOTES.md
└─ README.md
```

## 安装与使用

```powershell
git clone https://github.com/w5711112/academic-native-ppt-design-HNU-style.git
# 将完整目录放入技能路径，例如 ~/.agents/skills/academic-native-ppt-design-HNU-style/
```

在 Codex、Claude Code、Kimi Code、MiMo Desktop 等宿主中，提到“湖大 PPT / HNU PPT / 湖南大学模板”即可触发。生成后执行：

```powershell
python scripts\audit_hnu_deck.py --input <你的.pptx>
```

依赖以 `scripts/` 内说明与 Python 环境为准，通常需要 `python-pptx`。

## 使用边界

母版与印章素材按学校视觉规范整理与再实现，使用时遵守学校标识规定。自动门禁检查字阶、溢出和结构问题；内容是否适合具体答辩由作者判断。

安全问题见 [SECURITY.md](SECURITY.md)。

## Related

- 科研检索、精读、绘图与治理：[Unified-Scholarflow-Skills](https://github.com/w5711112/Unified-Scholarflow-Skills)

## 许可

见 [LICENSE](LICENSE)。第三方组件见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

---

## English summary

**academic-native-ppt-design-HNU-style** builds Hunan University–style academic decks as fully native, editable PowerPoint files. It encodes layout archetypes, dual-mode type scales, seal keep-out zones, and automated quality gates (Rules/QG 1–41).

This repo is the slide-side sibling of [Unified-Scholarflow-Skills](https://github.com/w5711112/Unified-Scholarflow-Skills). Ten showcase renders are under `docs/images/`. Defaults reflect one maintainer’s presentation habits and can be edited.
