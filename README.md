# academic-native-ppt-design-HNU-style

> 湖南大学风格、纯原生可编辑的学术 / 工业汇报 PPT 设计系统。  
> Hunan University–style academic PPT system with fully native, editable `.pptx` output.

**Skill 版本 v5.0.0（formal_release）** · 仓库 v1.0 姊妹仓库：[Unified-Scholarflow-Skills](https://github.com/w5711112/Unified-Scholarflow-Skills)

---

## 解决什么问题

组会与答辩 PPT 反复出现三类麻烦：官方模板版式不够统一，边框和字号靠手感，每次手工排版耗时且质量不稳。本系统把可复用的视觉与排版规则写成 Skill，用湖南大学标准母版作底板，生成**完全原生、可继续编辑**的 `.pptx`，而不是贴图稿或损坏结构的自动排版。

适合：湖南大学学术汇报、组会答辩、工程方案汇报。  
不适合：随意海报、非湖大视觉体系、需要像素级品牌定制却不愿改规则的场景。

规则、阈值和示例带有维护者的使用习惯，可以按自己的汇报密度修改；不改也能按默认门禁出稿。

## 效果（10 张）

下列为实际成品截屏，未再裁切压缩出内容损失之外的修饰。完整分辨率见 `docs/images/`。

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

- 湖大母版、印章禁飞区、字阶门禁和通用学术 PPT 不是一套规则，拆开后检索和安装都更清楚。  
- 主仓库 [Unified-Scholarflow-Skills](https://github.com/w5711112/Unified-Scholarflow-Skills) 专注科研阅读与治理；旧的 `academic-native-ppt-design` / `codex-ppt` 已撤下，演示能力以本仓库为准。  
- 两仓互链：主仓看科研链时顺手找到高质量汇报出口；本仓做 PPT 时能回到文献与笔记来源。

## 实现里几条代表规则

完整条文见 `SKILL.md` 与 `references/`（Rules 1–41、QG-01–QG-41）。

1. **零捏造事实**  
   参考材料写什么技术栈和指标就用什么。不编造中间件名、时延或精度。

2. **双模态字阶（QG-03）**  
   精炼页正文约 16–18 pt；高密页 13–14.5 pt 并配合多子区域拆分。大标题不超过 28 pt。与部分历史参考中的“绝对 ≥14 pt”表述存在冲突时，以最新主 Skill 与 QG-03 为准。

3. **印章禁飞区**  
   左下角印章/建筑水印区域避开正文卡片，防止遮挡与压线。

4. **原生可编辑**  
   文本框、形状、表格保持 DrawingML 对象，不用整页位图冒充“PPT”。

5. **自动门禁**  
   `scripts/audit_hnu_deck.py` 检查字阶、溢出、Markdown 字符泄漏、卡片净高，避免手工目检遗漏。

## 仓库结构

```text
academic-native-ppt-design-HNU-style/
├─ SKILL.md              # 主规则
├─ references/           # 模板、字阶、版式、门禁等强制参考
├─ scripts/              # 模板引擎与审计脚本
├─ templates/            # 湖大母版
├─ examples/             # 示例
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

在 Codex、Claude Code、Kimi Code、**MiMo Desktop** 等宿主中，提到“湖大 PPT / HNU PPT / 湖南大学模板”即可触发。生成后可用：

```powershell
python scripts\audit_hnu_deck.py --input <你的.pptx>
```

依赖以 `scripts/` 内说明与 Python 环境为准（通常为 `python-pptx`）。

## 公开版边界

- 母版与印章素材来自校内视觉规范的整理与再实现，请遵守学校标识使用规定，不要用于与湖南大学无关的商业品牌冒充。  
- 示例工程数据与讲稿内容仅作排版演示，不含未授权内部资料。  
- 自动门禁通过只说明结构与字阶达标，内容是否适合答辩仍需作者自己判断。

## Related

- 科研检索 / 精读 / 绘图 / 治理：[Unified-Scholarflow-Skills](https://github.com/w5711112/Unified-Scholarflow-Skills)

## 许可

见 [LICENSE](LICENSE)。第三方组件见仓库内说明。安全问题见 [SECURITY.md](SECURITY.md)。

---

## English summary

**academic-native-ppt-design-HNU-style** builds Hunan University–style academic decks as fully native, editable PowerPoint files. It encodes layout archetypes, dual-mode type scales, seal keep-out zones, and automated quality gates (Rules/QG 1–41).

This repo is the slide-side sibling of [Unified-Scholarflow-Skills](https://github.com/w5711112/Unified-Scholarflow-Skills). Ten showcase renders are under `docs/images/`. Defaults reflect one maintainer’s presentation habits and are safe to edit.
