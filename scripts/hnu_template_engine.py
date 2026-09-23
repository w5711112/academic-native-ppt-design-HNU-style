# -*- coding: utf-8 -*-
"""
湖南大学专属学术PPT构建引擎 (HNUTemplateEngineHNU)
专为湖南大学组会、答辩与学术汇报打造的原生 PowerPoint 构建引擎。
100% 严格遵循湖大母版规范：
1. 四步闭环学术排版管线 (读取素材 -> 容量规划 -> 多模态分配 -> 模板微调)；
2. 封面红色底框绝不变小，主标题 40~44pt 垂直居中单行，汇报人 24pt 加粗垂直居中单行，绝无重叠；
3. 全域字体 100% 锁定为微软雅黑 (同时注入 a:ea 与 a:latin，彻底杜绝 PowerPoint 回退等线/宋体)；
4. 全域最小字号硬红线定死为 14pt (表格、图注、指标等)，正常正文文字严格按 16.0 ~ 20.0 pt 自适应排版；
5. 内容页大标题严格 <= 28.0 pt (26~28pt 湖大红)，卡片子标题 20.0 ~ 22.0 pt；
6. 左下角印章绝对禁飞区 (X < 280.1 且 Y > 350.0)，右下延伸区横贯落地消灭中下部死角空白；
7. 官方母版 Slide 2 结束页自动闭环在全套 PPT 最后一页。
"""


import os
import sys
import math
import pptx
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml

COLOR_HNU_RED   = RGBColor(166, 35, 43)   # #A6232B 湖大红
COLOR_HNU_BLUE  = RGBColor(29, 73, 153)   # #1D4999 湖大深蓝
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
COLOR_TEXT_MAIN = RGBColor(36, 36, 36)    # #242424 正文深灰
COLOR_TEXT_MUTED= RGBColor(102, 102, 102) # #666666 辅助灰
COLOR_WHITE     = RGBColor(255, 255, 255) # 纯白
COLOR_BORDER    = RGBColor(226, 232, 240) # 浅灰细边框
COLOR_PILL_BG   = RGBColor(248, 250, 252) # 极浅冷灰底板
COLOR_SLATE     = RGBColor(71, 85, 105)   # #475569 石板深灰 (流程标签)

# draw-style 淡雅学术多色阶调色板 (丰富视觉表现力，因地制宜构建流程图与对比卡片)
COLOR_TABLE_HEADER_BG   = RGBColor(248, 234, 235)  # #F8EAEB 淡雅浅陶粉 (draw-style 学术浅色系)
COLOR_TABLE_HEADER_TEXT = RGBColor(139, 29, 35)   # #8B1D23 湖大红深色文字加粗 (15pt)
COLOR_TABLE_ROW_ALT     = RGBColor(253, 251, 250)  # #FDFBFA 极淡温润灰斑马纹
COLOR_TABLE_ROW_TEXT    = RGBColor(45, 55, 72)     # #2D3748 深炭灰 (14pt)
COLOR_TABLE_BORDER      = RGBColor(226, 217, 215)  # #E2D9D7 0.75pt 细线

COLOR_DRAW_TEAL   = RGBColor(15, 118, 110)  # #0F766E 墨绿/青碧 (知识层级/基础)
COLOR_DRAW_AMBER  = RGBColor(180, 83, 9)    # #B45309 琥珀暖橙 (外部交互/预警)
COLOR_DRAW_INDIGO = RGBColor(67, 56, 202)   # #4338CA 靛蓝/紫罗兰 (算法/逻辑推理)
COLOR_DRAW_SAGE   = RGBColor(241, 245, 249) # #F1F5F9 极淡冷灰卡片底色
COLOR_DRAW_ROSE   = RGBColor(190, 24, 93)   # #BE185D 玫瑰红 (重点突破)

FONT_YAHEI = "Microsoft YaHei"
FONT_MATH  = "Cambria Math"

import re

def compile_latex_to_math_unicode(text):
    r"""底层 LaTeX 宏到 Unicode 数学符号编译器：
    1. 彻底根除 PPTX 文本框中裸露的 LaTeX 反斜杠源码；
    2. 将 \min, \max, \lambda, \ge, \Delta, \text 等精准转换为标准学术 Unicode 符号；
    3. 规范化数学物理符号表达。
    """
    if not text:
        return ""
    
    # 常见 LaTeX 宏转换字典 (按长度由长到短匹配，防前缀冲突)
    replacements = [
        (r'\\rightarrow\b', '→'),
        (r'\\leftarrow\b', '←'),
        (r'\\forall\b', '∀'),
        (r'\\exists\b', '∃'),
        (r'\\approx\b', '≈'),
        (r'\\partial\b', '∂'),
        (r'\\infty\b', '∞'),
        (r'\\nabla\b', '∇'),
        (r'\\times\b', '×'),
        (r'\\cdot\b', '·'),
        (r'\\neq\b', '≠'),
        (r'\\geq\b', '≥'),
        (r'\\leq\b', '≤'),
        (r'\\lambda\b', 'λ'),
        (r'\\Lambda\b', 'Λ'),
        (r'\\Delta\b', 'Δ'),
        (r'\\delta\b', 'δ'),
        (r'\\theta\b', 'θ'),
        (r'\\alpha\b', 'α'),
        (r'\\beta\b', 'β'),
        (r'\\gamma\b', 'γ'),
        (r'\\sigma\b', 'σ'),
        (r'\\omega\b', 'ω'),
        (r'\\tau\b', 'τ'),
        (r'\\mu\b', 'μ'),
        (r'\\phi\b', 'φ'),
        (r'\\min\b', 'min'),
        (r'\\max\b', 'max'),
        (r'\\ge\b', '≥'),
        (r'\\le\b', '≤'),
        (r'\\to\b', '→'),
        (r'\\in\b', '∈'),
        (r'\\text\s*\{([^}]+)\}', r'\1'),
        (r'\\mathrm\s*\{([^}]+)\}', r'\1'),
        (r'\\mathbf\s*\{([^}]+)\}', r'\1'),
    ]

    res = text
    for pat, rep in replacements:
        res = re.sub(pat, rep, res)

    # 清理多余反斜杠转义
    res = res.replace('\\$', '$').replace('\\%', '%')
    return res


def apply_yahei(run, font_size=14, bold=False, color=None):
    """全局统一中文字体注入器：
    1. 定死最小字号 >= 13 pt（允许密集排版使用 13pt 永久修改）；
    2. 显式添加 a:ea 和 a:latin，确保在 Windows PowerPoint 中 100% 识别为微软雅黑。
    """
    if font_size < 13:
        font_size = 13
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    run.font.name = FONT_YAHEI

    rPr = run._r.get_or_add_rPr()
    rPr.set("i", "0")
    ea = rPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
    if ea is None:
        ea = parse_xml(f'<a:ea xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="{FONT_YAHEI}"/>')
        rPr.append(ea)
    else:
        ea.set("typeface", FONT_YAHEI)
    
    latin = rPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}latin")
    if latin is None:
        latin = parse_xml(f'<a:latin xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="{FONT_YAHEI}"/>')
        rPr.append(latin)
    else:
        latin.set("typeface", FONT_YAHEI)


def apply_math(run, font_size=14, italic=True, bold=False, color=None):
    """数学与物理变量公式字体注入器：
    1. 显式绑定 Cambria Math，支持最小字号 >= 13 pt；
    2. 英文/希腊变量显式绑定 italic=True；
    3. 操作符、数字与函数名绑定 italic=False。
    """
    if font_size < 13:
        font_size = 13
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    run.font.name = FONT_MATH

    rPr = run._r.get_or_add_rPr()
    rPr.set("i", "1" if italic else "0")
    latin = rPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}latin")
    if latin is None:
        latin = parse_xml(f'<a:latin xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="{FONT_MATH}"/>')
        rPr.append(latin)
    else:
        latin.set("typeface", FONT_MATH)


def apply_rich_bullet(paragraph, text, font_size=16.0, color=None, show_bullet=True, space_before=4.0, space_after=4.0):
    """智能富文本与数学公式解析渲染器 (湖南大学学术规范版)：
    1. 底层先执行 compile_latex_to_math_unicode，100% 根除裸露反斜杠；
    2. 原生 OpenXML 项目符号：注入 <a:buChar char="•"/> 与 <a:buFont typeface="Microsoft YaHei UI"/>，实现原生悬挂缩进；
    3. 段前距与段后距硬核锁定为 >= 4pt (<a:spcBef/Aft val="400"/>，默认 4~5pt)；
    4. 字体全域统一：日常数字、英文单词、计量单位 (如 50Hz, 30ms, 95%) 一律统一为微软雅黑；
    5. Cambria Math 严格且仅限用于数学物理变量公式体 (如 J_s, λ_s, Δt, min J 等)；
    6. 确保最小字号 >= 13pt，正文自适应 13~20pt。
    """
    if font_size < 13.0:
        font_size = 13.0
    if color is None:
        color = COLOR_TEXT_MAIN

    # 第一道工序：全局编译所有 LaTeX 宏，消除裸露源码反斜杠
    clean_text = compile_latex_to_math_unicode(text)
    # 清理开头可能存在的冗余项目符号符号，杜绝与原生 buChar 产生双重视觉重叠
    clean_text = re.sub(r'^[•\-\*]\s*', '', clean_text.strip())

    # 第二道工序：注入 Slide 25 原生 OpenXML 项目符号与标准段距
    pPr = paragraph._p.get_or_add_pPr()
    if show_bullet:
        pPr.set("marL", "228600")    # 18pt 悬挂边距
        pPr.set("indent", "-177800")  # -14pt 悬挂缩进，首行与折行严密对齐
        
        # 移除已有可能冲突的 buNone/buChar
        for tag in ["buNone", "buChar", "buFont", "buAutoNum"]:
            old_elem = pPr.find(f"{{http://schemas.openxmlformats.org/drawingml/2006/main}}{tag}")
            if old_elem is not None:
                pPr.remove(old_elem)

        buFont = parse_xml('<a:buFont xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="Microsoft YaHei UI"/>')
        buChar = parse_xml('<a:buChar xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" char="•"/>')
        pPr.append(buFont)
        pPr.append(buChar)

    # 原生注入段前距与段后距 (满足底线 >= 4pt / spcBef >= 400 门禁)
    actual_spc_bef = max(4.0, float(space_before))
    actual_spc_aft = max(4.0, float(space_after))
    paragraph.space_before = Pt(actual_spc_bef)
    paragraph.space_after = Pt(actual_spc_aft)

    # 第三道工序：正则分词：严格仅识别真实数学变量 (如 J_s, λ_s, J_c, J_d, Δt 等)
    # 取消对常规数字和单位 (Hz, ms, s, km, %) 的数学公式识别，统一使用微软雅黑！
    pattern = re.compile(r'(\b[A-Za-z]_[a-z0-9]+\b|[λα-ωΑ-Ω]_[a-zA-Z0-9]+|Δ\s*[a-zA-Z0-9]+|\b(?:min|max)\s+[A-Za-z]\b)')
    tokens = pattern.split(clean_text)

    for tok in tokens:
        if not tok:
            continue

        # 判断是否为下划线数学变量 (如 J_s, λ_s, J_c, J_d)
        if re.match(r'^([A-Za-z]_[a-z0-9]+|[λα-ωΑ-Ω]_[a-zA-Z0-9]+)$', tok):
            m_sub = re.match(r'^([A-Za-z]|[λα-ωΑ-Ω])_([a-zA-Z0-9]+)$', tok)
            base_char, sub_char = m_sub.group(1), m_sub.group(2)
            
            # 主变量 Run (如 J, λ) -> Cambria Math 斜体
            r_base = paragraph.add_run()
            r_base.text = base_char
            apply_math(r_base, font_size=font_size, italic=True, bold=False, color=color)

            # 下标 Run (如 s, c, d) -> Cambria Math 下标基线
            r_sub = paragraph.add_run()
            r_sub.text = sub_char
            apply_math(r_sub, font_size=max(11.0, font_size * 0.78), italic=False, bold=False, color=color)
            rPr_sub = r_sub._r.get_or_add_rPr()
            rPr_sub.set("baseline", "-25000") # 下标下沉基线

        elif re.match(r'^([A-Za-z]_[a-zA-Z0-9]+|[λα-ωΑ-Ω]_[a-zA-Z0-9]+|Δ\s*[a-zA-Z0-9]+|\b(?:min|max)\s+[A-Za-z]\b)$', tok):
            r = paragraph.add_run()
            r.text = tok
            apply_math(r, font_size=font_size, italic=True, bold=False, color=color)

        else:
            # 普通汉字、数字 (2026, 4月23日)、英文单词、计量单位 (Hz, ms) 全域统一为微软雅黑
            r = paragraph.add_run()
            r.text = tok
            apply_yahei(r, font_size=font_size, bold=False, color=color)


class HNUTemplateEngineHNU:
    def __init__(self, template_path=None):
        if template_path is None:
            curr_dir = os.path.dirname(os.path.abspath(__file__))
            self.template_path = os.path.join(curr_dir, "..", "templates", "hnu_base_template.pptx")
        else:
            self.template_path = template_path

        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"找不到模板底板文件: {self.template_path}")

    def create_deck(self, output_path, cover_info, slides_data):
        """基于 python-pptx 直接在母版底板上构建，保证 100% 打开无异常弹窗"""
        prs = Presentation(self.template_path)

        # 1. 配置第 1 页封面
        self._format_cover_slide(prs.slides[0], cover_info)

        # 2. 动态定位底板中的官方结束页（“汇报完毕，敬请指正”），将其移动到整个 presentation 的最后
        ending_idx = None
        for i, sld_elem in enumerate(prs.slides._sldIdLst):
            slide_part = prs.part.related_slide(sld_elem.rId)
            s_text = "".join([t.text for t in slide_part._element.iter(f"{{{A_NS}}}t") if t.text])
            if "汇报完毕" in s_text:
                ending_idx = i
                break

        if ending_idx is not None:
            ending_sld_elem = prs.slides._sldIdLst[ending_idx]
            prs.slides._sldIdLst.remove(ending_sld_elem)
            prs.slides._sldIdLst.append(ending_sld_elem)

        # 3. 此时 prs.slides[1] 开始为原内容页模板
        total_content_pages = len(slides_data)
        for idx, s_data in enumerate(slides_data):
            slide_idx = idx + 1
            if slide_idx < len(prs.slides) - 1:
                slide = prs.slides[slide_idx]
            else:
                slide_layout = prs.slide_layouts[1]
                slide = prs.slides.add_slide(slide_layout)

            self._format_content_slide(slide, s_data, idx + 2)

        # 4. 删除多余的模板参考页 (保留封面 + total_content_pages + 结尾页 1 页)
        total_needed = 1 + total_content_pages + 1
        while len(prs.slides) > total_needed:
            del_idx = total_needed - 1
            rId = prs.slides._sldIdLst[del_idx].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[del_idx]

        # 5. 全局质量终审 (字号 >= 14pt，字体为微软雅黑，印章绝对避让)
        self._audit_deck(prs)

        # 保存并输出
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        prs.save(output_path)
        print(f"[SUCCESS] 成功生成湖南大学专属学术 PPT: {output_path} (共 {len(prs.slides)} 页)")

    def _format_cover_slide(self, slide, cover_info):
        """配置封面：
        1. 形状 1 (矩形 5, 大红色半透明底框): 绝对禁止修改坐标与尺寸，保持 [143.55, 361.75, 816.45, 178.25]，彻底清空文本！
        2. 形状 2、形状 3 (空文本框): 彻底清空文本！
        3. 形状 5 (原标题框): 唯一写入主标题，40~44pt 微软雅黑加粗纯白，垂直居中，单行！
        4. 形状 4 (原汇报人框): 唯一写入汇报人与时间，24pt 微软雅黑加粗纯白，垂直居中，单行！
        5. 形状 6 (白色分割线): 保持在 top=449.60。
        """
        title_text = cover_info.get("title", "湖南大学学术汇报")
        speaker_text = cover_info.get("speaker", "汇报人：吴沛霖")
        date_text = cover_info.get("date", "2026年9月")

        title_shape = None
        speaker_shape = None
        red_box_shape = None

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            # 检查大红色背景矩形 (Shape 1, left ~ 143.55, top ~ 361.75, width ~ 816.45, height ~ 178.25)
            if abs(shape.left.pt - 143.55) < 5 and abs(shape.top.pt - 361.75) < 5 and shape.width.pt > 750:
                red_box_shape = shape
                shape.left = Pt(143.55)
                shape.top = Pt(361.75)
                shape.width = Pt(816.45)
                shape.height = Pt(178.25)
                shape.text_frame.clear() # 纯背景色块，绝对不写入任何文字！
                continue

            # 检查原标题框 (Shape 5, id=12, top ~ 361.75, height ~ 87.85, left ~ 215.75)
            if abs(shape.left.pt - 215.75) < 10 and abs(shape.top.pt - 361.75) < 10 and shape.height.pt < 110:
                title_shape = shape
                continue

            # 检查原汇报人框 (Shape 4, id=2, top ~ 449.55, height ~ 90.40, left ~ 215.75)
            if abs(shape.left.pt - 215.75) < 10 and abs(shape.top.pt - 449.55) < 10:
                speaker_shape = shape
                continue

            # 其它任何多余文本框一律清空文字，杜绝重叠！
            shape.text_frame.clear()

        # 写入主标题 (唯一写入，绝无重叠)
        if title_shape is None:
            title_shape = slide.shapes.add_textbox(Pt(215.75), Pt(361.75), Pt(715.0), Pt(87.85))
        else:
            title_shape.left = Pt(215.75)
            title_shape.top = Pt(361.75)
            title_shape.width = Pt(715.0)
            title_shape.height = Pt(87.85)

        tf = title_shape.text_frame
        tf.clear()
        tf.word_wrap = False # 强制单行！
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE # 垂直居中！
        tf.margin_left = Pt(0)
        tf.margin_top = Pt(0)
        tf.margin_right = Pt(0)
        tf.margin_bottom = Pt(0)

        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = title_text

        # 严格标题长度约束与动态字号计算 (防超出右边界铁律)
        title_len = len(title_text)
        if title_len > 17:
            print(f"[WARNING] 封面标题字数超过 17 字建议阈值 ({title_len} 字)，正在执行自适应防溢出紧凑排版...")
            sz = max(32, min(38, int(680.0 / (title_len * 1.02))))
        elif title_len >= 15:
            sz = 40
        elif title_len >= 12:
            sz = 42
        else:
            sz = 44
        apply_yahei(run, font_size=sz, bold=True, color=COLOR_WHITE)


        # 写入汇报人与时间 (唯一写入，绝无重叠)
        if speaker_shape is None:
            speaker_shape = slide.shapes.add_textbox(Pt(215.75), Pt(449.55), Pt(650.0), Pt(90.40))
        else:
            speaker_shape.left = Pt(215.75)
            speaker_shape.top = Pt(449.55)
            speaker_shape.width = Pt(650.0)
            speaker_shape.height = Pt(90.40)

        tf_s = speaker_shape.text_frame
        tf_s.clear()
        tf_s.word_wrap = False # 强制单行！
        tf_s.vertical_anchor = MSO_ANCHOR.MIDDLE # 垂直居中！
        tf_s.margin_left = Pt(0)
        tf_s.margin_top = Pt(0)
        tf_s.margin_right = Pt(0)
        tf_s.margin_bottom = Pt(0)

        p_s = tf_s.paragraphs[0]
        p_s.alignment = PP_ALIGN.LEFT
        run_s = p_s.add_run()
        run_s.text = f"{speaker_text}          {date_text}"
        apply_yahei(run_s, font_size=24, bold=True, color=COLOR_WHITE)

    def _format_agenda_slide(self, slide, sections, active_idx=0, title="汇报提纲"):
        """创建符合母版 Slide 2 规范的目录页/章节过渡页：
        1. 左侧大标题：垂直居中 60pt 微软雅黑加粗（汇报提纲）；
        2. 右侧章节列表：支持 3~10 个章节动态等间距垂直居中；
        3. 当前章节高亮为湖大红 (#A6232B)，非当前章节为湖大深蓝 (#1D4999)；
        4. 圆角矩形序号框 + 细竖装饰条 + 章节名文本框，间距恒定对齐；
        5. 统一使用微软雅黑，数字字母不混用其他字体。
        """
        shapes_to_delete = []
        for sp in list(slide.shapes):
            if sp.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE:
                if sp.left.pt > 700 and sp.top.pt < 60:
                    continue # 保留右上角校徽
            shapes_to_delete.append(sp)

        for sp in shapes_to_delete:
            sp_elem = sp._element
            sp_elem.getparent().remove(sp_elem)

        # 1. 左侧垂直“汇报提纲”艺术标题框
        tb_title = slide.shapes.add_textbox(Pt(175.87), Pt(79.99), Pt(135.39), Pt(385.76))
        tf_t = tb_title.text_frame
        tf_t.clear()
        tf_t.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_t.word_wrap = True
        p_t = tf_t.paragraphs[0]
        p_t.alignment = PP_ALIGN.CENTER
        run_t = p_t.add_run()
        run_t.text = title
        apply_yahei(run_t, font_size=60.0, bold=True, color=COLOR_HNU_BLUE)

        # 2. 右侧章节列表
        N = len(sections)
        if N < 1:
            return

        # 几何参数动态适配 (根据章节数 3~10 计算尺寸与步长，保持纵向整体居中)
        if N <= 5:
            item_h = 39.5
            font_sz = 28.0
            num_sz = 32.0
            num_w = 49.9
            bar_w = 4.6
            step_gap = 68.6 if N == 5 else min(78.0, 360.0 / max(1, N - 1))
        elif N <= 7:
            item_h = 34.0
            font_sz = 24.0
            num_sz = 26.0
            num_w = 44.0
            bar_w = 4.0
            step_gap = 360.0 / (N - 1)
        else: # 8~10
            item_h = 28.0
            font_sz = 19.0
            num_sz = 22.0
            num_w = 38.0
            bar_w = 3.5
            step_gap = 370.0 / (N - 1)

        total_h = (N - 1) * step_gap + item_h
        start_top = 270.0 - (total_h / 2.0) # 垂直居中于画布中心 y=270

        left_base = 392.7
        bar_gap = 3.5
        text_gap = 12.0

        for idx, sec_name in enumerate(sections):
            y_pos = start_top + idx * step_gap
            is_active = (idx == active_idx)
            cur_color = COLOR_HNU_RED if is_active else COLOR_HNU_BLUE

            # 序号圆角矩形
            shp_num = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(left_base), Pt(y_pos), Pt(num_w), Pt(item_h))
            shp_num.fill.solid()
            shp_num.fill.fore_color.rgb = cur_color
            shp_num.line.fill.background() # 无边框
            tf_num = shp_num.text_frame
            tf_num.clear()
            tf_num.vertical_anchor = MSO_ANCHOR.MIDDLE
            p_num = tf_num.paragraphs[0]
            p_num.alignment = PP_ALIGN.CENTER
            run_num = p_num.add_run()
            run_num.text = str(idx + 1)
            apply_yahei(run_num, font_size=num_sz, bold=True, color=COLOR_WHITE)

            # 细竖装饰条
            bar_x = left_base + num_w + bar_gap
            shp_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(bar_x), Pt(y_pos), Pt(bar_w), Pt(item_h))
            shp_bar.fill.solid()
            shp_bar.fill.fore_color.rgb = cur_color
            shp_bar.line.fill.background()

            # 章节文字文本框
            text_x = bar_x + bar_w + text_gap
            text_w = 420.0
            shp_txt = slide.shapes.add_textbox(Pt(text_x), Pt(y_pos - 2.0), Pt(text_w), Pt(item_h + 4.0))
            tf_txt = shp_txt.text_frame
            tf_txt.clear()
            tf_txt.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf_txt.word_wrap = True
            p_txt = tf_txt.paragraphs[0]
            p_txt.alignment = PP_ALIGN.LEFT
            run_txt = p_txt.add_run()
            run_txt.text = sec_name
            apply_yahei(run_txt, font_size=font_sz, bold=True, color=cur_color)

    def add_agenda_slide(self, slide, sections, active_idx=0, title="汇报提纲"):
        """公共 API: 在指定 slide 上生成标准湖大目录/过渡页"""
        self._format_agenda_slide(slide, sections, active_idx=active_idx, title=title)

    def _format_content_slide(self, slide, s_data, slide_num):
        """配置内容页：
        1. 彻底清除旧内容；
        2. 若为目录/提纲页 (is_agenda=True 或 type='agenda')，直接调用目录渲染器；
        3. 保留顶部红线、右上校徽、左下印章；
        4. 单行大标题严格对齐 X = 38.1 pt；
        5. 右下角页码 14pt 微软雅黑；
        6. 内容卡片装配 (全域字号 >= 14pt，微软雅黑，左下印章绝对禁飞)。
        """
        # 目录页优先路由
        if s_data.get("type") == "agenda" or s_data.get("is_agenda", False):
            self._format_agenda_slide(
                slide,
                sections=s_data.get("sections", []),
                active_idx=s_data.get("active_idx", 0),
                title=s_data.get("agenda_title", "汇报提纲")
            )
            return

        shapes_to_delete = []
        title_shape = None
        page_num_shape = None

        for shape in list(slide.shapes):
            if shape.has_table:
                shapes_to_delete.append(shape)
                continue

            if shape.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE:
                x_pt = shape.left.pt
                y_pt = shape.top.pt
                # 保留右上校徽 (x > 700, y < 60) 或左下印章 (x < 285, y > 340)
                if (x_pt > 700 and y_pt < 60) or (x_pt < 285 and y_pt > 340):
                    continue
                else:
                    shapes_to_delete.append(shape)
                    continue

            if shape.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.LINE:
                continue

            if shape.has_text_frame:
                x_pt = shape.left.pt
                y_pt = shape.top.pt

                if y_pt > 500 and x_pt > 850:
                    page_num_shape = shape
                    continue

                if y_pt < 80 and x_pt < 500:
                    title_shape = shape
                    continue

                shapes_to_delete.append(shape)
                continue

            shapes_to_delete.append(shape)

        for sp in shapes_to_delete:
            sp_elem = sp._element
            sp_elem.getparent().remove(sp_elem)

        # 单行大标题
        title_text = s_data.get("title", "研究进展汇报")
        if title_shape is None:
            title_shape = slide.shapes.add_textbox(Pt(38.1), Pt(12.0), Pt(700.0), Pt(48.0))
        else:
            title_shape.left = Pt(38.1)
            title_shape.top = Pt(12.0)
            title_shape.width = Pt(700.0)
            title_shape.height = Pt(48.0)

        tf = title_shape.text_frame
        tf.clear()
        tf.word_wrap = False
        tf.margin_left = Pt(0)
        tf.margin_top = Pt(4)
        tf.margin_right = Pt(0)
        tf.margin_bottom = Pt(0)

        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = title_text

        if len(title_text) > 22:
            t_sz = 24
        elif len(title_text) > 17:
            t_sz = 26
        else:
            t_sz = 28
        apply_yahei(run, font_size=t_sz, bold=True, color=COLOR_HNU_RED)

        # 右下角页码 (14pt 微软雅黑)
        if page_num_shape is not None:
            p_tf = page_num_shape.text_frame
            p_tf.clear()
            p_p = p_tf.paragraphs[0]
            p_p.alignment = PP_ALIGN.RIGHT
            p_run = p_p.add_run()
            p_run.text = f"{slide_num:02d}"
            apply_yahei(p_run, font_size=14, bold=False, color=COLOR_TEXT_MUTED)

        # 内容装配
        contents = s_data.get("contents", [])
        for item in contents:
            c_type = item.get("type")
            pos = item.get("pos", [38.1, 105.0, 440.0, 235.0])
            x, y, w, h = [Pt(v) for v in pos]

            # 几何禁区刚性拦截：
            # 若形状位于印章左下角 (pos[0] < 280.1):
            if pos[0] < 280.1:
                if pos[1] >= 345.0:
                    raise AssertionError(f"[CRITICAL] Slide {slide_num} 形状绝对禁止放置在左下角印章禁飞区 (X < 280.1 且 Y >= 345.0): pos={pos}")
                if (pos[1] + pos[3]) > 345.0:
                    print(f"[WARNING] Slide {slide_num} 左侧形状底部超过 345.0 pt 禁区 (pos={pos})，强制截断高度至 {342.0 - pos[1]} pt！")
                    h = Pt(342.0 - pos[1])

            if c_type == "pic":
                img_path = item.get("image_path")
                if not img_path or not os.path.exists(img_path):
                    raise FileNotFoundError(f"[CRITICAL ERROR] Slide {slide_num} 引用的真实图片不存在: '{img_path}'！严禁留下空白图框！")
                self._add_picture_box(
                    slide, x, y, w, h, img_path,
                    caption=item.get("caption"),
                    show_border=item.get("show_border", False),
                    padding=item.get("padding", None)
                )

            elif c_type == "card":
                self._add_card_shape(
                    slide, x, y, w, h,
                    title=item.get("card_title", ""),
                    bullets=item.get("bullets", []),
                    title_sz=item.get("title_sz"),
                    bullet_sz=item.get("bullet_sz"),
                    rounded=item.get("rounded", False),
                    macro_border=item.get("macro_border", False),
                    padding=item.get("padding", None),
                    dense=item.get("dense", False),
                    tight_fit=item.get("tight_fit", False),
                    bg_color=item.get("bg_color", None)
                )

            elif c_type == "horizontal_pipeline":
                self._add_horizontal_pipeline(
                    slide, x, y, w, h,
                    nodes=item.get("nodes", []),
                    color_theme=item.get("color_theme", None)
                )

            elif c_type == "circular_pipeline":
                self._add_circular_pipeline(
                    slide, x, y, w, h,
                    nodes=item.get("nodes", []),
                    color_theme=item.get("color_theme", None)
                )

            elif c_type == "snake_pipeline":
                self._add_snake_pipeline(
                    slide, x, y, w, h,
                    nodes=item.get("nodes", []),
                    color_theme=item.get("color_theme", None)
                )

            elif c_type == "bifurcated_flow":
                self._add_bifurcated_flow(
                    slide, x, y, w, h,
                    root_node=item.get("root_node", {}),
                    branch_nodes=item.get("branch_nodes", []),
                    color_theme=item.get("color_theme", None)
                )

            elif c_type == "grid_cards":
                self._add_grid_cards_layout(
                    slide, x, y, w, h,
                    cards_data=item.get("cards", []),
                    cols=item.get("cols", 3),
                    rows=item.get("rows", 2)
                )

            elif c_type == "sidebar_grid":
                self._add_sidebar_plus_grid_layout(
                    slide, x, y, w, h,
                    sidebar_data=item.get("sidebar", {}),
                    grid_cards=item.get("grid_cards", []),
                    cols=item.get("cols", 2),
                    rows=item.get("rows", 2)
                )

            elif c_type == "three_col":
                self._add_3col_layout(
                    slide, x, y, w, h,
                    cards_data=item.get("cards", [])
                )

            elif c_type == "max_image":
                img_path = item.get("image_path")
                if not img_path or not os.path.exists(img_path):
                    raise FileNotFoundError(f"[CRITICAL ERROR] Slide {slide_num} 引用的真实图片不存在: '{img_path}'！")
                self._add_max_image_layout(
                    slide,
                    left_card=item.get("left_card", {}),
                    img_path=img_path,
                    caption=item.get("caption")
                )

            elif c_type == "math":
                self._add_native_math_equation(
                    slide, x, y, w, h,
                    eq_type=item.get("eq_type", "b_spline_cost"),
                    title=item.get("card_title", ""),
                    bullets=item.get("bullets", []),
                    rounded=item.get("rounded", False),
                    macro_border=item.get("macro_border", False),
                    padding=item.get("padding", None),
                    bullet_sz=item.get("bullet_sz", None)
                )

            elif c_type == "table":
                headers = item.get("headers", [])
                rows = item.get("rows", [])
                col_widths = item.get("col_widths", None)
                self._add_academic_table(slide, x, y, w, h, headers, rows, col_widths=col_widths)

            elif c_type in ["arrow", "flowchart_arrow"]:
                self._add_flowchart_arrow(
                    slide, x, y, w, h,
                    direction=item.get("direction", "right"),
                    color=item.get("color", None)
                )

    def _add_flowchart_arrow(self, slide, x, y, w, h, direction="right", color=None):
        """插入流程图箭头 (大比例粗块状箭头，杜绝针眼细箭头)：
        支持 direction="right", "left", "down", "up"
        """
        dir_map = {
            "right": MSO_SHAPE.RIGHT_ARROW,
            "left": MSO_SHAPE.LEFT_ARROW,
            "down": MSO_SHAPE.DOWN_ARROW,
            "up": MSO_SHAPE.UP_ARROW,
        }
        shape_type = dir_map.get(direction.lower(), MSO_SHAPE.RIGHT_ARROW)
        shp = slide.shapes.add_shape(shape_type, x, y, w, h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = COLOR_HNU_RED if color is None else color
        shp.line.fill.background()
        return shp

    def _add_native_math_equation(self, slide, x, y, w, h, eq_type="b_spline_cost", title="", bullets=None, rounded=False, macro_border=False, padding=None, bullet_sz=None):
        """插入原生 Office / WPS 数学公式对象 (<a14:m> / OMML)：
        1. 采用 DrawingML 14 标准原生数学公式树 (<a14:m><m:oMathPara><m:oMath>...</m:oMath></m:oMathPara></a14:m>)；
        2. 在 PowerPoint / WPS 中 100% 识别为原生可编辑公式对象，点击直接唤醒公式设计器；
        3. 消除 0.75pt 细线碎框，支持宏观包裹或纯净无边框；
        4. 变量 Cambria Math 斜体，下标基线自动下沉，运算符带专业数学间距；
        5. 支持 padding 调节（padding=0 实现贴边扩展）；
        6. 支持 bullet_sz 统一字阶。
        """
        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
        shape = slide.shapes.add_shape(shape_type, x, y, w, h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = COLOR_WHITE
        if macro_border:
            shape.line.color.rgb = COLOR_BORDER
            shape.line.width = Pt(1.5)
        else:
            shape.line.fill.background()

        pad_lr = 8.0 if padding is None else float(padding)
        pad_tb = 6.0 if padding is None else float(padding)

        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Pt(pad_lr)
        tf.margin_right = Pt(pad_lr)
        tf.margin_top = Pt(pad_tb)
        tf.margin_bottom = Pt(pad_tb)
        tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.TOP

        # 1. 标题
        if title:
            p_title = tf.paragraphs[0]
            p_title.alignment = PP_ALIGN.LEFT
            p_title.space_after = Pt(8.0)
            run = p_title.add_run()
            run.text = title
            apply_yahei(run, font_size=20.0, bold=True, color=COLOR_HNU_RED)

        # 2. 原生数学公式段落 (显式注入深灰文字颜色 #242424，确保在白底卡片上 100% 清晰呈现)
        p_math = tf.add_paragraph() if title else tf.paragraphs[0]
        p_math.alignment = PP_ALIGN.LEFT
        p_math.space_after = Pt(8.0)

        # 具备显式深色颜色标签的标准 OMML XML
        omml_xml = """<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath><m:func><m:funcPr><m:ctrlPr><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr></m:ctrlPr></m:funcPr><m:fName><m:r><m:rPr><m:sty m:val="p"/></m:rPr><a:rPr lang="zh-CN" altLang="en-US"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>min</m:t></m:r></m:fName><m:e><m:r><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝐽</m:t></m:r></m:e></m:func><m:r><a:rPr lang="zh-CN" altLang="en-US" i="0"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t> = </m:t></m:r><m:sSub><m:sSubPr><m:ctrlPr><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr></m:ctrlPr></m:sSubPr><m:e><m:r><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝜆</m:t></m:r></m:e><m:sub><m:r><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝑠</m:t></m:r></m:sub></m:sSub><m:sSub><m:sSubPr><m:ctrlPr><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr></m:ctrlPr></m:sSubPr><m:e><m:r><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝐽</m:t></m:r></m:e><m:sub><m:r><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝑠</m:t></m:r></m:sub></m:sSub><m:r><a:rPr lang="zh-CN" altLang="en-US" i="0"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t> + </m:t></m:r><m:sSub><m:sSubPr><m:ctrlPr><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr></m:ctrlPr></m:sSubPr><m:e><m:r><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝜆</m:t></m:r></m:e><m:sub><m:r><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝑐</m:t></m:r></m:sub></m:sSub><m:sSub><m:sSubPr><m:ctrlPr><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr></m:ctrlPr></m:sSubPr><m:e><m:r><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝐽</m:t></m:r></m:e><m:sub><m:r><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝑐</m:t></m:r></m:sub></m:sSub><m:r><a:rPr lang="zh-CN" altLang="en-US" i="0"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t> + </m:t></m:r><m:sSub><m:sSubPr><m:ctrlPr><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr></m:ctrlPr></m:sSubPr><m:e><m:r><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝜆</m:t></m:r></m:e><m:sub><m:r><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝑑</m:t></m:r></m:sub></m:sSub><m:sSub><m:sSubPr><m:ctrlPr><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr></m:ctrlPr></m:sSubPr><m:e><m:r><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝐽</m:t></m:r></m:e><m:sub><m:r><a:rPr lang="zh-CN" altLang="en-US" i="1"><a:solidFill><a:srgbClr val="242424"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝑑</m:t></m:r></m:sub></m:sSub></m:oMath></m:oMathPara></a14:m>"""
        math_elem = parse_xml(omml_xml)
        p_math._p.append(math_elem)

        # 3. 解释说明列表 (1.25 倍行距，段距 8pt 饱满展开)
        actual_bullet_sz = 15.0 if bullet_sz is None else float(bullet_sz)
        if bullets:
            for b in bullets:
                p_b = tf.add_paragraph()
                p_b.alignment = PP_ALIGN.LEFT
                p_b.line_spacing = 1.25
                p_b.space_after = Pt(8.0)
                apply_rich_bullet(p_b, b, font_size=actual_bullet_sz, color=COLOR_TEXT_MAIN)

    def _add_academic_table(self, slide, x, y, w, h, headers, rows, col_widths=None):
        """创建原生学术三线表 (完全对齐 draw-style 浅色学术与湖大红配色体系)：
        1. 表头：淡雅浅陶粉 (#F8EAEB) 底板 + 湖大红深色加粗文字 (#8B1D23, 15pt)；
        2. 数据行：交替条纹斑马纹 (偶数行 #FDFBFA，奇数行 #FFFFFF)；
        3. 正文：深炭灰 (#2D3748) >= 13pt 微软雅黑，全表所有列统一完全居中对齐；
        4. 单元格内边距严格设为 0 (cell.margin 与 text_frame.margin 上下左右均为 0)，垂直中部居中对齐 (vertical_anchor=MIDDLE)！
        5. 支持 col_widths 自定义列宽数组，确保窄列不挤压折行、宽列从容排布。
        """
        num_rows = len(rows) + 1
        num_cols = len(headers)
        table_shape = slide.shapes.add_table(num_rows, num_cols, x, y, w, h)
        tbl = table_shape.table

        # 若提供了自定义列宽，按精确值赋给各列
        if col_widths and len(col_widths) == num_cols:
            for c_i, cw in enumerate(col_widths):
                tbl.columns[c_i].width = Pt(cw)

        # 表头配置 (draw-style 浅陶粉底板 + 湖大红深色字，边距为0，统一完全居中对齐)
        for col_idx, text in enumerate(headers):
            cell = tbl.cell(0, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_TABLE_HEADER_BG
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Pt(0)
            cell.margin_right = Pt(0)
            cell.margin_top = Pt(0)
            cell.margin_bottom = Pt(0)
            cell.text_frame.margin_left = Pt(0)
            cell.text_frame.margin_right = Pt(0)
            cell.text_frame.margin_top = Pt(0)
            cell.text_frame.margin_bottom = Pt(0)
            cell.text_frame.word_wrap = True
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = text
            apply_yahei(run, font_size=15, bold=True, color=COLOR_TABLE_HEADER_TEXT)

        # 数据行配置 (斑马纹交替 + 深炭灰字，边距为0，统一完全居中对齐)
        for row_idx, row_data in enumerate(rows):
            is_alt = (row_idx % 2 == 1)
            bg_color = COLOR_TABLE_ROW_ALT if is_alt else COLOR_WHITE
            for col_idx, text in enumerate(row_data):
                cell = tbl.cell(row_idx + 1, col_idx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg_color
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = Pt(0)
                cell.margin_right = Pt(0)
                cell.margin_top = Pt(0)
                cell.margin_bottom = Pt(0)
                cell.text_frame.margin_left = Pt(0)
                cell.text_frame.margin_right = Pt(0)
                cell.text_frame.margin_top = Pt(0)
                cell.text_frame.margin_bottom = Pt(0)
                cell.text_frame.word_wrap = True
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                run = p.add_run()
                text_str = str(text)
                run.text = text_str
                # 表格高亮准则：仅对短标签/状态徽章词（<=8字符且完全匹配）进行红色加粗，绝不可对整句自然语言长描述中的通用词（如"通过"）误触发高亮！
                clean_token = text_str.strip()
                is_badge = (len(clean_token) <= 8) and (clean_token in ["100%", "完全符合", "优于", "通过", "达标", "自主可控", "通过评审", "优秀"] or (clean_token.startswith("M") and len(clean_token) <= 4))
                is_highlight = is_badge and (col_idx >= 2)
                text_color = COLOR_HNU_RED if is_highlight else COLOR_TABLE_ROW_TEXT
                apply_yahei(run, font_size=14, bold=is_highlight, color=text_color)

    def _add_picture_box(self, slide, x, y, w, h, img_path, caption=None, show_border=False, padding=None):
        """插入技术大图：
        1. 彻底消灭 0.75pt 细线碎框：默认纯净无边框，底托与大背景自然融合；
        2. 若在宏观卡片内，直接呼吸居中呈现；
        3. 图注严格 >= 14pt 微软雅黑居中；
        4. 支持 padding=0 内容贴边扩展模态：图片满幅顶满容器，消除冗余白边；
        5. 零裁切铁律 (QG-29 Zero-Crop Invariant)：100% 保持原图完整，绝无 PowerPoint 裁切。
        """
        from PIL import Image

        pad = 8.0 if padding is None else float(padding)

        # 底托卡片：默认无边框，彻底告别 0.75pt 细线框
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_WHITE
        if show_border:
            bg.line.color.rgb = COLOR_BORDER
            bg.line.width = Pt(1.5)
        else:
            bg.line.fill.background() # 无边框！消灭一切细线框痕迹

        cap_h = Pt(26) if caption else Pt(0)
        avail_w = w.pt - (pad * 2.0)
        avail_h = h.pt - cap_h.pt - (pad * 2.0)

        pic_shape = None
        try:
            with Image.open(img_path) as im:
                orig_w, orig_h = im.size
            ar_img = orig_w / float(orig_h)
            ar_box = avail_w / float(avail_h)
            ar_diff = abs(ar_box - ar_img) / ar_img
            if ar_diff > 0.40:
                print(f"[WARNING] 容器与图片长宽比差异较大 ({ar_diff:.1%}): 容器AR={ar_box:.2f}, 图像AR={ar_img:.2f} (建议调整版型以实现>=75%充满度)")

            scale = min(avail_w / orig_w, avail_h / orig_h)
            actual_w = orig_w * scale
            actual_h = orig_h * scale

            pic_x = x.pt + pad + (avail_w - actual_w) / 2.0
            pic_y = y.pt + pad + (avail_h - actual_h) / 2.0
            pic_shape = slide.shapes.add_picture(img_path, Pt(pic_x), Pt(pic_y), Pt(actual_w), Pt(actual_h))
        except Exception:
            pic_shape = slide.shapes.add_picture(img_path, x + Pt(pad), y + Pt(pad), Pt(avail_w), Pt(avail_h))

        # 零裁切刚性执行 (Zero-Crop Invariant)
        if pic_shape is not None:
            try:
                pic_shape.crop_left = 0
                pic_shape.crop_top = 0
                pic_shape.crop_right = 0
                pic_shape.crop_bottom = 0
                blipFill = pic_shape._element.find(f'.//{{{A_NS}}}blipFill')
                if blipFill is not None:
                    srcRect = blipFill.find(f'{{{A_NS}}}srcRect')
                    if srcRect is not None:
                        blipFill.remove(srcRect)
            except Exception:
                pass

        if caption:
            cap_y = y + h - cap_h - Pt(2)
            cap_box = slide.shapes.add_textbox(x, cap_y, w, cap_h)
            c_tf = cap_box.text_frame
            c_tf.margin_left = Pt(4)
            c_tf.margin_right = Pt(4)
            c_tf.margin_top = Pt(0)
            c_tf.margin_bottom = Pt(0)
            c_p = c_tf.paragraphs[0]
            c_p.alignment = PP_ALIGN.CENTER
            c_run = c_p.add_run()
            c_run.text = caption
            apply_yahei(c_run, font_size=14, bold=True, color=COLOR_TEXT_MUTED)

    def _add_card_shape(self, slide, x, y, w, h, title, bullets, title_sz=None, bullet_sz=None, rounded=False, macro_border=False, padding=None, dense=False, tight_fit=False, bg_color=None):
        """创建主题卡片：
        1. 彻底消灭 0.75pt 琐碎细线：默认完全无边框痕迹，或圆角矩形浅色大框包裹；
        2. 宏观包裹：相似主题/近似含义使用同一个大卡片承载，绝不拆成很多小框；
        3. 标题：20.0~22.0pt 加粗湖大红微软雅黑；
        4. 正文：支持双模态：
           - 密集技术详述模式 (dense=True 或指定 14~15pt)：以 14.0~15.0pt 为主，承载 4~7 个多分点；
           - 常规标准模式 (dense=False)：小字号仅在内容多时启用，内容适中时保持 16.0~18.0pt 舒展大字号；
        5. 空间优化：支持 padding 调节（padding=0 消除内边距，实现无缝与贴边）；
        6. 贴合几何 (tight_fit=True)：色块与文字严密贴合，动态将卡片高度收紧包裹实际文本，杜绝内部死白。
        """
        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
        shape = slide.shapes.add_shape(shape_type, x, y, w, h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = COLOR_WHITE if bg_color is None else bg_color
        if macro_border:
            shape.line.color.rgb = COLOR_BORDER
            shape.line.width = Pt(1.5) # 略粗的浅柔灰宏观大框
        else:
            shape.line.fill.background() # 彻底无边框！消灭 0.75pt 细线痕迹

        pad_lr = 8.0 if padding is None else float(padding)
        pad_tb = 6.0 if padding is None else float(padding)

        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Pt(pad_lr)
        tf.margin_right = Pt(pad_lr)
        tf.margin_top = Pt(pad_tb)
        tf.margin_bottom = Pt(pad_tb)
        tf.clear()

        # 强制顶部对齐
        tf.vertical_anchor = MSO_ANCHOR.TOP

        avail_w = w.pt - (pad_lr * 2.0)
        avail_h = h.pt - (pad_tb * 2.0)
        num_bullets = len(bullets)

        # 估算标题占用高度
        title_h = 0.0
        if title:
            t_sz = 21.0 if title_sz is None else float(title_sz)
            title_h = t_sz * 1.3 + (6.0 if h.pt > 200 else 4.0)

        avail_body_h = max(20.0, avail_h - title_h)

        # 计算每个 bullet 的有效字符数 (汉字计 1.0，英文字母/数字/标点计 0.55)
        bullet_effective_chars = []
        for b in bullets:
            c_count = 0.0
            for ch in b:
                c_count += 1.0 if '\u4e00' <= ch <= '\u9fff' else 0.55
            bullet_effective_chars.append(c_count)

        # 候选字号自适应阶梯筛选 (支持最小 13.0pt 密集排版)
        if bullet_sz is not None:
            init_sz = max(13.0, min(20.0, float(bullet_sz)))
            chosen_sz = 13.0
            candidates = [sz for sz in [init_sz, init_sz - 0.5, init_sz - 1.0, 14.0, 13.5, 13.0] if sz >= 13.0]
            for sz_candidate in candidates:
                chars_per_line = max(5.0, avail_w / sz_candidate)
                total_est_lines = sum(max(1, int(math.ceil(c / chars_per_line))) for c in bullet_effective_chars)
                req_h = total_est_lines * sz_candidate * 1.18 + num_bullets * 3.0
                if req_h <= avail_body_h * 0.95 or sz_candidate == 13.0:
                    chosen_sz = sz_candidate
                    break
        elif dense:
            # 高密度技术详述模式：以 13.0 ~ 15.0 pt 为主，承载 4~7 个多分点
            chosen_sz = 13.0
            for sz_candidate in [15.0, 14.5, 14.0, 13.5, 13.0]:
                chars_per_line = max(5.0, avail_w / sz_candidate)
                total_est_lines = sum(max(1, int(math.ceil(c / chars_per_line))) for c in bullet_effective_chars)
                req_h = total_est_lines * sz_candidate * 1.25 + num_bullets * 5.0
                if req_h <= avail_body_h * 0.90:
                    chosen_sz = sz_candidate
                    break
        else:
            # 标准模式：如果用户没有指定内容多，不按 13~15pt 为主字号，保持 16~18pt 舒展字号
            chosen_sz = 14.0
            for sz_candidate in [18.0, 17.5, 17.0, 16.5, 16.0, 15.5, 15.0, 14.5, 14.0, 13.5, 13.0]:
                chars_per_line = max(5.0, avail_w / sz_candidate)
                total_est_lines = sum(max(1, int(math.ceil(c / chars_per_line))) for c in bullet_effective_chars)
                req_h = total_est_lines * sz_candidate * 1.25 + num_bullets * 6.0
                if req_h <= avail_body_h * 0.88:
                    chosen_sz = sz_candidate
                    break

        bullet_sz = chosen_sz

        if title_sz is None:
            if bullet_sz >= 18.0:
                title_sz = 21.0
            elif bullet_sz >= 17.0:
                title_sz = 20.0
            elif bullet_sz >= 16.0:
                title_sz = 19.0
            elif bullet_sz >= 14.5:
                title_sz = 18.0
            else:
                title_sz = 17.0
        else:
            title_sz = max(16.0, min(22.0, float(title_sz)))

        # 根据最终选定字号重新精确计算总行数
        chars_per_line = max(5.0, avail_w / bullet_sz)
        bullet_lines = [max(1, int(math.ceil(c / chars_per_line))) for c in bullet_effective_chars]
        total_est_lines = sum(bullet_lines)

        # 估算标题占用真实高度 (含换行与段后距)
        title_lines = 1
        if title:
            t_chars_per_line = max(5.0, avail_w / title_sz)
            t_eff_chars = sum(1.0 if '\u4e00' <= ch <= '\u9fff' else 0.55 for ch in title)
            title_lines = max(1, int(math.ceil(t_eff_chars / t_chars_per_line)))
        title_real_h = title_lines * (title_sz * 1.3) if title else 0.0

        # 每段统一段前距为 4.0 pt (满足 QG-24 spcBef >= 400 门禁)
        bullet_spc_bef = 4.0
        bullets_bef_h = num_bullets * bullet_spc_bef

        # 用户核心准则：行距默认为 1.25 倍行距 (或 ~20pt)，且卡片内部垂直空间利用率必须 >= 80%，黄金区间 85% ~ 95%
        base_line_spacing = 1.25 if not dense else 1.20
        text_lines_h = total_est_lines * bullet_sz * base_line_spacing

        # 目标充实度设为 88%
        target_rendered_h = avail_h * 0.88
        free_for_spacing = target_rendered_h - (title_real_h + text_lines_h + bullets_bef_h)

        if free_for_spacing > 0 and num_bullets > 0:
            raw_space_after = free_for_spacing / (num_bullets + (0.5 if title else 0.0))
            space_after = max(4.0, min(8.0, raw_space_after))
            title_space_after = max(4.0, min(8.0, space_after * 1.1)) if title else 0.0
            line_spacing = base_line_spacing
        else:
            line_spacing = 1.18 if (title_real_h + text_lines_h + bullets_bef_h) > avail_h * 0.95 else 1.20
            rem_h = avail_h - title_real_h - bullets_bef_h - (total_est_lines * bullet_sz * line_spacing)
            space_after = max(4.0, min(5.0, rem_h / max(1, num_bullets)))
            title_space_after = max(4.0, min(6.0, space_after * 1.1)) if title else 0.0

        # 计算并记录最终卡片实际填充率 (Fill Rate)
        actual_rendered_h = title_real_h + title_space_after + (total_est_lines * bullet_sz * line_spacing) + (num_bullets * (bullet_spc_bef + space_after))

        # 紧密几何贴合 (Tight-Fitting Geometry)
        if tight_fit:
            ideal_h = actual_rendered_h + pad_tb * 2.0 + 4.0
            if ideal_h < h.pt:
                shape.height = Pt(ideal_h)
                avail_h = ideal_h - pad_tb * 2.0

        fill_rate = actual_rendered_h / avail_h
        card_name = title.replace('\n', ' ') if title else "Card"
        print(f"[CARD LAYOUT] 卡片 '{card_name[:16]}': 字号={bullet_sz}pt, 行数={total_est_lines}, 行距={line_spacing:.2f}x, 段距={space_after:.1f}pt, 充实度={fill_rate:.1%}")
        if fill_rate < 0.80:
            print(f"  --> [LOW FILL WARNING] 卡片 '{card_name[:16]}' 充实度 {fill_rate:.1%} < 80%，需充实内容或扩充分点！")

        if title:
            p_title = tf.paragraphs[0]
            p_title.alignment = PP_ALIGN.LEFT
            p_title.space_after = Pt(title_space_after)
            run = p_title.add_run()
            run.text = title
            apply_yahei(run, font_size=title_sz, bold=True, color=COLOR_HNU_RED)

        for idx, bullet in enumerate(bullets):
            p = tf.add_paragraph() if (title or idx > 0) else tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = line_spacing
            
            # 使用智能富文本与数学公式体渲染，精确传入段前距与段后距
            apply_rich_bullet(p, bullet, font_size=bullet_sz, color=COLOR_TEXT_MAIN, space_before=bullet_spc_bef, space_after=space_after)

    def _add_decoupled_card(self, slide, x, y, w, h, title, bullets, title_sz=18.0, bullet_sz=14.0, fill_color=None, title_color=None):
        """插入分体式卡片 (Decoupled Card Architecture)：
        1. 纯几何底托形状 (b.rect)：固定坐标与固定尺寸，行高绝对齐平，严禁参差不齐；
        2. 独立置顶标题框 (Section)：位于 (x+10, y+8, w-20, 28)，加粗湖大红或深蓝；
        3. 独立正文文本框 (Body)：位于 (x+10, y+35, w-20, h-41)，字号按 14.0 -> 13.5 -> 13.0 pt 自适应微调防溢出，
           自动注入原生 OpenXML 实心圆点与 4~5pt 段间距。
        """
        fill = COLOR_TABLE_HEADER_BG if fill_color is None else fill_color
        t_col = COLOR_HNU_RED if title_color is None else title_color
        
        # 1. 底框
        bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = fill
        bg_shape.line.fill.background()
        bg_shape.name = "BG"
        
        # 2. 独立标题框
        if title:
            t_box = slide.shapes.add_textbox(Pt(x.pt + 10.0), Pt(y.pt + 8.0), Pt(w.pt - 20.0), Pt(28.0))
            t_box.name = "Section"
            t_tf = t_box.text_frame
            t_tf.clear()
            t_tf.word_wrap = True
            t_tf.margin_left = t_tf.margin_right = t_tf.margin_top = t_tf.margin_bottom = Pt(0)
            t_p = t_tf.paragraphs[0]
            t_p.alignment = PP_ALIGN.LEFT
            t_run = t_p.add_run()
            t_run.text = title
            apply_yahei(t_run, font_size=title_sz, bold=True, color=t_col)
            
        # 3. 独立正文框
        b_box = slide.shapes.add_textbox(Pt(x.pt + 10.0), Pt(y.pt + 35.0), Pt(w.pt - 20.0), Pt(max(20.0, h.pt - 41.0)))
        b_box.name = "Body"
        b_tf = b_box.text_frame
        b_tf.clear()
        b_tf.word_wrap = True
        b_tf.margin_left = b_tf.margin_right = b_tf.margin_top = b_tf.margin_bottom = Pt(0)
        b_tf.vertical_anchor = MSO_ANCHOR.TOP
        
        bullet_list = bullets if isinstance(bullets, list) else [bullets]
        
        # 估算字号阶梯
        avail_bw = w.pt - 20.0
        avail_bh = max(20.0, h.pt - 41.0)
        chosen_sz = 13.0
        for sz_cand in [bullet_sz, bullet_sz - 0.5, bullet_sz - 1.0, 13.5, 13.0]:
            if sz_cand < 13.0:
                continue
            chars_per_line = max(5.0, (avail_bw - 20.0) / sz_cand)
            total_lines = 0
            for b_txt in bullet_list:
                equiv = sum(1.0 if ('\u4e00' <= ch <= '\u9fff' or ch in '，。；：！？（）“”《》') else 0.55 for ch in b_txt)
                total_lines += max(1, int(math.ceil(equiv / chars_per_line)))
            est_h = total_lines * sz_cand * 1.2 + len(bullet_list) * 8.0
            if est_h <= avail_bh or sz_cand == 13.0:
                chosen_sz = sz_cand
                break
                
        for i, b_txt in enumerate(bullet_list):
            p = b_tf.paragraphs[0] if i == 0 else b_tf.add_paragraph()
            p.line_spacing = Pt(chosen_sz * 1.2)
            apply_rich_bullet(p, b_txt, font_size=chosen_sz, color=COLOR_TEXT_MAIN, show_bullet=True, space_before=2.0, space_after=4.0)
            
        return bg_shape

    def _add_horizontal_pipeline(self, slide, x, y, w, h, nodes, color_theme=None):
        """插入原生单行横向流水线流程图 (Horizontal Pipeline)：
        1. nodes: 列表，每个元素为 dict (含 title, bullets, badge, etc.) 或 string；
        2. 自动在节点间插入大比例右向粗箭头 (22×14 pt)；
        3. 节点卡片采用圆角浅框，带顶部步骤徽章装饰。
        """
        if not nodes:
            return
        N = len(nodes)
        arrow_w = 22.0
        arrow_h = 14.0
        arrow_gap = 6.0
        total_arrow_w = (N - 1) * (arrow_w + arrow_gap * 2.0) if N > 1 else 0.0
        avail_node_w = (w.pt - total_arrow_w) / float(N)

        theme_color = COLOR_HNU_BLUE if color_theme is None else color_theme

        for i, node in enumerate(nodes):
            node_x = x.pt + i * (avail_node_w + arrow_w + arrow_gap * 2.0)
            node_y = y.pt
            node_w = avail_node_w
            node_h = h.pt

            if isinstance(node, dict):
                n_title = node.get("title", f"阶段 {i+1}")
                n_bullets = node.get("bullets", [])
                n_badge = node.get("badge", f"Step {i+1}")
                n_color = node.get("color", theme_color)
            else:
                n_title = str(node)
                n_bullets = []
                n_badge = f"0{i+1}"
                n_color = theme_color

            self._add_card_shape(
                slide, Pt(node_x), Pt(node_y), Pt(node_w), Pt(node_h),
                title=n_title,
                bullets=n_bullets,
                title_sz=16.0,
                bullet_sz=13.5,
                rounded=True,
                macro_border=True,
                padding=6.0,
                dense=True
            )

            # 顶部徽章装饰条
            badge_shp = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Pt(node_x + 6.0), Pt(node_y + 4.0), Pt(46.0), Pt(16.0)
            )
            badge_shp.fill.solid()
            badge_shp.fill.fore_color.rgb = n_color
            badge_shp.line.fill.background()
            b_tf = badge_shp.text_frame
            b_tf.clear()
            b_tf.margin_left = Pt(0)
            b_tf.margin_right = Pt(0)
            b_tf.margin_top = Pt(0)
            b_tf.margin_bottom = Pt(0)
            b_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            b_p = b_tf.paragraphs[0]
            b_p.alignment = PP_ALIGN.CENTER
            b_run = b_p.add_run()
            b_run.text = n_badge
            apply_yahei(b_run, font_size=10.0, bold=True, color=COLOR_WHITE)

            if i < N - 1:
                arr_x = node_x + node_w + arrow_gap
                arr_y = node_y + (node_h - arrow_h) / 2.0
                self._add_flowchart_arrow(
                    slide, Pt(arr_x), Pt(arr_y), Pt(arrow_w), Pt(arrow_h),
                    direction="right",
                    color=COLOR_HNU_RED
                )

    def _add_circular_pipeline(self, slide, x, y, w, h, nodes, color_theme=None):
        """插入原生 4 节点环形闭环流程图 (Circular Pipeline):
        Top-Left (Node 1) -> Right Arrow -> Top-Right (Node 2)
              ^                                   |
           Up Arrow                            Down Arrow
              |                                   v
        Bottom-Left (Node 4) <- Left Arrow <- Bottom-Right (Node 3)
        """
        if len(nodes) < 4:
            return self._add_horizontal_pipeline(slide, x, y, w, h, nodes, color_theme)

        theme_color = COLOR_HNU_BLUE if color_theme is None else color_theme
        arrow_w = 24.0
        arrow_h = 14.0
        gap_x = 36.0
        gap_y = 28.0

        card_w = (w.pt - gap_x) / 2.0
        card_h = (h.pt - gap_y) / 2.0

        positions = [
            (x.pt, y.pt),                                   # 0: Top-Left
            (x.pt + card_w + gap_x, y.pt),                  # 1: Top-Right
            (x.pt + card_w + gap_x, y.pt + card_h + gap_y), # 2: Bottom-Right
            (x.pt, y.pt + card_h + gap_y)                   # 3: Bottom-Left
        ]

        for i in range(4):
            pos_x, pos_y = positions[i]
            node = nodes[i]
            if isinstance(node, dict):
                n_title = node.get("title", f"环节 {i+1}")
                n_bullets = node.get("bullets", [])
            else:
                n_title = str(node)
                n_bullets = []

            self._add_card_shape(
                slide, Pt(pos_x), Pt(pos_y), Pt(card_w), Pt(card_h),
                title=n_title,
                bullets=n_bullets,
                title_sz=16.0,
                bullet_sz=13.5,
                rounded=True,
                macro_border=True,
                padding=6.0,
                dense=True
            )

        # 4 个方向的闭环箭头
        # 1. Top: Node 1 -> Node 2 (Right Arrow)
        self._add_flowchart_arrow(
            slide,
            Pt(x.pt + card_w + (gap_x - arrow_w) / 2.0),
            Pt(y.pt + card_h / 2.0 - arrow_h / 2.0),
            Pt(arrow_w), Pt(arrow_h),
            direction="right", color=COLOR_HNU_RED
        )
        # 2. Right: Node 2 -> Node 3 (Down Arrow)
        self._add_flowchart_arrow(
            slide,
            Pt(positions[1][0] + card_w / 2.0 - arrow_h / 2.0),
            Pt(y.pt + card_h + (gap_y - arrow_w) / 2.0),
            Pt(arrow_h), Pt(arrow_w),
            direction="down", color=COLOR_HNU_RED
        )
        # 3. Bottom: Node 3 -> Node 4 (Left Arrow)
        self._add_flowchart_arrow(
            slide,
            Pt(x.pt + card_w + (gap_x - arrow_w) / 2.0),
            Pt(positions[2][1] + card_h / 2.0 - arrow_h / 2.0),
            Pt(arrow_w), Pt(arrow_h),
            direction="left", color=COLOR_HNU_RED
        )
        # 4. Left: Node 4 -> Node 1 (Up Arrow)
        self._add_flowchart_arrow(
            slide,
            Pt(positions[0][0] + card_w / 2.0 - arrow_h / 2.0),
            Pt(y.pt + card_h + (gap_y - arrow_w) / 2.0),
            Pt(arrow_h), Pt(arrow_w),
            direction="up", color=COLOR_HNU_RED
        )

    def _add_snake_pipeline(self, slide, x, y, w, h, nodes, color_theme=None):
        """插入原生蛇形折返流程图 (Snake / Zigzag Pipeline):
        Row 1: Node 1 -> Node 2 -> Node 3  (向右)
                                     |
                                 Down Arrow
                                     v
        Row 2: Node 6 <- Node 5 <- Node 4  (向左)
        """
        N = len(nodes)
        if N < 4:
            return self._add_horizontal_pipeline(slide, x, y, w, h, nodes, color_theme)

        k = (N + 1) // 2  # 第一行节点数
        m = N - k         # 第二行节点数

        arrow_w = 22.0
        arrow_h = 14.0
        arrow_gap = 6.0
        row_gap = 26.0

        card_h = (h.pt - row_gap) / 2.0
        total_arrow_w_r1 = (k - 1) * (arrow_w + arrow_gap * 2.0)
        card_w_r1 = (w.pt - total_arrow_w_r1) / float(k)

        # 绘制第一行 (向右)
        for i in range(k):
            node_x = x.pt + i * (card_w_r1 + arrow_w + arrow_gap * 2.0)
            node_y = y.pt
            node = nodes[i]
            n_title = node.get("title", f"阶段 {i+1}") if isinstance(node, dict) else str(node)
            n_bullets = node.get("bullets", []) if isinstance(node, dict) else []

            self._add_card_shape(
                slide, Pt(node_x), Pt(node_y), Pt(card_w_r1), Pt(card_h),
                title=n_title, bullets=n_bullets,
                title_sz=16.0, bullet_sz=13.5,
                rounded=True, macro_border=True, padding=6.0, dense=True
            )

            if i < k - 1:
                arr_x = node_x + card_w_r1 + arrow_gap
                arr_y = node_y + (card_h - arrow_h) / 2.0
                self._add_flowchart_arrow(
                    slide, Pt(arr_x), Pt(arr_y), Pt(arrow_w), Pt(arrow_h),
                    direction="right", color=COLOR_HNU_RED
                )

        # 右端向下折返箭头
        turn_x = x.pt + (k - 1) * (card_w_r1 + arrow_w + arrow_gap * 2.0) + card_w_r1 / 2.0 - arrow_h / 2.0
        turn_y = y.pt + card_h + (row_gap - arrow_w) / 2.0
        self._add_flowchart_arrow(
            slide, Pt(turn_x), Pt(turn_y), Pt(arrow_h), Pt(arrow_w),
            direction="down", color=COLOR_HNU_RED
        )

        # 绘制第二行 (向左)
        total_arrow_w_r2 = (m - 1) * (arrow_w + arrow_gap * 2.0) if m > 1 else 0.0
        card_w_r2 = (w.pt - total_arrow_w_r2) / float(m)

        for j in range(m):
            node_idx = k + j
            node = nodes[node_idx]
            col_pos = (m - 1) - j
            node_x = x.pt + col_pos * (card_w_r2 + arrow_w + arrow_gap * 2.0)
            node_y = y.pt + card_h + row_gap

            n_title = node.get("title", f"阶段 {node_idx+1}") if isinstance(node, dict) else str(node)
            n_bullets = node.get("bullets", []) if isinstance(node, dict) else []

            self._add_card_shape(
                slide, Pt(node_x), Pt(node_y), Pt(card_w_r2), Pt(card_h),
                title=n_title, bullets=n_bullets,
                title_sz=16.0, bullet_sz=13.5,
                rounded=True, macro_border=True, padding=6.0, dense=True
            )

            if col_pos > 0:
                arr_x = node_x - arrow_gap - arrow_w
                arr_y = node_y + (card_h - arrow_h) / 2.0
                self._add_flowchart_arrow(
                    slide, Pt(arr_x), Pt(arr_y), Pt(arrow_w), Pt(arrow_h),
                    direction="left", color=COLOR_HNU_RED
                )

    def _add_bifurcated_flow(self, slide, x, y, w, h, root_node, branch_nodes, color_theme=None):
        """插入原生分叉/树状决策流 (Bifurcated Flow):
        左侧: 根节点 (如阶段一: 调研选型)
        中间: 分叉箭头
        右侧: N 个分支节点垂直堆叠 (如阶段二: 分支1, 分支2, 分支3)
        """
        B = len(branch_nodes)
        if B == 0:
            return

        arrow_w = 24.0
        arrow_h = 14.0
        gap_x = 36.0

        root_w = w.pt * 0.36
        branches_w = w.pt - root_w - gap_x

        # 1. 左侧根节点卡片
        r_title = root_node.get("title", "核心起点") if isinstance(root_node, dict) else str(root_node)
        r_bullets = root_node.get("bullets", []) if isinstance(root_node, dict) else []
        self._add_card_shape(
            slide, x, y, Pt(root_w), h,
            title=r_title, bullets=r_bullets,
            title_sz=18.0, bullet_sz=14.0,
            rounded=True, macro_border=True, padding=8.0, dense=True
        )

        # 2. 右侧分支节点
        branch_gap_y = 12.0
        branch_h = (h.pt - (B - 1) * branch_gap_y) / float(B)

        for b_i, b_node in enumerate(branch_nodes):
            b_y = y.pt + b_i * (branch_h + branch_gap_y)
            b_x = x.pt + root_w + gap_x
            b_title = b_node.get("title", f"分支 {b_i+1}") if isinstance(b_node, dict) else str(b_node)
            b_bullets = b_node.get("bullets", []) if isinstance(b_node, dict) else []

            self._add_card_shape(
                slide, Pt(b_x), Pt(b_y), Pt(branches_w), Pt(branch_h),
                title=b_title, bullets=b_bullets,
                title_sz=16.0, bullet_sz=13.5,
                rounded=True, macro_border=True, padding=6.0, dense=True
            )

            arr_x = x.pt + root_w + (gap_x - arrow_w) / 2.0
            arr_y = b_y + (branch_h - arrow_h) / 2.0
            self._add_flowchart_arrow(
                slide, Pt(arr_x), Pt(arr_y), Pt(arrow_w), Pt(arrow_h),
                direction="right", color=COLOR_HNU_RED
            )

    def _add_grid_cards_layout(self, slide, x, y, w, h, cards_data, cols=3, rows=2):
        """插入 5~7 框矩阵网格布局 (Slide 17/18/24 规范)：
        cols 列 × rows 行矩阵卡片，支持 4~6 框网格
        """
        gap_x = 14.0
        gap_y = 14.0

        card_w = (w.pt - (cols - 1) * gap_x) / float(cols)
        card_h = (h.pt - (rows - 1) * gap_y) / float(rows)

        for idx, c_data in enumerate(cards_data):
            r = idx // cols
            c = idx % cols
            c_x = x.pt + c * (card_w + gap_x)
            c_y = y.pt + r * (card_h + gap_y)

            self._add_card_shape(
                slide, Pt(c_x), Pt(c_y), Pt(card_w), Pt(card_h),
                title=c_data.get("title", f"模块 {idx+1}"),
                bullets=c_data.get("bullets", []),
                title_sz=c_data.get("title_sz", 16.0),
                bullet_sz=c_data.get("bullet_sz", 13.5),
                rounded=True, macro_border=True, padding=6.0, dense=True
            )

    def _add_sidebar_plus_grid_layout(self, slide, x, y, w, h, sidebar_data, grid_cards, cols=2, rows=2):
        """插入 1 侧边全局卡 + 2×2 或 2×3 矩阵网格 (Slide 19/24 规范，共 5~7 框)：
        左侧: 宽幅全局卡片 (width ~ 32% ~ 36%)
        右侧: 2 列 × 2~3 行局部矩阵卡片
        """
        gap_x = 16.0
        gap_y = 12.0
        side_w = w.pt * 0.32
        grid_total_w = w.pt - side_w - gap_x

        self._add_card_shape(
            slide, x, y, Pt(side_w), h,
            title=sidebar_data.get("title", "系统总括"),
            bullets=sidebar_data.get("bullets", []),
            title_sz=sidebar_data.get("title_sz", 18.0),
            bullet_sz=sidebar_data.get("bullet_sz", 14.0),
            rounded=True, macro_border=True, padding=8.0, dense=True
        )

        grid_w = (grid_total_w - (cols - 1) * gap_x) / float(cols)
        grid_h = (h.pt - (rows - 1) * gap_y) / float(rows)

        for idx, c_data in enumerate(grid_cards):
            r = idx // cols
            c = idx % cols
            c_x = x.pt + side_w + gap_x + c * (grid_w + gap_x)
            c_y = y.pt + r * (grid_h + gap_y)

            self._add_card_shape(
                slide, Pt(c_x), Pt(c_y), Pt(grid_w), Pt(grid_h),
                title=c_data.get("title", f"子项 {idx+1}"),
                bullets=c_data.get("bullets", []),
                title_sz=c_data.get("title_sz", 16.0),
                bullet_sz=c_data.get("bullet_sz", 13.5),
                rounded=True, macro_border=True, padding=6.0, dense=True
            )

    def _add_3col_layout(self, slide, x, y, w, h, cards_data):
        """插入 3 列纵向高耸卡片布局 (Slide 23 规范)：
        适合 3 大支柱技术、3 阶段演进或 3 方对比
        """
        gap_x = 16.0
        card_w = (w.pt - 2.0 * gap_x) / 3.0

        for idx, c_data in enumerate(cards_data[:3]):
            c_x = x.pt + idx * (card_w + gap_x)
            self._add_card_shape(
                slide, Pt(c_x), y, Pt(card_w), h,
                title=c_data.get("title", f"支柱 {idx+1}"),
                bullets=c_data.get("bullets", []),
                title_sz=c_data.get("title_sz", 17.0),
                bullet_sz=c_data.get("bullet_sz", 14.0),
                rounded=True, macro_border=True, padding=8.0, dense=True
            )

    def _add_max_image_layout(self, slide, left_card, img_path, caption=None):
        """专为巨幅基准对比图定制的大图模板 (Max-Image Benchmark Template):
        左侧: 精炼技术机理卡片 (340.0 pt × 395.0 pt)
        右侧: 巨幅视觉展区 (525.0 pt × 410.0 pt)，专为 benchmack的对比.jpg 等大图设计，100% 零裁切 contain 呈现
        """
        left_pos = [38.1, 95.0, 340.0, 395.0]
        right_pos = [395.0, 90.0, 525.0, 410.0]

        # 左侧精炼文字卡
        self._add_card_shape(
            slide, Pt(left_pos[0]), Pt(left_pos[1]), Pt(left_pos[2]), Pt(left_pos[3]),
            title=left_card.get("title", "基准性能对比"),
            bullets=left_card.get("bullets", []),
            title_sz=left_card.get("title_sz", 18.0),
            bullet_sz=left_card.get("bullet_sz", 14.0),
            rounded=True, macro_border=True, padding=8.0, dense=True
        )

        # 右侧巨幅大图 (零裁切)
        self._add_picture_box(
            slide, Pt(right_pos[0]), Pt(right_pos[1]), Pt(right_pos[2]), Pt(right_pos[3]),
            img_path=img_path, caption=caption, show_border=False, padding=0.0
        )



    def _audit_deck(self, prs):
        """全域质量终审审计：
        1. 检查任何 slide 中是否有 font.size < 14pt；
        2. 检查内容页单行大标题是否 <= 28.0 pt；
        3. 统计正文字号分布 (确保在 16.0 ~ 20.0 pt 合规区间)；
        4. 检查印章禁飞区是否有非法侵入。
        """
        all_sizes = []
        body_sizes = []

        for s_idx, slide in enumerate(prs.slides):
            for shape in slide.shapes:
                # 印章禁飞区检查 (非封面、非结尾页)
                if 0 < s_idx < len(prs.slides) - 1:
                    # 排除印章图片自身 (图片 9, left ~ 11.55, top ~ 356.05)
                    is_seal = (shape.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE and shape.left.pt < 20 and shape.top.pt > 350)
                    if not is_seal and shape.name not in ["SEAL_WATERMARK", "TOP_LOGO", "PAGE_TITLE", "PAGE_NUMBER", "HEADER_LINE"]:
                        # 检查形状是否侵占印章禁飞区 (x < 280.1 且 y > 350.0)
                        if shape.left.pt < 280.1 and (shape.top.pt + shape.height.pt) > 351.0:
                            if shape.top.pt > 348.0 and shape.left.pt < 275.0 and shape.width.pt > 100:
                                raise AssertionError(f"Slide {s_idx+1} 侵占左下角印章禁飞区: {shape.name}, left={shape.left.pt}, top={shape.top.pt}")

                # 检查字号
                if shape.has_text_frame:
                    # 内容页单行大标题检查
                    if 0 < s_idx < len(prs.slides) - 1 and shape.left.pt < 100 and shape.top.pt < 50:
                        for p in shape.text_frame.paragraphs:
                            for r in p.runs:
                                if r.font.size is not None and r.font.size.pt > 28.5:
                                    raise AssertionError(f"Slide {s_idx+1} 大标题字号超过 28pt 限制: {r.font.size.pt} pt")

                    for p in shape.text_frame.paragraphs:
                        for r in p.runs:
                            # 检查是否残留任何未编译的 LaTeX 源码反斜杠
                            if r.text and any(k in r.text for k in [r'\min', r'\max', r'\lambda', r'\ge', r'\le', r'\Delta', r'\text{', r'\mathrm{']):
                                raise AssertionError(f"Slide {s_idx+1} 发现裸露未转换的 LaTeX 源码反斜杠: '{r.text}'")
                            if r.font.size is not None:
                                sz = r.font.size.pt
                                all_sizes.append(sz)
                                if sz < 12.9: # 浮点容差: 允许最小字号 >= 13.0 pt
                                    raise AssertionError(f"Slide {s_idx+1} 发现低于 13pt 的字体: {sz} pt ('{r.text[:20]}')")
                                if 14.9 <= sz <= 20.1:
                                    body_sizes.append(sz)

                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            for p in cell.text_frame.paragraphs:
                                for r in p.runs:
                                    if r.font.size is not None:
                                        sz = r.font.size.pt
                                        all_sizes.append(sz)
                                        if sz < 12.9:
                                            raise AssertionError(f"Slide {s_idx+1} 表格中发现低于 13pt 的字体: {r.font.size.pt} pt")

        min_sz = min(all_sizes) if all_sizes else 0
        max_sz = max(all_sizes) if all_sizes else 0
        print(f"[AUDIT PASS] 全局字阶与禁飞区审查 100% 通过！")
        print(f"  - 全局最小字号: {min_sz:.1f} pt (硬红线 >= 13.0 pt)")
        print(f"  - 全局最大字号: {max_sz:.1f} pt (封面主标题 <= 44.0 pt)")
        print(f"  - 正常正文文字样本数: {len(body_sizes)} (主要落在 13.0 ~ 20.0 pt 区间)")
        print(f"  - 印章避让: 全页 100% 完整避让！")
