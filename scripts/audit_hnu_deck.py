# -*- coding: utf-8 -*-
"""
湖南大学专属学术 PPT 质量门禁与底层审计脚本 (audit_hnu_deck.py) - v5.0 Formal Release (2026-09-23)
对生成的 .pptx 文件执行底层 OpenXML 深度审计与规范核验：
1. QG-01 原生可编辑性：非整页位图贴片
2. QG-02 导引框虚拟化：无残留蓝色多边形/矩形框
3. QG-03 字阶硬红线：最小字号 >= 13.0 pt（密集排版永久放宽至 13.0 pt），正文常规 14.0 ~ 18.0 pt，大标题 <= 28.0 pt
4. QG-04 普通中英文数字显式微软雅黑；数学变量字体单独核验
5. QG-05 封面红框保护：红色底框尺寸不变，主标题单行居中，绝无文字重叠
6. QG-06 禁区避让与非对称充实：印章禁飞区零侵入，右下横贯落地消除方形空白
7. QG-07 事实真实性：0 违禁词 (绝无 MINCO 等捏造内容)
8. QG-08 主题色保真：大标题使用 #A6232B
9. QG-09 图片数量统计；素材真实性与实装完整性须任务证据核对
10. QG-10 学术三线表配色：表头淡陶粉 #F8EAEB + 湖大红深色字 #8B1D23
11. QG-11 形状防重叠安全边界：卡片与表格 0 物理碰撞
12. QG-12 宽大主图视觉重心：主图高展宽阔，占比优于文字
13. QG-13 原生数学公式体：DrawingML 14 Office Math (<a14:m>)
14. QG-14 纯净边框准则：0.75pt 浅灰细线碎框彻底清零
15. QG-15 表格单元格内边距清零与垂直居中：marL/R/T/B=0, anchor='ctr'
16. QG-16 宽扁图容器长宽比适配：展区 AR 与图片天然 AR 偏差 <= 75%
17. QG-17 任务卫生需单独核验；只读审计不删除共享文件
18. QG-18 同一页内正文字号极差 <= 2.0 pt：严禁突兀大小混用
19. QG-19 图大文精视觉重心：图片为主视觉焦点，文字充实凝练
20. QG-20 卡片垂直空间有机饱满度：垂直利用率 67%~95%（宽方差人工呼吸感，底线 67%），杜绝空洞死白
21. QG-21 三层间距解耦与动态收紧：间距 <= 18pt，消除多重留白叠加
22. QG-22 仅检查显式当前项目资产目录与包内媒体哈希，来源/引用须人工复核
24. QG-24 原生项目符号、悬挂缩进和段距
25. QG-25 结构化目录识别、动态红蓝章节与几何
26. QG-26 显式功能页清单的 Skill 名称与具体作用标题绑定
27. QG-27 块箭头尺寸/连接箭头线宽；内容驱动布局仍需人工判断
28. QG-28 视觉节奏铁律：连续文字/表格页上限 <= 4 页，必有图表/流程图/视觉化呈现
29. QG-29 用户素材绝对零裁切铁律 (Zero-Crop Invariant)：严禁 PowerPoint 裁切属性
"""

import os
import sys
import zipfile
import argparse
import math
import re
import json
import hashlib
from pathlib import Path
import xml.etree.ElementTree as ET

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

FORBIDDEN_WORDS = [
    "MINCO", "minco",
    "CycloneDDS", "cyclonedds", "Cyclone",
    "18ms", "18 ms",
    "250Hz", "250 Hz",
    "1.5% 推力", "1.5%以内", "1.5% 以内"
]

# QG-23 反浮夸、反虚构与客观工程表述违禁词清单
HYPE_AND_FABRICATED_WORDS = [
    "无差拍", "全工况无差拍",
    "系统底噪极限", "底噪极限",
    "彻底消除",
    "零失真",
    "绝对优势", "绝对综合优势",
    "变电架构", "导线绝缘子",
    "颠覆性", "降维打击", "全天候无死角", "无死角"
]

def emu_to_pt(val):
    if val is None:
        return 0.0
    return round(float(val) / 12700.0, 2)

# Helpers merged into the shared auditor by the guarded repair script.
YAHEI = {'Microsoft YaHei', '微软雅黑'}
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

def shape_box(sp):
    xf = sp.find(f'.//{{{A_NS}}}xfrm')
    if xf is None: xf = sp.find(f'{{{P_NS}}}xfrm')
    if xf is None: return (0, 0, 0, 0)
    off, ext = xf.find(f'{{{A_NS}}}off'), xf.find(f'{{{A_NS}}}ext')
    if off is None or ext is None: return (0, 0, 0, 0)
    return tuple(emu_to_pt(v) for v in [off.get('x'), off.get('y'), ext.get('cx'), ext.get('cy')])

def shape_text(sp):
    return ''.join(t.text or '' for t in sp.iter(f'{{{A_NS}}}t')).strip()

def rgb_at(el, path):
    n = el.find(path)
    return n.get('val', '').upper() if n is not None else ''

def agenda_structure(tree):
    """Recognize geometry/content, never object names; return only agenda nodes.

    A malformed candidate receives QG25 errors and receives no QG06 exemption.
    Extra objects are deliberately absent from the exempt set.
    """
    shapes = tree.findall(f'.//{{{P_NS}}}sp')
    candidates = [s for s in shapes if shape_text(s).replace('\n','').strip() in ('汇报提纲', '提纲', '目录')
                  or (not shape_text(s).startswith('汇报完毕') and all(abs(a-b)<1 for a,b in zip(shape_box(s),(175.87,79.99,135.39,385.76))))]
    if not candidates: return set(), [], None
    errors=[]; title=candidates[0]; nodes={title}
    if len(candidates)!=1: errors.append('目录竖排标题数量不是 1')
    if not all(abs(a-b)<1 for a,b in zip(shape_box(title),(175.87,79.99,135.39,385.76))): errors.append('目录标题几何与 Slide 2 不符')
    for r in title.findall(f'.//{{{A_NS}}}r'):
        rp=r.find(f'{{{A_NS}}}rPr')
        if rp is None or rp.get('sz')!='6000' or rp.get('b')!='1': errors.append('目录标题必须为 60pt 加粗')
    nums=sorted([s for s in shapes if abs(shape_box(s)[0]-392.7)<1 and shape_text(s).isdigit()],key=lambda s:shape_box(s)[1])
    n=len(nums)
    if not 3<=n<=10: errors.append('目录必须包含 3~10 个章节')
    ys=[]; colors=[]; labels=[]
    for i,num in enumerate(nums):
        x,y,w,h=shape_box(num);ys.append(y);nodes.add(num)
        if shape_text(num)!=str(i+1): errors.append('目录序号不连续')
        bars=[s for s in shapes if abs(shape_box(s)[0]-(x+w+3.5))<1 and abs(shape_box(s)[1]-y)<1 and 3<=shape_box(s)[2]<=5]
        texts=[s for s in shapes if shape_text(s) and shape_box(s)[0]>x+w+10 and abs(shape_box(s)[1]-(y-2))<1]
        if len(bars)!=1 or len(texts)!=1: errors.append('目录序号/装饰条/章节文字不成组');continue
        bar,txt=bars[0],texts[0];nodes.update([bar,txt]);labels.append(shape_text(txt))
        c=rgb_at(num,f'{{{P_NS}}}spPr/{{{A_NS}}}solidFill/{{{A_NS}}}srgbClr');colors.append(c)
        tc=[rgb_at(r,f'{{{A_NS}}}rPr/{{{A_NS}}}solidFill/{{{A_NS}}}srgbClr') for r in txt.findall(f'.//{{{A_NS}}}r') if shape_text(r)]
        bc=rgb_at(bar,f'{{{P_NS}}}spPr/{{{A_NS}}}solidFill/{{{A_NS}}}srgbClr')
        if c not in ('A6232B','1D4999') or bc!=c or not tc or any(t!=c for t in tc): errors.append('目录当前红/其他蓝的序号、条与文字色不一致')
    if colors.count('A6232B')!=1: errors.append('目录必须恰有一个当前章节高亮')
    if len(ys)>1:
        steps=[b-a for a,b in zip(ys,ys[1:])]
        if max(steps)-min(steps)>1: errors.append('目录章节未等间距排列')
        if abs((ys[0]+ys[-1]+shape_box(nums[-1])[3])/2-270)>1: errors.append('目录章节整体未垂直居中')
    # An actual content-page seal/header makes this an invalid agenda, not an escape.
    if any(shape_box(s)[0]<280.1 and shape_box(s)[1]>350 for s in tree.findall(f'.//{{{P_NS}}}pic')): errors.append('目录混入内容页印章')
    return (nodes if not errors else set()), errors, {'labels':labels,'active':colors.index('A6232B') if colors.count('A6232B')==1 else None}

def check_updated_rules(tree, s_num, skill_names, issues):
    def fail(code,msg,severity='CRITICAL'): issues.append({'slide':s_num,'code':code,'severity':severity,'msg':msg})
    parent={child:node for node in tree.iter() for child in node}
    def under_math(node):
        while node in parent:
            node=parent[node]
            if node.tag in (f'{{{M_NS}}}oMath',f'{{{M_NS}}}oMathPara'):return True
        return False
    # Every displayed run, including table runs and fields; no silent font fallback.
    for tag in ('r','fld'):
        for r in tree.iter(f'{{{A_NS}}}{tag}'):
            t=r.find(f'{{{A_NS}}}t')
            if t is None or not (t.text or '').strip():continue
            txt=t.text.strip();rp=r.find(f'{{{A_NS}}}rPr')
            if rp is None: fail('QG-04',f'缺少显式字体属性: {txt[:35]}');continue
            lf=rp.find(f'{{{A_NS}}}latin');ef=rp.find(f'{{{A_NS}}}ea')
            latin=lf.get('typeface','') if lf is not None else '';ea=ef.get('typeface','') if ef is not None else ''
            # The specific cover link is a user-authorized project exception only.
            link_exception=(s_num==1 and txt=='https://github.com/w5711112/Unified-Scholarflow-Skills')
            # Inline math requires variable/operator syntax, not arbitrary English or digits.
            inline_math=(latin=='Cambria Math' and ((rp.get('i')=='1' and re.fullmatch(r'[A-Za-zα-ωΑ-Ω∂∇]',txt))
                         or (rp.get('baseline') is not None and re.fullmatch(r'[A-Za-z0-9α-ωΑ-Ω_]',txt))
                         or re.fullmatch(r'(?:Δ\s*[A-Za-z]|min|max|sin|cos|∂|∇|[α-ωΑ-Ω])',txt)))
            if not link_exception and not under_math(r) and not inline_math and (latin not in YAHEI or ea not in YAHEI):
                fail('QG-04',f'普通文字必须显式使用微软雅黑 (latin={latin}, ea={ea}): {txt[:35]}')
    for p in tree.iter(f'{{{A_NS}}}p'):
        txt=shape_text(p)
        if not txt:continue
        if re.match(r'^(?:•\s*|[-*]\s+)',txt):fail('QG-24',f'手工项目符号前缀: {txt[:35]}')
        pp=p.find(f'{{{A_NS}}}pPr')
        if pp is None:continue
        bu=pp.find(f'{{{A_NS}}}buChar');auto=pp.find(f'{{{A_NS}}}buAutoNum')
        if bu is None and auto is None:continue
        bf=pp.find(f'{{{A_NS}}}buFont')
        if bu is None or bu.get('char')!='•' or bf is None or bf.get('typeface')!='Microsoft YaHei UI' or pp.get('marL')!='228600' or pp.get('indent')!='-177800':
            fail('QG-24',f'项目符号/字体/悬挂缩进不符合原生规范: {txt[:35]}')
        for spacing in ('spcBef','spcAft'):
            sp=pp.find(f'{{{A_NS}}}{spacing}/{{{A_NS}}}spcPts')
            if sp is None or int(sp.get('val','0'))<400:fail('QG-24',f'列表 {spacing} 必须至少 4pt（默认 5pt）: {txt[:35]}')
    if skill_names:
        titles=[shape_text(s) for s in tree.findall(f'.//{{{P_NS}}}sp') if shape_box(s)[0]<100 and shape_box(s)[1]<50 and shape_text(s)]
        title=' '.join(titles).strip('【】')
        parts=re.split('[：:]',title,maxsplit=1)
        if len(parts)!=2 or not parts[1].strip() or any(not re.search(r'(?<![A-Za-z0-9_-])'+re.escape(name)+r'(?![A-Za-z0-9_-])',parts[0]) for name in skill_names):
            fail('QG-26',f'Skill 功能页标题未绑定 {skill_names} 与具体作用: {title}','HIGH')
    for sp in tree.findall(f'.//{{{P_NS}}}sp'):
        geom=sp.find(f'.//{{{A_NS}}}prstGeom')
        if geom is None or geom.get('prst') not in ('rightArrow','leftArrow','upArrow','downArrow'):continue
        _,_,w,h=shape_box(sp)
        if not (18<=w<=28 and 12<=h<=16):fail('QG-27',f'流程块箭头尺寸必须宽 18~28pt、高 12~16pt，实际 {w} × {h}','HIGH')
    for sp in tree.findall(f'.//{{{P_NS}}}cxnSp'):
        ln=sp.find(f'.//{{{A_NS}}}ln')
        if ln is None:continue
        ends=[ln.find(f'{{{A_NS}}}{t}') for t in ('headEnd','tailEnd')]
        if any(e is not None and e.get('type','none')!='none' for e in ends):
            width=emu_to_pt(ln.get('w','0'))
            if not 2<=width<=2.5:fail('QG-27',f'流程连接箭头线宽须为 2~2.5pt，实际 {width}','HIGH')


def audit_deck(pptx_path, asset_dir=None, skill_pages=None, json_report=None):
    print(f"==================================================")
    print(f"[AUDIT START] 正在启动湖南大学专属学术 PPT 质量门禁审查...")
    print(f"目标文件: {pptx_path}")
    print(f"==================================================")

    if not os.path.exists(pptx_path):
        print(f"[CRITICAL FAIL] 文件不存在: {pptx_path}")
        return False

    with zipfile.ZipFile(pptx_path, 'r') as z:
        # 按照 presentation.xml 的真实逻辑顺序解析 slide 列表
        pres_xml = z.read('ppt/presentation.xml').decode('utf-8')
        rels_xml = z.read('ppt/_rels/presentation.xml.rels').decode('utf-8')
        pres_tree = ET.fromstring(pres_xml)
        rels_tree = ET.fromstring(rels_xml)

        rel_map = {}
        for rel in rels_tree:
            rel_map[rel.attrib['Id']] = rel.attrib['Target']

        sld_id_lst = pres_tree.find(f'.//{{{P_NS}}}sldIdLst')
        ordered_slide_files = []
        for sldId in sld_id_lst:
            rId = sldId.attrib[f'{{{R_NS}}}id']
            target = rel_map[rId]
            # target 可能是 'slides/slide1.xml' 或 'slide1.xml'
            if not target.startswith('ppt/'):
                target = 'ppt/' + target.lstrip('/')
            ordered_slide_files.append(target)

        total_slides = len(ordered_slide_files)
        print(f"检测到幻灯片逻辑总页数: {total_slides} 页")

        issues = []
        all_font_sizes = []
        body_font_sizes = []
        all_fonts = set()
        ea_fonts = set()
        total_real_images = 0
        total_tables = 0
        agenda_pages = []
        agenda_details = []
        visual_status_by_slide = {}
        skill_pages = {int(k): v for k, v in (skill_pages or {}).items()}

        for idx, s_file in enumerate(ordered_slide_files):
            s_num = idx + 1 # 真实逻辑页码 1-indexed
            xml_content = z.read(s_file).decode('utf-8')
            tree = ET.fromstring(xml_content)

            # 官方母版结束页跳过内容页规则检查
            if s_num == total_slides and '汇报完毕' in xml_content:
                continue

            slide_body_sizes = []
            agenda_nodes, agenda_errors, agenda_info = agenda_structure(tree)
            is_agenda = bool(agenda_nodes)
            if is_agenda:
                agenda_pages.append(s_num)
                agenda_details.append(agenda_info)
            for message in agenda_errors:
                issues.append({'slide': s_num, 'code': 'QG-25', 'severity': 'CRITICAL', 'msg': message})
            check_updated_rules(tree, s_num, skill_pages.get(s_num), issues)

            # 检查违禁词 (QG-07 事实真实性)
            for word in FORBIDDEN_WORDS:
                if word in xml_content:
                    issues.append({
                        "slide": s_num, "code": "QG-07", "severity": "CRITICAL",
                        "msg": f"发现违禁/虚构术语: '{word}'，违反事实真实性契约！"
                    })

            # 检查夸大修饰词与虚构实体 (QG-23 反浮夸反虚构)
            for word in HYPE_AND_FABRICATED_WORDS:
                if word in xml_content:
                    issues.append({
                        "slide": s_num, "code": "QG-23", "severity": "CRITICAL",
                        "msg": f"发现违禁夸大修饰/伪学术黑话/虚构实体: '{word}'，违反客观求真与事实真实性契约！"
                    })

            # 检查图片 (QG-01 & QG-09 & QG-29)
            pics = tree.findall(f'.//{{{P_NS}}}pic')
            user_pics = []
            for p in pics:
                p_box = shape_box(p)
                # 排除右上校徽 (x > 700, y < 60) 和左下水印印章 (x < 285, y > 340)
                if (p_box[0] > 700 and p_box[1] < 60) or (p_box[0] < 285 and p_box[1] > 340):
                    continue
                user_pics.append(p)

            total_real_images += len(pics)
            texts = [t.text for t in tree.findall(f'.//{{{A_NS}}}t') if t.text and t.text.strip()]
            if s_num > 1 and s_num < total_slides and len(pics) == 1 and len(texts) <= 1:
                issues.append({
                    "slide": s_num, "code": "QG-01", "severity": "CRITICAL",
                    "msg": "疑似整页位图贴片，缺失原生可编辑对象！"
                })

            # QG-29: 用户素材与工程原图绝对零裁切铁律 (Zero-Crop Invariant)
            if 1 < s_num < total_slides:
                for p in user_pics:
                    blipFill = p.find(f'.//{{{A_NS}}}blipFill')
                    if blipFill is not None:
                        srcRect = blipFill.find(f'{{{A_NS}}}srcRect')
                        if srcRect is not None:
                            crops = {k: srcRect.get(k) for k in ('l', 't', 'r', 'b') if srcRect.get(k) not in (None, '0', 0)}
                            if crops:
                                p_cNvPr = p.find(f'.//{{{P_NS}}}cNvPr')
                                p_name = p_cNvPr.get('name', 'Picture') if p_cNvPr is not None else 'Picture'
                                issues.append({
                                    "slide": s_num, "code": "QG-29", "severity": "CRITICAL",
                                    "msg": f"图片 '{p_name}' 存在 PowerPoint 裁切属性 (srcRect={crops})，违反用户素材绝对零裁切 (Zero-Crop) 铁律！必须保持 100% 原图完整呈现 (使用 contain 等比缩放)！"
                                })

            # QG-28: 统计当前页是否为视觉页 (有图、有公式、有流程箭头、或者是目录页)
            has_pics = len(user_pics) > 0
            has_math = ('<a14:m' in xml_content or ':m' in xml_content or 'oMath' in xml_content)
            has_flow_arrow = False
            for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                geom = sp.find(f'.//{{{A_NS}}}prstGeom')
                if geom is not None and geom.get('prst') in ('rightArrow', 'leftArrow', 'upArrow', 'downArrow'):
                    has_flow_arrow = True
                    break
            if not has_flow_arrow:
                for cxn in tree.findall(f'.//{{{P_NS}}}cxnSp'):
                    ln = cxn.find(f'.//{{{A_NS}}}ln')
                    if ln is not None and any(ln.find(f'{{{A_NS}}}{t}') is not None for t in ('headEnd', 'tailEnd')):
                        has_flow_arrow = True
                        break
            visual_status_by_slide[s_num] = has_pics or has_math or has_flow_arrow or is_agenda

            # 遍历形状
            for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                cNvPr = sp.find(f'.//{{{P_NS}}}cNvPr')
                name = cNvPr.attrib.get('name', '') if cNvPr is not None else ''

                # QG-02: 检查残留参考多边形
                cust = sp.find(f'.//{{{A_NS}}}custGeom')
                if '任意多边形' in name or (cust is not None and s_num > 1 and s_num < total_slides):
                    issues.append({
                        "slide": s_num, "code": "QG-02", "severity": "CRITICAL",
                        "msg": f"残留未清除的参考多边形导引框: {name}"
                    })

                # 获取形状位置
                xfrm = sp.find(f'.//{{{A_NS}}}xfrm')
                x_pt, y_pt, w_pt, h_pt = 0.0, 0.0, 0.0, 0.0
                if xfrm is not None:
                    off = xfrm.find(f'{{{A_NS}}}off')
                    ext = xfrm.find(f'{{{A_NS}}}ext')
                    if off is not None:
                        x_pt = emu_to_pt(off.attrib.get('x', 0))
                        y_pt = emu_to_pt(off.attrib.get('y', 0))
                    if ext is not None:
                        w_pt = emu_to_pt(ext.attrib.get('cx', 0))
                        h_pt = emu_to_pt(ext.attrib.get('cy', 0))

                # QG-06: 印章禁飞区避让检查 (仅针对内容页: 1 < s_num < total_slides)
                if 1 < s_num < total_slides:
                    # 排除全幅背景底板、顶部红线、大标题等系统母版形状
                    if sp not in agenda_nodes and x_pt < 280.1 and (y_pt + h_pt) > 348.0:
                        issues.append({
                            "slide": s_num, "code": "QG-06", "severity": "CRITICAL",
                            "msg": f"形状侵入左下角印章禁飞区: {name} (x={x_pt}, y={y_pt}, w={w_pt}, h={h_pt}, 底部y={y_pt+h_pt} > 348.0)"
                        })

                # QG-05: 封面第 1 页红色大底框与标题长度检查
                if s_num == 1:
                    if abs(x_pt - 143.55) < 10 and abs(y_pt - 361.75) < 10 and w_pt > 700:
                        if w_pt < 800.0 or h_pt < 170.0:
                            issues.append({
                                "slide": 1, "code": "QG-05", "severity": "CRITICAL",
                                "msg": f"封面红色大底框尺寸发生缩小违规: 宽度={w_pt} pt (标准 816.45 pt)"
                            })
                    # 检查封面标题是否单行且字数 <= 17 字
                    if y_pt > 150.0 and y_pt < 260.0 and x_pt > 180.0:
                        title_texts = "".join([t.text for t in sp.findall(f'.//{{{A_NS}}}t') if t.text])
                        if len(title_texts) > 17:
                            issues.append({
                                "slide": 1, "code": "QG-05", "severity": "CRITICAL",
                                "msg": f"封面主标题字数 ({len(title_texts)}字) 超过17字硬上限，存在溢出右边缘风险: '{title_texts}'"
                            })

                # QG-03 & QG-04 & QG-12: 字号、字体与数学公式反斜杠检查
                for r in sp.findall(f'.//{{{A_NS}}}r'):
                    rPr = r.find(f'{{{A_NS}}}rPr')
                    t = r.find(f'{{{A_NS}}}t')
                    if t is not None and t.text and t.text.strip():
                        # QG-12: 检查裸露 LaTeX 源码反斜杠
                        for bad_tex in [r'\min', r'\max', r'\lambda', r'\ge', r'\le', r'\Delta', r'\text{', r'\mathrm{', r'\mathbf{']:
                            if bad_tex in t.text:
                                issues.append({
                                    "slide": s_num, "code": "QG-12", "severity": "CRITICAL",
                                    "msg": f"发现裸露未转换的 LaTeX 源码反斜杠 '{bad_tex}': '{t.text.strip()}'"
                                })

                        # QG-39: 检查 renhua 违禁 AI 腔套话与标点 (R1/R7/R12)
                        for bad_cliche in ['不仅如此', '显著提升', '显著提高', '全面赋能', '极大地赋能', '深度赋能', '划时代的', '具有重要意义', '毋庸置疑', '毫无疑问', '未来可期', '必将大放异彩', '旨在打造', '致力于打造', '作为 AI', '正如我们所见', '——']:
                            if bad_cliche in t.text:
                                issues.append({
                                    "slide": s_num, "code": "QG-39", "severity": "CRITICAL",
                                    "msg": f"发现违背 renhua 规范的空洞 AI 腔或违规标点 '{bad_cliche}': '{t.text.strip()}'！必须改写为客观、平实的工科白描语言！"
                                })

                        # QG-41: 零 Markdown 字符泄漏检查 (大忌)
                        if '**' in t.text or '__' in t.text or '`' in t.text or re.search(r'^\s*#{1,6}\s+', t.text):
                            issues.append({
                                "slide": s_num, "code": "QG-41", "severity": "CRITICAL",
                                "msg": f"发现裸露未解析的 Markdown 控制字符 ('**' / '__' / '`' / '#'): '{t.text.strip()}'！加粗必须使用 OpenXML 原生 run.font.bold 属性，严禁 Markdown 字符直接泄漏到 PPT 文本中！"
                            })

                        if rPr is not None and 'sz' in rPr.attrib:
                            sz_pt = float(rPr.attrib['sz']) / 100.0
                            all_font_sizes.append(sz_pt)

                            if sz_pt < 12.9:
                                issues.append({
                                    "slide": s_num, "code": "QG-03", "severity": "CRITICAL",
                                    "msg": f"发现低于 13.0 pt 的微小字体违规: {sz_pt} pt ('{t.text.strip()[:15]}')"
                                })

                            if 1 < s_num < total_slides and x_pt < 100.0 and y_pt < 50.0:
                                if sz_pt > 28.5:
                                    issues.append({
                                        "slide": s_num, "code": "QG-03", "severity": "HIGH",
                                        "msg": f"内容页大标题超过 28.0 pt 上限: {sz_pt} pt ('{t.text.strip()[:15]}')"
                                    })

                            if 13.0 <= sz_pt <= 20.1:
                                body_font_sizes.append(sz_pt)

                        latin = rPr.find(f'{{{A_NS}}}latin') if rPr is not None else None
                        ea = rPr.find(f'{{{A_NS}}}ea') if rPr is not None else None
                        if latin is not None:
                            all_fonts.add(latin.attrib.get('typeface', ''))
                        if ea is not None:
                            ea_fonts.add(ea.attrib.get('typeface', ''))

            # 提取真正的卡片正文字号 (QG-18: 严格针对卡片正文要点字阶收敛，排除卡片标题与图片图注)
            card_body_sizes = []
            if 1 < s_num < total_slides and not is_agenda:
                for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                    xfrm = sp.find(f'.//{{{A_NS}}}xfrm')
                    txBody = sp.find(f'.//{{{P_NS}}}txBody')
                    if xfrm is None or txBody is None:
                        continue
                    off = xfrm.find(f'{{{A_NS}}}off')
                    ext = xfrm.find(f'{{{A_NS}}}ext')
                    if off is None or ext is None:
                        continue
                    sp_x = emu_to_pt(off.attrib.get('x', 0))
                    sp_y = emu_to_pt(off.attrib.get('y', 0))
                    sp_w = emu_to_pt(ext.attrib.get('cx', 0))
                    sp_h = emu_to_pt(ext.attrib.get('cy', 0))
                    if sp_y < 80.0 or sp_w >= 850.0 or sp_h < 40.0:
                        continue

                    paras = txBody.findall(f'.//{{{A_NS}}}p')
                    for p_idx, p in enumerate(paras):
                        p_txt = "".join([t.text for t in p.findall(f'.//{{{A_NS}}}t') if t.text]).strip()
                        if not p_txt:
                            continue
                        if p_txt.startswith('图 ') or p_txt.startswith('图1') or p_txt.startswith('图2') or p_txt.startswith('图3') or p_txt.startswith('图4'):
                            continue
                        if p_idx == 0 and len(paras) > 1 and p.find(f'{{{A_NS}}}pPr/{{{A_NS}}}buChar') is None:
                            continue

                        for rPr in p.findall(f'.//{{{A_NS}}}rPr'):
                            if 'sz' in rPr.attrib:
                                r_sz = float(rPr.attrib['sz']) / 100.0
                                if 'baseline' not in rPr.attrib:
                                    card_body_sizes.append(r_sz)

            # QG-18: 同一页内正文字号极差必须 <= 2.0 pt (严禁 19pt 与 14pt 混用)
            if 1 < s_num < total_slides and not is_agenda and len(card_body_sizes) >= 2:
                body_min = min(card_body_sizes)
                body_max = max(card_body_sizes)
                body_range = body_max - body_min
                if body_range > 2.05: # 浮点容差
                    issues.append({
                        "slide": s_num, "code": "QG-18", "severity": "CRITICAL",
                        "msg": f"同一页内正文字号差距过大 (极差 {body_range:.1f} pt > 2.0 pt，最大={body_max} pt, 最小={body_min} pt)！严重破坏页面协调感！"
                    })

            # 表格检查 (QG-10 draw-style 淡雅配色 & QG-15 0边距垂直居中)
            tables = tree.findall(f'.//{{{A_NS}}}tbl')
            total_tables += len(tables)
            for tbl in tables:
                # 检查是否残留 1E293B 沉重黑块
                for clr in tbl.findall(f'.//{{{A_NS}}}srgbClr'):
                    c_val = clr.attrib.get('val', '').upper()
                    if c_val == '1E293B':
                        issues.append({
                            "slide": s_num, "code": "QG-10", "severity": "CRITICAL",
                            "msg": "表格表头使用了违禁的沉重深蓝黑底板 (#1E293B)！必须遵循 draw-style 淡雅浅色规范！"
                        })
                # QG-15: 检查单元格内边距与垂直对齐
                for tc in tbl.findall(f'.//{{{A_NS}}}tc'):
                    tcPr = tc.find(f'{{{A_NS}}}tcPr')
                    if tcPr is not None:
                        marL = tcPr.attrib.get('marL', None)
                        marR = tcPr.attrib.get('marR', None)
                        marT = tcPr.attrib.get('marT', None)
                        marB = tcPr.attrib.get('marB', None)
                        anchor = tcPr.attrib.get('anchor', None)
                        # 若未清零边距 (marL/R/T/B 不为 0)
                        if marL not in ('0', 0, None) or marR not in ('0', 0, None) or marT not in ('0', 0, None) or marB not in ('0', 0, None):
                            # 如果值大于 0 则报警
                            if any(int(m or 0) > 0 for m in (marL, marR, marT, marB)):
                                issues.append({
                                    "slide": s_num, "code": "QG-15", "severity": "CRITICAL",
                                    "msg": f"表格单元格内边距未清零 (marL={marL}, marT={marT})！必须设置 0 边距保证内容舒展！"
                                })
                                break
                        if anchor != 'ctr':
                            issues.append({
                                "slide": s_num, "code": "QG-15", "severity": "CRITICAL",
                                "msg": f"表格单元格垂直对齐方式不是垂直居中 (anchor='{anchor}')！必须设置 anchor='ctr'！"
                            })
                            break
                        # QG-15: 检查单元格内段落是否水平居中对齐 (包含第一列)
                        for p in tc.findall(f'.//{{{A_NS}}}p'):
                            pPr = p.find(f'{{{A_NS}}}pPr')
                            algn = pPr.attrib.get('algn', '') if pPr is not None else ''
                            if algn != 'ctr':
                                issues.append({
                                    "slide": s_num, "code": "QG-15", "severity": "CRITICAL",
                                    "msg": f"表格单元格水平对齐未居中 (algn='{algn}')！表格第一列及所有列文字必须强制水平居中对齐！"
                                })
                                break

                for r in tbl.findall(f'.//{{{A_NS}}}r'):
                    rPr = r.find(f'{{{A_NS}}}rPr')
                    t = r.find(f'{{{A_NS}}}t')
                    if rPr is not None and t is not None and t.text and t.text.strip():
                        if 'sz' in rPr.attrib:
                            sz_pt = float(rPr.attrib['sz']) / 100.0
                            all_font_sizes.append(sz_pt)
                            if sz_pt < 12.9:
                                issues.append({
                                    "slide": s_num, "code": "QG-03", "severity": "CRITICAL",
                                    "msg": f"表格单元格字号低于 13.0 pt: {sz_pt} pt ('{t.text.strip()[:15]}')"
                                })
                        latin = rPr.find(f'{{{A_NS}}}latin') if rPr is not None else None
                        ea = rPr.find(f'{{{A_NS}}}ea') if rPr is not None else None
                        if latin is not None:
                            all_fonts.add(latin.attrib.get('typeface', ''))
                        if ea is not None:
                            ea_fonts.add(ea.attrib.get('typeface', ''))

            # QG-13: 原生数学公式体检查 (<a14:m>)
            if 'B样条' in xml_content or '轨迹优化' in xml_content or 'J_s' in xml_content:
                has_a14m = '<a14:m' in xml_content or ':m' in xml_content or 'oMath' in xml_content
                if not has_a14m:
                    issues.append({
                        "slide": s_num, "code": "QG-13", "severity": "CRITICAL",
                        "msg": "核心轨迹优化数学公式页未检测到原生 DrawingML 14 Office Math (<a14:m>) 结构！"
                    })
                # 检查公式中文字颜色是否显式注入深灰 (防白底白字隐形)
                if has_a14m and '242424' not in xml_content and 'solidFill' not in xml_content:
                    issues.append({
                        "slide": s_num, "code": "QG-13", "severity": "CRITICAL",
                        "msg": "原生公式未显式指定深色字体颜色，存在继承母版默认白色导致隐形的重大风险！"
                    })

            # QG-14: 细线碎框检查 (0.75pt #E2E8F0 拦截)
            for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                ln = sp.find(f'.//{{{A_NS}}}ln')
                if ln is not None:
                    w = ln.attrib.get('w', '')
                    srgb = ln.find(f'.//{{{A_NS}}}srgbClr')
                    clr_val = srgb.attrib.get('val', '').upper() if srgb is not None else ''
                    # 9525 EMU 对应 0.75 pt, E2E8F0 为常见默认细描边
                    if w == '9525' and clr_val in ('E2E8F0', 'CBD5E1'):
                        issues.append({
                            "slide": s_num, "code": "QG-14", "severity": "HIGH",
                            "msg": f"发现组件带有 0.75pt 浅灰细线碎框 ({clr_val})，违反纯净无边框或宏观大框包裹准则！"
                        })

            # QG-16: 宽扁图容器长宽比适配检查 (以 Slide 9 软件链路图为标杆)
            if '3-18' in xml_content or '软件链路' in xml_content:
                # 检查宽扁图卡片宽度是否 >= 350 pt
                for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                    sp_text = "".join([t.text for t in sp.findall(f'.//{{{A_NS}}}t') if t.text])
                    if '软件链路' in sp_text or '链路' in sp_text:
                        xfrm = sp.find(f'.//{{{A_NS}}}xfrm')
                        if xfrm is not None:
                            ext = xfrm.find(f'{{{A_NS}}}ext')
                            if ext is not None:
                                card_w = emu_to_pt(ext.attrib.get('cx', 0))
                                if card_w < 340.0:
                                    issues.append({
                                        "slide": s_num, "code": "QG-16", "severity": "CRITICAL",
                                        "msg": f"宽扁软件链路图 (AR=2.4) 卡片展区过窄 ({card_w} pt < 350 pt)，导致图片压缩严重并产生大面积死白！"
                                    })

            # QG-11: 形状防重叠安全边界检算 (重点针对 Slide 14 等多块页面)
            slide_boxes = []
            for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                xfrm = sp.find(f'.//{{{A_NS}}}xfrm')
                if xfrm is not None:
                    off = xfrm.find(f'{{{A_NS}}}off')
                    ext = xfrm.find(f'{{{A_NS}}}ext')
                    if off is not None and ext is not None:
                        bx = emu_to_pt(off.attrib.get('x', 0))
                        by = emu_to_pt(off.attrib.get('y', 0))
                        bw = emu_to_pt(ext.attrib.get('cx', 0))
                        bh = emu_to_pt(ext.attrib.get('cy', 0))
                        # 排除全屏背景、标题和校徽水印
                        if bw < 850 and by > 80:
                            slide_boxes.append((bx, by, bw, bh))

            # 表格包围盒
            for tbl_sp in tree.findall(f'.//{{{A_NS}}}tbl'):
                # 找到所属 graphicFrame
                gf = tree.find(f'.//{{{P_NS}}}graphicFrame')
                if gf is not None:
                    xfrm = gf.find(f'.//{{{P_NS}}}xfrm')
                    if xfrm is not None:
                        off = xfrm.find(f'{{{A_NS}}}off')
                        ext = xfrm.find(f'{{{A_NS}}}ext')
                        if off is not None and ext is not None:
                            tbx = emu_to_pt(off.attrib.get('x', 0))
                            tby = emu_to_pt(off.attrib.get('y', 0))
                            tbw = emu_to_pt(ext.attrib.get('cx', 0))
                            tbh = emu_to_pt(ext.attrib.get('cy', 0))
                            # 检算表格与卡片是否物理碰撞重叠
                            for (bx, by, bw, bh) in slide_boxes:
                                x_overlap = max(0.0, min(tbx + tbw, bx + bw) - max(tbx, bx))
                                y_overlap = max(0.0, min(tby + tbh, by + bh) - max(tby, by))
                                if x_overlap > 10.0 and y_overlap > 10.0:
                                    issues.append({
                                        "slide": s_num, "code": "QG-11", "severity": "CRITICAL",
                                        "msg": f"发现表格与内容卡片严重物理重叠碰撞！(X重叠 {x_overlap:.1f}pt, Y重叠 {y_overlap:.1f}pt)"
                                    })

            # QG-20: 卡片垂直空间高充实度检查 (卡片垂直空间利用率必须 >= 80%, 杜绝下半部死白留白)
            if 1 < s_num < total_slides and not is_agenda:
                for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                    xfrm = sp.find(f'.//{{{A_NS}}}xfrm')
                    txBody = sp.find(f'.//{{{P_NS}}}txBody')
                    if xfrm is None or txBody is None:
                        continue
                    off = xfrm.find(f'{{{A_NS}}}off')
                    ext = xfrm.find(f'{{{A_NS}}}ext')
                    if off is None or ext is None:
                        continue
                    by = emu_to_pt(off.attrib.get('y', 0))
                    bw = emu_to_pt(ext.attrib.get('cx', 0))
                    bh = emu_to_pt(ext.attrib.get('cy', 0))
                    
                    # 仅检查正文内容卡片 (排除顶部页面大标题、母版水印、全屏背景)
                    if by < 85.0 or bw >= 850.0 or bh < 80.0:
                        continue
                    
                    paras = txBody.findall(f'.//{{{A_NS}}}p')
                    non_empty_p = []
                    for p in paras:
                        txt = "".join([t.text for t in p.findall(f'.//{{{A_NS}}}t') if t.text]).strip()
                        p_xml_str = ET.tostring(p).decode('utf-8')
                        is_math = (':m' in p_xml_str or 'oMath' in p_xml_str)
                        if txt or is_math:
                            non_empty_p.append((p, txt, is_math))
                    if not non_empty_p:
                        continue
                    
                    bodyPr = txBody.find(f'{{{A_NS}}}bodyPr')
                    l_ins = emu_to_pt(bodyPr.attrib.get('lIns', 101600)) if bodyPr is not None else 8.0
                    r_ins = emu_to_pt(bodyPr.attrib.get('rIns', 101600)) if bodyPr is not None else 8.0
                    t_ins = emu_to_pt(bodyPr.attrib.get('tIns', 76200)) if bodyPr is not None else 6.0
                    b_ins = emu_to_pt(bodyPr.attrib.get('bIns', 76200)) if bodyPr is not None else 6.0
                    avail_w = max(20.0, bw - (l_ins + r_ins))
                    avail_h = max(20.0, bh - (t_ins + b_ins))
                    if avail_h <= 10.0 or avail_w <= 10.0:
                        continue
                    
                    # 检算该卡片内文本的实际渲染占用高度
                    rendered_h = 0.0
                    for p_elem, p_txt, is_math in non_empty_p:
                        p_sz = 14.0
                        for rPr in p_elem.findall(f'.//{{{A_NS}}}rPr'):
                            if 'sz' in rPr.attrib:
                                p_sz = float(rPr.attrib['sz']) / 100.0
                                break
                        
                        p_spc = 0.0
                        p_before = 0.0
                        p_margin = 0.0
                        p_ln_spc = 1.25
                        p_ln_pts = None
                        pPr = p_elem.find(f'{{{A_NS}}}pPr')
                        if pPr is not None:
                            before = pPr.find(f'{{{A_NS}}}spcBef/{{{A_NS}}}spcPts')
                            p_before = float(before.get('val', '0')) / 100.0 if before is not None else 0.0
                            p_margin = emu_to_pt(pPr.get('marL', '0'))
                            spcAft = pPr.find(f'{{{A_NS}}}spcAft')
                            if spcAft is not None:
                                spcPts = spcAft.find(f'{{{A_NS}}}spcPts')
                                if spcPts is not None and 'val' in spcPts.attrib:
                                    p_spc = float(spcPts.attrib['val']) / 100.0
                            lnSpc = pPr.find(f'{{{A_NS}}}lnSpc')
                            if lnSpc is not None:
                                fixed = lnSpc.find(f'{{{A_NS}}}spcPts')
                                if fixed is not None and 'val' in fixed.attrib:
                                    p_ln_pts = float(fixed.attrib['val']) / 100.0
                                spcPct = lnSpc.find(f'{{{A_NS}}}spcPct')
                                if spcPct is not None and 'val' in spcPct.attrib:
                                    p_ln_spc = float(spcPct.attrib['val']) / 100000.0
                        
                        # 原生数学公式段落高度
                        if is_math and not p_txt:
                            p_h = 32.0 + p_spc
                            rendered_h += p_h
                            continue

                        # Explicit soft breaks are real line boundaries. Long segments
                        # still use the existing width-based automatic-wrap estimate.
                        segments = ['']
                        for child in p_elem:
                            if child.tag == f'{{{A_NS}}}br':
                                segments.append('')
                            elif child.tag in (f'{{{A_NS}}}r', f'{{{A_NS}}}fld'):
                                segments[-1] += ''.join(t.text or '' for t in child.iter(f'{{{A_NS}}}t'))
                        segments = [part for segment in segments for part in segment.split('\n')]
                        chars_per_line = max(1.0, max(1.0, avail_w - p_margin) / p_sz)
                        line_count = 0
                        for segment in segments:
                            equiv_chars = sum(1.0 if ('\u4e00' <= ch <= '\u9fff' or ch in '，。；：！？（）“”《》') else 0.55 for ch in segment)
                            if bodyPr is not None and bodyPr.get('wrap') == 'none':
                                line_count += 1  # Horizontal fit still needs rendered QA.
                            else:
                                line_count += max(1, math.ceil(max(0.0, equiv_chars - 0.25) / chars_per_line))
                        line_height = p_ln_pts if p_ln_pts is not None else p_sz * p_ln_spc
                        p_h = line_count * line_height + p_spc + p_before
                        rendered_h += p_h
                    
                    fill_rate = rendered_h / avail_h
                    card_title = non_empty_p[0][1][:12] if non_empty_p[0][1] else "MathCard"
                    if fill_rate < 0.60:
                        issues.append({
                            "slide": s_num, "code": "QG-20", "severity": "CRITICAL",
                            "msg": f"卡片 '{card_title}' 垂直空间填充率过低 ({fill_rate*100:.1f}% < 60%)！存在大面积留白死白，需增加分点/充实技术细节/调整行距至 1.25x！"
                        })
                    elif fill_rate > 1.08:
                        issues.append({
                            "slide": s_num, "code": "QG-20", "severity": "CRITICAL",
                            "msg": f"卡片 '{card_title}' 垂直空间填充率溢出 ({fill_rate*100:.1f}% > 108%)！文字超出卡片物理边界！"
                        })

            # QG-20 / QG-30: 底托卡片物理边界不可逾越检算 (Physical Card Boundary Confinement)
            bg_card_rects = []
            if 1 < s_num < total_slides and not is_agenda:
                for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                    xfrm = sp.find(f'.//{{{A_NS}}}xfrm')
                    if xfrm is None: continue
                    off = xfrm.find(f'{{{A_NS}}}off'); ext = xfrm.find(f'{{{A_NS}}}ext')
                    if off is None or ext is None: continue
                    bx = emu_to_pt(off.attrib.get('x', 0)); by = emu_to_pt(off.attrib.get('y', 0))
                    bw = emu_to_pt(ext.attrib.get('cx', 0)); bh = emu_to_pt(ext.attrib.get('cy', 0))
                    cNvPr = sp.find(f'.//{{{P_NS}}}cNvPr')
                    s_name = cNvPr.attrib.get('name', '') if cNvPr is not None else ''
                    if by >= 85.0 and 80.0 <= bw < 850.0 and bh >= 50.0:
                        has_fill = sp.find(f'.//{{{P_NS}}}spPr/{{{A_NS}}}solidFill') is not None
                        txBody = sp.find(f'.//{{{P_NS}}}txBody')
                        has_text = False
                        if txBody is not None:
                            has_text = bool("".join(t.text or "" for t in txBody.iter(f'{{{A_NS}}}t')).strip())
                        if s_name == 'BG' or (has_fill and not has_text):
                            bg_card_rects.append((bx, by, bw, bh, s_name))

                if bg_card_rects:
                    for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                        xfrm = sp.find(f'.//{{{A_NS}}}xfrm')
                        txBody = sp.find(f'.//{{{P_NS}}}txBody')
                        if xfrm is None or txBody is None: continue
                        off = xfrm.find(f'{{{A_NS}}}off'); ext = xfrm.find(f'{{{A_NS}}}ext')
                        if off is None or ext is None: continue
                        tx = emu_to_pt(off.attrib.get('x', 0)); ty = emu_to_pt(off.attrib.get('y', 0))
                        tw = emu_to_pt(ext.attrib.get('cx', 0)); th = emu_to_pt(ext.attrib.get('cy', 0))
                        cNvPr = sp.find(f'.//{{{P_NS}}}cNvPr')
                        s_name = cNvPr.attrib.get('name', '') if cNvPr is not None else ''
                        if ty < 85.0 or tw >= 850.0 or th < 20.0:
                            continue
                        matched_bg = None
                        for (bx, by, bw, bh, bg_name) in bg_card_rects:
                            if bx - 5.0 <= tx and tx + tw <= bx + bw + 5.0 and by - 5.0 <= ty <= by + bh + 5.0:
                                matched_bg = (bx, by, bw, bh, bg_name)
                                break
                        if matched_bg:
                            card_bottom = matched_bg[1] + matched_bg[3]
                            tx_bottom = ty + th
                            if tx_bottom > card_bottom + 2.0:
                                issues.append({
                                    "slide": s_num, "code": "QG-20", "severity": "CRITICAL",
                                    "msg": f"文本框 '{s_name}' (top={ty:.1f}, h={th:.1f}, 底部={tx_bottom:.1f}) 穿透底托卡片物理边界 (卡片底部={card_bottom:.1f})，超出 {tx_bottom - card_bottom:.1f} pt！存在严重穿模压盖事故！"
                                })

            # QG-21: 三层间距解耦与动态收紧检查 (并排构件框间距 <= 18.0 pt，杜绝 30~40pt 冗余空旷间隔)
            if 1 < s_num < total_slides:
                top_row_boxes = []
                for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                    xfrm = sp.find(f'.//{{{A_NS}}}xfrm')
                    if xfrm is not None:
                        off = xfrm.find(f'{{{A_NS}}}off')
                        ext = xfrm.find(f'{{{A_NS}}}ext')
                        if off is not None and ext is not None:
                            bx = emu_to_pt(off.attrib.get('x', 0))
                            by = emu_to_pt(off.attrib.get('y', 0))
                            bw = emu_to_pt(ext.attrib.get('cx', 0))
                            bh = emu_to_pt(ext.attrib.get('cy', 0))
                            if 85.0 <= by <= 110.0 and bw < 850.0 and bh >= 100.0:
                                top_row_boxes.append((bx, by, bw, bh))

                for gf in tree.findall(f'.//{{{P_NS}}}graphicFrame'):
                    xfrm = gf.find(f'.//{{{P_NS}}}xfrm')
                    if xfrm is not None:
                        off = xfrm.find(f'{{{A_NS}}}off')
                        ext = xfrm.find(f'{{{A_NS}}}ext')
                        if off is not None and ext is not None:
                            bx = emu_to_pt(off.attrib.get('x', 0))
                            by = emu_to_pt(off.attrib.get('y', 0))
                            bw = emu_to_pt(ext.attrib.get('cx', 0))
                            bh = emu_to_pt(ext.attrib.get('cy', 0))
                            if 85.0 <= by <= 110.0 and bw < 850.0 and bh >= 100.0:
                                top_row_boxes.append((bx, by, bw, bh))

                unique_boxes = []
                top_row_boxes.sort(key=lambda b: b[0])
                for b in top_row_boxes:
                    matched = False
                    for idx_u, ub in enumerate(unique_boxes):
                        if abs(ub[0] - b[0]) < 10.0:
                            if b[2] > ub[2]:
                                unique_boxes[idx_u] = b
                            matched = True
                            break
                    if not matched:
                        unique_boxes.append(b)

                unique_boxes.sort(key=lambda b: b[0])
                if len(unique_boxes) >= 2:
                    for i in range(len(unique_boxes) - 1):
                        gap = unique_boxes[i+1][0] - (unique_boxes[i][0] + unique_boxes[i][2])
                        if gap > 18.0:
                            issues.append({
                                "slide": s_num, "code": "QG-21", "severity": "HIGH",
                                "msg": f"相邻组件水平间距过大 (gap={gap:.1f} pt > 18.0 pt)！存在'A-A框-B框-B'三层多重留白叠加，需收紧间距至 6~12pt 以释放空间！"
                            })

            # QG-30: 行高齐平与分体式卡片几何规范 (Row Height Invariance)
            # 在多栏网格中 (同行并排 >= 2 个同级卡片)，卡片高度必须严格齐平 (高度极差 <= 12.0 pt)
            is_official_ending = (s_num == total_slides and ('汇报完毕' in xml_content or '批评指正' in xml_content))
            is_content_slide = (s_num > 1 and not is_official_ending and not is_agenda)
            if is_content_slide:

                # 收集页面中的真实图片 (排除校徽印章)
                pics_in_slide = []
                for p_pic in tree.findall(f'.//{{{P_NS}}}pic'):
                    p_xfrm = p_pic.find(f'.//{{{A_NS}}}xfrm')
                    if p_xfrm is not None:
                        p_off = p_xfrm.find(f'{{{A_NS}}}off')
                        p_ext = p_xfrm.find(f'{{{A_NS}}}ext')
                        if p_off is not None and p_ext is not None:
                            px = emu_to_pt(p_off.attrib.get('x', 0))
                            py = emu_to_pt(p_off.attrib.get('y', 0))
                            pw = emu_to_pt(p_ext.attrib.get('cx', 0))
                            ph = emu_to_pt(p_ext.attrib.get('cy', 0))
                            p_cNvPr = p_pic.find(f'{{{P_NS}}}nvPicPr/{{{P_NS}}}cNvPr')
                            p_name = p_cNvPr.attrib.get('name', '') if p_cNvPr is not None else ''
                            if not any(kw in p_name for kw in ('ͼƬ 4', 'ͼƬ 9', 'Picture 4', 'Picture 9', 'Logo')):
                                pics_in_slide.append((px, py, pw, ph))

                card_shapes = []
                for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                    xfrm = sp.find(f'.//{{{A_NS}}}xfrm')
                    if xfrm is not None:
                        off = xfrm.find(f'{{{A_NS}}}off')
                        ext = xfrm.find(f'{{{A_NS}}}ext')
                        if off is not None and ext is not None:
                            bx = emu_to_pt(off.attrib.get('x', 0))
                            by = emu_to_pt(off.attrib.get('y', 0))
                            bw = emu_to_pt(ext.attrib.get('cx', 0))
                            bh = emu_to_pt(ext.attrib.get('cy', 0))
                            if by > 80.0 and 80.0 <= bw < 850.0 and bh >= 50.0:
                                sp_name = sp.find(f'{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr')
                                name_val = sp_name.attrib.get('name', '') if sp_name is not None else ''
                                if any(kw in name_val for kw in ('Logo', 'ֱ', 'ͼƬ', 'Picture', 'EvidenceTable', '02', '校徽')):
                                    continue
                                # 检查是否为承托图片的背景框 (非独立文字卡片)
                                is_pic_backing = False
                                for px, py, pw, ph in pics_in_slide:
                                    if (bx <= px + 20 and by <= py + 20 and
                                        bx + bw >= px + pw - 20 and by + bh >= py + ph - 20):
                                        is_pic_backing = True
                                        break
                                if not is_pic_backing:
                                    card_shapes.append((bx, by, bw, bh, name_val))

                # 聚合同一行的卡片
                row_groups = []
                for cs in sorted(card_shapes, key=lambda c: c[1]):
                    placed = False
                    for rg in row_groups:
                        if abs(rg[0][1] - cs[1]) <= 10.0:
                            rg.append(cs)
                            placed = True
                            break
                    if not placed:
                        row_groups.append([cs])

                for rg in row_groups:
                    # 当一行包含 2 个或更多并排卡片时
                    if len(rg) >= 2:
                        heights = [c[3] for c in rg]
                        h_range = max(heights) - min(heights)
                        if h_range > 12.0:
                            issues.append({
                                "slide": s_num, "code": "QG-30", "severity": "CRITICAL",
                                "msg": f"同一行卡片高度参差不齐 (最高={max(heights):.1f}pt, 最矮={min(heights):.1f}pt, 极差={h_range:.1f}pt > 12.0pt)！严重违反行高齐平铁律！严禁对并排卡片独立动态缩减高度！"
                            })
                            break

                # QG-31: 垂直紧凑布局与杜绝中间断层死白 (No Middle Dead Zone)
                # 当页面存在上下两排多卡片结构时，若中间既无任何文本框、流程图亦无图片，且间隙 > 70pt，判定为死白断层
                multi_rows = [rg for rg in row_groups if len(rg) >= 2]
                if len(multi_rows) >= 2:
                    top_r = multi_rows[0]
                    bot_r = multi_rows[1]
                    top_r_bottom = max(c[1] + c[3] for c in top_r)
                    bot_r_top = min(c[1] for c in bot_r)
                    row_gap = bot_r_top - top_r_bottom
                    if row_gap > 70.0:
                        # 检查中间是否有实质内容承接 (如大段文本框、水平箭头、流程拓扑或表格)
                        has_middle_content = False
                        for m_sp in tree.findall(f'.//{{{P_NS}}}sp'):
                            m_xfrm = m_sp.find(f'.//{{{A_NS}}}xfrm')
                            if m_xfrm is not None:
                                m_off = m_xfrm.find(f'{{{A_NS}}}off')
                                m_ext = m_xfrm.find(f'{{{A_NS}}}ext')
                                if m_off is not None and m_ext is not None:
                                    my = emu_to_pt(m_off.attrib.get('y', 0))
                                    mh = emu_to_pt(m_ext.attrib.get('cy', 0))
                                    if top_r_bottom - 5.0 <= my <= bot_r_top - 20.0 and mh >= 25.0:
                                        has_middle_content = True
                                        break
                        if not has_middle_content:
                            issues.append({
                                "slide": s_num, "code": "QG-31", "severity": "CRITICAL",
                                "msg": f"上下排卡片之间存在巨大空白断层 (间隙={row_gap:.1f}pt > 70.0pt)！中间既无过渡内容亦未撑满母版垂直展区，产生严重空洞！"
                            })

                # QG-32: 语义完整性与反机械截字检查 (Semantic Heading & Anti-Word-Chopping)
                for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                    for p in sp.findall(f'.//{{{A_NS}}}p'):
                        txt = "".join([t.text for t in p.findall(f'.//{{{A_NS}}}t') if t.text]).strip()
                        if txt and re.match(r'^[\u4e00-\u9fa5][，。；：！？、]', txt):
                            issues.append({
                                "slide": s_num, "code": "QG-32", "severity": "CRITICAL",
                                "msg": f"检测到疑似机械截字导致的单字残缺流落正文 ('{txt[:16]}')！必须按完整标点分句，严禁暴力截断！"
                            })

                # QG-34: 学术语义色彩调和律与 <=4 色法则 (Semantic Color Harmony)
                is_official_ending = (s_num == total_slides and ('汇报完毕' in xml_content or '批评指正' in xml_content))
                is_content_slide = (s_num > 1 and not is_official_ending and not is_agenda)
                if is_content_slide:
                    slide_fills = set()
                    NEUTRAL_COLORS = {'FFFFFF', '000000', '242424', '1F2937', '2D3748', '666666', '888888', 'E2E8F0', 'CBD5E1', 'F1F5F9', 'F8FAFC'}
                    for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                        for sf in sp.findall(f'.//{{{P_NS}}}spPr/{{{A_NS}}}solidFill/{{{A_NS}}}srgbClr'):
                            val = sf.attrib.get('val', '').upper()
                            if val and val not in NEUTRAL_COLORS:
                                slide_fills.add(val)
                    if len(slide_fills) > 4:
                        issues.append({
                            "slide": s_num, "code": "QG-34", "severity": "CRITICAL",
                            "msg": f"单页显性非中性主题色数量超过上限 ({len(slide_fills)} 种 > 4 种: {sorted(list(slide_fills))})！违反学术语义色彩调和律与 <=4 色法则，色彩杂乱如彩虹！"
                        })

                    # QG-34 子项: 湖大红等深色背景文字高对比度纯白铁律
                    red_card_rects = []
                    HNU_RED_VARIANTS = {'A6232B', '8B1D23', 'C00000', '9E1B22', 'A61C1C'}
                    for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                        sf = sp.find(f'.//{{{P_NS}}}spPr/{{{A_NS}}}solidFill/{{{A_NS}}}srgbClr')
                        if sf is not None and sf.attrib.get('val', '').upper() in HNU_RED_VARIANTS:
                            xfrm = sp.find(f'.//{{{A_NS}}}xfrm')
                            if xfrm is not None:
                                off = xfrm.find(f'{{{A_NS}}}off')
                                ext = xfrm.find(f'{{{A_NS}}}ext')
                                if off is not None and ext is not None:
                                    rx = emu_to_pt(off.attrib.get('x', 0))
                                    ry = emu_to_pt(off.attrib.get('y', 0))
                                    rw = emu_to_pt(ext.attrib.get('cx', 0))
                                    rh = emu_to_pt(ext.attrib.get('cy', 0))
                                    if rw >= 50.0 and rh >= 25.0 and ry > 75.0:
                                        red_card_rects.append((rx, ry, rw, rh))

                    if red_card_rects:
                        for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                            xfrm = sp.find(f'.//{{{A_NS}}}xfrm')
                            if xfrm is None:
                                continue
                            off = xfrm.find(f'{{{A_NS}}}off')
                            ext = xfrm.find(f'{{{A_NS}}}ext')
                            if off is None or ext is None:
                                continue
                            tx = emu_to_pt(off.attrib.get('x', 0))
                            ty = emu_to_pt(off.attrib.get('y', 0))
                            for rx, ry, rw, rh in red_card_rects:
                                if rx - 5.0 <= tx <= rx + rw + 5.0 and ry - 5.0 <= ty <= ry + rh + 5.0:
                                    for r in sp.findall(f'.//{{{A_NS}}}r'):
                                        t_el = r.find(f'.//{{{A_NS}}}t')
                                        t_txt = t_el.text if t_el is not None else ''
                                        if not t_txt or not t_txt.strip():
                                            continue
                                        srgb = r.find(f'.//{{{A_NS}}}rPr/{{{A_NS}}}solidFill/{{{A_NS}}}srgbClr')
                                        sch = r.find(f'.//{{{A_NS}}}rPr/{{{A_NS}}}solidFill/{{{A_NS}}}schemeClr')
                                        is_white = False
                                        if srgb is not None and srgb.attrib.get('val', '').upper() == 'FFFFFF':
                                            is_white = True
                                        elif sch is not None and sch.attrib.get('val', '') in ('bg1', 'lt1'):
                                            is_white = True
                                        if not is_white:
                                            clr_desc = srgb.attrib.get('val') if srgb is not None else (sch.attrib.get('val') if sch is not None else 'default/dark')
                                            issues.append({
                                                "slide": s_num, "code": "QG-34", "severity": "CRITICAL",
                                                "msg": f"湖大红背景卡片上的文字 ('{t_txt[:20]}') 颜色为 {clr_desc}，未设为纯白！严重影响可读性！必须强制使用纯白色 (#FFFFFF / bg1)！"
                                            })
                                            break

                # QG-35 & QG-36: 核心分点多维充实与有机饱满度检查 (Substantive Bullets & Fill Rate)
                if is_content_slide:
                    for sp in tree.findall(f'.//{{{P_NS}}}sp'):

                        cNvPr = sp.find(f'.//{{{P_NS}}}cNvPr')
                        sp_name = cNvPr.attrib.get('name', '') if cNvPr is not None else ''
                        if sp_name == 'Body' or (sp_name and 'Body' in sp_name):
                            xfrm = sp.find(f'.//{{{A_NS}}}xfrm')
                            if xfrm is not None:
                                ext = xfrm.find(f'{{{A_NS}}}ext')
                                if ext is not None:
                                    bh_pt = emu_to_pt(ext.attrib.get('cy', 0))
                                    bw_pt = emu_to_pt(ext.attrib.get('cx', 0))
                                    paras = sp.findall(f'.//{{{A_NS}}}p')
                                    bullet_paras = [p for p in paras if "".join([t.text for t in p.findall(f'.//{{{A_NS}}}t') if t.text]).strip()]
                                    all_card_text = "".join([t.text for t in sp.findall(f'.//{{{A_NS}}}t') if t.text]).strip()
                                    # QG-35: 检查单薄分点 (高度 >= 80pt 时只有 1 条分点且字数少)
                                    if bh_pt >= 80.0 and len(bullet_paras) <= 1 and len(all_card_text) < 40:
                                        issues.append({
                                            "slide": s_num, "code": "QG-35", "severity": "CRITICAL",
                                            "msg": f"卡片 (高={bh_pt:.1f}pt) 仅有 {len(bullet_paras)} 条单薄分点 (仅 {len(all_card_text)} 字: '{all_card_text[:20]}')！严重悬空！必须扩充至 2~3 条独立客观事实分点！"
                                        })
                                    # QG-36: 饱满度下限检测
                                    if bh_pt >= 120.0 and bw_pt >= 150.0:
                                        total_text_lines = max(1, len(all_card_text) / max(1.0, bw_pt / 14.0))
                                        est_h = total_text_lines * 14.0 * 1.3 + len(bullet_paras) * 6.0
                                        fill_rate = est_h / bh_pt
                                        if fill_rate < 0.60:
                                            issues.append({
                                                "slide": s_num, "code": "QG-36", "severity": "CRITICAL",
                                                "msg": f"卡片空间利用率极低 (估算 {fill_rate*100:.1f}% < 60%, 高={bh_pt:.1f}pt, '{all_card_text[:20]}')！下半部大面积死白，必须充实分点或提升字号！"
                                            })

                    # QG-36 子项: 检查正文段落行间距及同一页一致性
                    slide_body_line_spacings = []
                    for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                        cNvPr = sp.find(f'.//{{{P_NS}}}cNvPr')
                        sp_name = cNvPr.attrib.get('name', '') if cNvPr is not None else ''
                        if sp_name == 'Body' or (sp_name and 'Body' in sp_name):
                            for p in sp.findall(f'.//{{{A_NS}}}p'):
                                p_txt = "".join([t.text for t in p.findall(f'.//{{{A_NS}}}t') if t.text]).strip()
                                if not p_txt:
                                    continue
                                pPr = p.find(f'{{{A_NS}}}pPr')
                                if pPr is not None:
                                    lnSpc = pPr.find(f'{{{A_NS}}}lnSpc')
                                    if lnSpc is not None:
                                        fixed = lnSpc.find(f'{{{A_NS}}}spcPts')
                                        if fixed is not None and 'val' in fixed.attrib:
                                            lsp = float(fixed.attrib['val']) / 100.0
                                            slide_body_line_spacings.append(lsp)
                                            if lsp < 16.9:
                                                issues.append({
                                                    "slide": s_num, "code": "QG-36", "severity": "CRITICAL",
                                                    "msg": f"正文段落行间距过紧 ({lsp:.1f} pt < 17.0 pt: '{p_txt[:20]}')！默认必须为 18~20 pt（紧凑情况最低 17.0 pt，严禁低于 17.0 pt）！"
                                                })
                    if len(slide_body_line_spacings) >= 2:
                        min_lsp = min(slide_body_line_spacings)
                        max_lsp = max(slide_body_line_spacings)
                        if max_lsp - min_lsp > 1.5:
                            issues.append({
                                "slide": s_num, "code": "QG-36", "severity": "HIGH",
                                "msg": f"同一页内文本框行间距不一致 (最小={min_lsp:.1f}pt, 最大={max_lsp:.1f}pt, 极差={max_lsp-min_lsp:.1f}pt > 1.5pt)！同一页内各文本框行间距必须保持一致！"
                            })

                    # QG-36 子项: 防虚假降级检查 (卡片空间充裕时严禁误降至 17.0 pt)
                    if slide_body_line_spacings and min(slide_body_line_spacings) < 18.0:
                        can_fit_at_18_5 = True
                        has_checked_any = False
                        for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                            cNvPr = sp.find(f'.//{{{P_NS}}}cNvPr')
                            s_name = cNvPr.attrib.get('name', '') if cNvPr is not None else ''
                            if s_name == 'Body' or (s_name and 'Body' in s_name):
                                xfrm = sp.find(f'.//{{{A_NS}}}xfrm')
                                if xfrm is None: continue
                                off = xfrm.find(f'{{{A_NS}}}off'); ext = xfrm.find(f'{{{A_NS}}}ext')
                                if off is None or ext is None: continue
                                tx = emu_to_pt(off.attrib.get('x', 0)); ty = emu_to_pt(off.attrib.get('y', 0))
                                tw = emu_to_pt(ext.attrib.get('cx', 0)); th = emu_to_pt(ext.attrib.get('cy', 0))
                                parent_card_h = th
                                for (bx, by, bw, bh, _) in bg_card_rects:
                                    if bx - 5.0 <= tx and tx + tw <= bx + bw + 5.0 and by - 5.0 <= ty <= by + bh + 5.0:
                                        parent_card_h = max(th, bh - 41.0)
                                        break
                                paras = sp.findall(f'.//{{{A_NS}}}p')
                                line_count = sum(len("".join(t.text or "" for t in p.iter(f'{{{A_NS}}}t')).split('\v')) for p in paras if "".join(t.text or "" for t in p.iter(f'{{{A_NS}}}t')).strip())
                                need_at_18_5 = line_count * 18.5 + len(paras) * 4.0
                                if need_at_18_5 > parent_card_h + 2.0:
                                    can_fit_at_18_5 = False
                                    break
                                has_checked_any = True
                        if has_checked_any and can_fit_at_18_5:
                            issues.append({
                                "slide": s_num, "code": "QG-36", "severity": "CRITICAL",
                                "msg": f"Slide {s_num} 正文行间距被虚假降级为 {min(slide_body_line_spacings):.1f} pt！底托卡片可用高度充裕，完全足以容纳 18.5 pt 标准行距！必须恢复 18~20 pt 默认行距！"
                            })

                # QG-39: renhua 人性化语言风格、学术零人称与禁忌标点检查
                if is_content_slide:
                    PRONOUN_PATTERN = re.compile(r'(\b我们\b|(?<![自无物知])我|(?<![为予])你|你们|(?<![其排])他们|大家)')
                    CLICHE_PATTERNS = [
                        (re.compile(r'不仅.*?而且'), "空洞连词 '不仅……而且……'"),
                        (re.compile(r'显著(提升|优势|增强|改善|提高|效果)'), "虚饰词 '显著...'"),
                        (re.compile(r'极大地?赋能'), "AI腔 '赋能'"),
                        (re.compile(r'旨在打造'), "官话套话 '旨在打造'"),
                        (re.compile(r'飞速发展'), "空洞套话 '飞速发展'"),
                        (re.compile(r'里程碑意义|划时代'), "浮夸修饰 '里程碑/划时代'"),
                        (re.compile(r'壁垒'), "大厂黑话 '壁垒'（改困难/阻碍）"),
                        (re.compile(r'(?:固化为)?可复用资产'), "大厂黑话 '可复用资产'（改具体工程规则/操作规范）"),
                        (re.compile(r'原子化'), "大厂黑话 '原子化'（改单项基础操作/细分步骤）"),
                        (re.compile(r'沉淀为(?:工程规则|资产)?'), "大厂黑话 '沉淀为...'（改整理成规范/写成规则）"),
                        (re.compile(r'高频科研方法'), "伪架构词 '高频科研方法'（改常用的科研方法）"),
                        (re.compile(r'标准化通信与状态感知'), "伪架构套话 '标准化通信与状态感知'（改把外部工具连接进来/了解运行状态）"),
                        (re.compile(r'自愈闭环'), "伪架构套话 '自愈闭环'（改排查并自动修正问题/完整执行流程）"),
                    ]
                    for sp in tree.findall(f'.//{{{P_NS}}}sp'):
                        for p in sp.findall(f'.//{{{A_NS}}}p'):
                            p_txt = "".join([t.text for t in p.findall(f'.//{{{A_NS}}}t') if t.text]).strip()
                            if not p_txt:
                                continue
                            # 1. 学术零人称检查
                            m_pronoun = PRONOUN_PATTERN.search(p_txt)
                            if m_pronoun:
                                issues.append({
                                    "slide": s_num, "code": "QG-39", "severity": "CRITICAL",
                                    "msg": f"正文包含主观人称代词 '{m_pronoun.group(0)}' ('{p_txt[:25]}')！学术叙述必须零人称，改用客观目的句（如'为了减少...'）！"
                                })
                            # 2. AI 套话与大厂黑话黑名单
                            for c_pat, c_desc in CLICHE_PATTERNS:
                                if c_pat.search(p_txt):
                                    issues.append({
                                        "slide": s_num, "code": "QG-39", "severity": "CRITICAL",
                                        "msg": f"检测到 {c_desc} ('{p_txt[:25]}')！违反 renhua 白描求真规范，必须去除 AI 腔与黑话！"
                                    })
                            # 3. 禁忌标点检查
                            if '——' in p_txt or '—' in p_txt:
                                issues.append({
                                    "slide": s_num, "code": "QG-39", "severity": "CRITICAL",
                                    "msg": f"正文包含破折号 ('{p_txt[:25]}')！学术排版严禁破折号，改用逗号或冒号！"
                                })
                            if '；' in p_txt:
                                issues.append({
                                    "slide": s_num, "code": "QG-39", "severity": "HIGH",
                                    "msg": f"正文包含分号 ('{p_txt[:25]}')！学术排版严禁分号，改用独立句号或分点！"
                                })


        # QG-28: 视觉节奏铁律（连续纯文字/表格页上限 <= 4 页）
        non_visual_streak = 0
        streak_start = None
        for s_idx in range(2, total_slides): # 仅检查正文内容页 (排除封面 1 和结尾页 total_slides)
            if not visual_status_by_slide.get(s_idx, True):
                if non_visual_streak == 0:
                    streak_start = s_idx
                non_visual_streak += 1
                if non_visual_streak > 4:
                    issues.append({
                        "slide": s_idx, "code": "QG-28", "severity": "CRITICAL",
                        "msg": f"违反视觉节奏铁律：连续 {non_visual_streak} 页无任何图表/流程图/公式/视觉化呈现 (Slide {streak_start}~{s_idx})，超过 <= 4 页上限！必须引入流程图或真实素材图！"
                    })
            else:
                non_visual_streak = 0
                streak_start = None

        # Task-scoped image verification. No legacy sample count / shared-folder scan.
        if total_slides > 30 and (2 not in agenda_pages or len(agenda_pages) < 2):
            issues.append({'slide': 0, 'code': 'QG-25', 'severity': 'CRITICAL',
                           'msg': '超过 30 页须首页后提纲及章节过渡页；章节语义位置仍需人工核对'})
        if agenda_details and any(d['labels'] != agenda_details[0]['labels'] for d in agenda_details):
            issues.append({'slide': 0, 'code': 'QG-25', 'severity': 'CRITICAL', 'msg': '各目录页章节清单不一致'})
        for page in skill_pages:
            if not 1 <= page <= total_slides:
                issues.append({'slide': page, 'code': 'QG-26', 'severity': 'HIGH', 'msg': 'Skill 功能页映射页码越界'})
        if asset_dir:
            asset_path = Path(asset_dir).resolve()
            if not asset_path.is_dir() or asset_path.parent.name != 'assets':
                issues.append({'slide':0,'code':'QG-22','severity':'CRITICAL','msg':'必须指定 assets/<项目>/ 专属资产目录'})
            else:
                readme = asset_path / 'README.md'
                if not readme.is_file():
                    issues.append({'slide':0,'code':'QG-22','severity':'HIGH','msg':'当前项目资产缺少 README 清册'})
                else:
                    content=readme.read_text(encoding='utf-8-sig')
                    if not all(any(w in content for w in group) for group in [('项目','标题'),('源','来源'),('模型','Model')]):
                        issues.append({'slide':0,'code':'QG-22','severity':'HIGH','msg':'当前 README 缺项目、源素材或模型声明'})
                media_hashes={hashlib.sha256(z.read(n)).hexdigest() for n in z.namelist() if n.startswith('ppt/media/') and not n.endswith('/')}
                for image in asset_path.iterdir():
                    if image.is_file() and image.suffix.lower() in ('.png','.jpg','.jpeg','.svg','.emf','.bmp','.gif'):
                        if hashlib.sha256(image.read_bytes()).hexdigest() not in media_hashes:
                            issues.append({'slide':0,'code':'QG-22','severity':'HIGH','msg':f'当前资产未在 PPT 包媒体中找到同一字节内容: {image.name}（若重编码须人工提供引用证据）'})

    min_sz = min(all_font_sizes) if all_font_sizes else 0
    max_sz = max(all_font_sizes) if all_font_sizes else 0

    print(f"\n---------------- 质量审查统计指标 ----------------")
    print(f"全篇真实图片实装数: {total_real_images} 张（含模板图片；真实性需来源核对）")
    print(f"全篇学术三线表总数: {total_tables} 张（统计不代表全门禁通过）")
    print(f"全域字体样本数: {len(all_font_sizes)}")
    print(f"全域最小字号: {min_sz:.1f} pt (标准: >= 13.0 pt)")
    print(f"全域最大字号: {max_sz:.1f} pt (含官方结束页)")
    print(f"正文文字(16~20pt)样本数: {len(body_font_sizes)}")
    print(f"注册 latin 字体集: {all_fonts}")
    print(f"注册 East Asian (ea) 字体集: {ea_fonts}")

    critical_count = sum(1 for i in issues if i['severity'] == 'CRITICAL')
    high_count = sum(1 for i in issues if i['severity'] == 'HIGH')

    if issues:
        print(f"\n[ISSUES FOUND] 共发现 {len(issues)} 处潜在违规项:")
        for iss in issues:
            print(f"  [{iss['severity']}] Slide {iss['slide']} ({iss['code']}): {iss['msg']}")

    # The audit is read-only. QG17 workspace hygiene remains an explicit task check.
    manual_review = [
        'QG01/QG07/QG09/QG23: 可编辑性、事实真实性与素材来源需逐页源证据核对',
        'QG17: 任务产物卫生需限定当前任务根检查；审计器不删除共享缓存',
        'QG20: 文本填充率是几何估算，真实换行/溢出须 PowerPoint 渲染检查',
        'QG25: 章节切换位置与当前高亮章节语义须人工核对',
        'QG26: Skill 功能页范围和三阶叙事链完整性须人工核对',
        'QG27: 内容多少、半幅图/2~4文字框及治理页流程语义须人工检查',
        'QG39: renhua 人性化语言风格、工科白描求真与自然呼吸长短句需逐页通读核对',
    ]
    if not asset_dir: manual_review.append('QG22: 未提供当前项目资产目录，资产隔离与引用完整性尚未核验')
    if not skill_pages: manual_review.append('QG26: 未提供 Skill 功能页清单，标题强绑定检查尚未应用')
    print('[MANUAL REVIEW REQUIRED] 自动审计不能替代以下检查:')
    for item in manual_review: print('  - ' + item)
    if json_report:
        Path(json_report).write_text(json.dumps({'input':str(pptx_path), 'sha256':hashlib.sha256(Path(pptx_path).read_bytes()).hexdigest(),
            'automated_pass':not issues, 'issues':issues, 'agenda_pages':agenda_pages,
            'manual_review_required':manual_review},ensure_ascii=False,indent=2),encoding='utf-8')

    if critical_count == 0 and high_count == 0:
        print(f"\n==================================================")
        print(f"[AUTOMATED CHECKS PASS] 已执行的自动检查通过；未代表人工门禁或全量质量通过。")
        print(f"==================================================")
        return True
    else:
        print(f"\n[AUDIT FAIL] 存在 {critical_count} 个严重错误与 {high_count} 个高危错误！")
        return False

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='湖南大学学术 PPT 质量门禁审查脚本')
    parser.add_argument('--input', default=r'C:\Users\w5711112\.agents\skills\academic-native-ppt-design-HNU-style\examples\投标文件--技术部分--湖大模板汇报.pptx')
    parser.add_argument('--asset-dir', help='仅检查当前任务 assets/<项目>/ 目录')
    parser.add_argument('--skill-pages', help='JSON 对象：逻辑页码映射到该功能页 Skill 名称数组')
    parser.add_argument('--json-report', help='写出机器可读结果及必须人工检查的范围')
    args = parser.parse_args()
    skill_pages = json.loads(Path(args.skill_pages).read_text(encoding='utf-8-sig')) if args.skill_pages else None
    success = audit_deck(args.input, args.asset_dir, skill_pages, args.json_report)
    sys.exit(0 if success else 1)

