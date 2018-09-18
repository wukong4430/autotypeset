# -*- coding: utf-8 -*-
# @Author: KICC
# @Date:   2018-09-16 17:34:46
# @Last Modified by:   KICC
# @Last Modified time: 2018-09-18 16:32:48

from docx.shared import Inches, Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_LINE_SPACING

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None


def modify_character(back_text, p, category):
    """
    修改段落的中文字的字体
    param: back_text:  需要修改的段落文字
    param: p: 空白的段落(修改后段落的内容)
    param: category: 是正文还是标题还是摘要...
    """
    for i, ch in enumerate(back_text):
        run = p.add_run(ch)
        font = run.font
        if category == '正文':
            font.size = Pt(10)
            if '0' <= ch <= '9' or 'a' <= ch <= 'z' or 'A' <= ch <= 'z':
                font.name = u'微软雅黑'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), u'微软雅黑')
            else:
                font.name = u'宋体'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
        elif category == '章标题':
            # 1 绪论
            pass

        elif category == '一级标题':
            # 1.1
            pass


def adjust_paragraph_style(formats, paragraph, category):
    """
    调整段落的style
    format是传入进来的段落格式

    目前写的是根据原来的段落格式修改,但是原来的段落格式不一定是对的,
    所以应该根据传入进来的段落属于 [正文, 标题, 摘要..] 来选择段落的格式

    """
    # p = paragraph  # p是需要修改的段落
    # p_format = p.paragraph_format
    # print(formats.first_line_indent,'\n')
    # p_format.alignment = formats.alignment
    # p_format.first_line_indent = formats.first_line_indent
    # p_format.left_indent = formats.left_indent
    # p_format.right_indent = formats.right_indent
    # p_format.line_spacing = formats.line_spacing
    # p_format.space_before = formats.space_before

    # 用category来确定段落格式的话, formats参数就没用了.

    if category == '正文':
        # 标题以外的文字行距为"固定值"23磅,字符间距为"标准"
        p_format = paragraph.paragraph_format
        p_format.alignment = None
        p_format.first_line_indent = Pt(20)
        p_format.left_indent = Pt(21)
        p_format.right_indent = None
        p_format.line_spacing = Pt(23)
        p_format.space_before = None

    elif category == '节标题':
        # 1.1.1
        # 节标题段前为0.5行, 段后为0.5行
        p_format = paragraph.paragraph_format
        p_format.alignment = None
        p_format.first_line_indent = Pt(30)
        p_format.left_indent = Pt(70)
        p_format.right_indent = None
        p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p_format.line_spacing = 1.0
        p_format.space_before = None

    elif category=='章标题':
        # 1.
        # 章标题的段前0.8行,段后0.5行;
        p_format = paragraph.paragraph_format
        p_format.alignment = None
        p_format.first_line_indent = Pt(30)
        p_format.left_indent = Pt(50)
        p_format.right_indent = None
        p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p_format.line_spacing = 1.3
        p_format.space_before = None

    elif category == '摘要':
        pass
