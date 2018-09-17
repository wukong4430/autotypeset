# -*- coding: utf-8 -*-
# @Author: KICC
# @Date:   2018-09-16 17:34:46
# @Last Modified by:   KICC
# @Last Modified time: 2018-09-17 17:17:43

from docx.shared import Inches, Pt
from docx.oxml.ns import qn


def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None


def modify_character(back_text, p):
    for i, ch in enumerate(back_text):
        run = p.add_run(ch)
        font = run.font
        font.size = Pt(10)
        if '0' <= ch <= '9' or 'a' <= ch <= 'z':
            font.name = u'微软雅黑'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), u'微软雅黑')
        else:
            font.name = u'宋体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')


def adjust_paragraph_style(formats, paragraph):
    """
    调整段落的style
    format是传入进来的段落格式
    """
    p = paragraph  # p是需要修改的段落
    p_format = p.paragraph_format
    p_format.alignment = formats.alignment
    p_format.first_line_indent = formats.first_line_indent
    p_format.left_indent = formats.left_indent
    p_format.right_indent = formats.right_indent
    p_format.line_spacing = formats.line_spacing
    p_format.space_before = formats.space_before
