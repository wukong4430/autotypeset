# -*- coding: utf-8 -*-
# @Author: KICC
# @Date:   2018-09-16 17:34:46
# @Last Modified by:   KICC
# @Last Modified time: 2018-09-16 17:35:37


def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None
