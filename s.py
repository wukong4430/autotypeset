#!/usr/bin/env ./autotype/bin/python

# -*- coding: utf-8 -*-
# @Author: KICC
# @Date:   2018-07-27 12:03:41
# @Last Modified by:   KICC
# @Last Modified time: 2018-09-15 20:31:01

from docx import Document
from PIL import Image, ImageDraw
from io import BytesIO
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT


def draw_circles():
    document = Document()
    p = document.add_paragraph()
    r = p.add_run()
    img_size = 20
    for x in range(255):
        im = Image.new("RGB", (img_size, img_size), "white")
        draw_obj = ImageDraw.Draw(im)
        draw_obj.ellipse((0, 0, img_size - 1, img_size - 1),
                         fill=255 - x)  # 画圈
        fake_buf_file = BytesIO()  # 用BytesIO将图片保存在内存里, 减少磁盘操作
        im.save(fake_buf_file, 'png')
        r.add_picture(fake_buf_file)
        fake_buf_file.close()
        document.save('circle.docx')


def fun():
    document = Document('demo1.docx')
    p = document.paragraphs
    s = document.sections[-1]
    print(p[0].text)
    print(p[1].text)
    print(p[2].text)
    print(p[3].text)
    # get style from p[0]
    # 修改行间距
    p[1].paragraph_format.line_spacing = 4.0
    # p[0].add_run(text='I wanna have a try')
    print(p[1].paragraph_format.line_spacing)

    # 直接修改页边距
    # s[0].top_margin = 1514400
    #document.sections = s
    # s[0].orientation = WD_ORIENT.LANDSCAPE  # index 1; another index 0
    # 页边距
    print(s.bottom_margin)
    print(s.top_margin)
    print(s.left_margin)
    print(s.right_margin)

    print(s.footer_distance)  # 页脚distance
    print(s.header_distance)  # yemei distance
    print(s.orientation)

    document.save('demo1.docx')


def main():
    fun()


if __name__ == '__main__':
    main()
