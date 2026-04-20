#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""桌牌PDF生成的核心逻辑，与GUI解耦"""

import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
import os
import platform
import glob


def find_chinese_font():
    """跨平台查找可用中文字体，返回 (font_name, font_path)"""
    system = platform.system()

    if system == 'Darwin':  # macOS
        candidates = [
            ('SourceHanSans', os.path.expanduser('~/Library/Fonts/SourceHanSansCN-Regular.ttf')),
            ('Heiti', '/System/Library/Fonts/STHeiti Medium.ttc'),
            ('HeitiLight', '/System/Library/Fonts/STHeiti Light.ttc'),
            ('PingFang', '/System/Library/Fonts/PingFang.ttc'),
            ('ArialUnicode', '/Library/Fonts/Arial Unicode.ttf'),
        ]
        # 也搜索 User Fonts 目录下的思源/Noto 字体
        for p in glob.glob(os.path.expanduser('~/Library/Fonts/*Hei*')):
            candidates.append(('UserHeiFont', p))
        for p in glob.glob(os.path.expanduser('~/Library/Fonts/*Noto*CJK*')):
            candidates.append(('NotoCJK', p))
        for p in glob.glob(os.path.expanduser('~/Library/Fonts/*Source*Han*')):
            candidates.append(('SourceHan', p))
    elif system == 'Windows':
        windir = os.environ.get('WINDIR', r'C:\Windows')
        candidates = [
            ('SimHei', os.path.join(windir, 'Fonts', 'simhei.ttf')),
            ('MicrosoftYaHei', os.path.join(windir, 'Fonts', 'msyh.ttc')),
            ('MicrosoftYaHeiBold', os.path.join(windir, 'Fonts', 'msyhbd.ttc')),
            ('SimSun', os.path.join(windir, 'Fonts', 'simsun.ttc')),
            ('FangSong', os.path.join(windir, 'Fonts', 'SIMFANG.TTF')),
        ]
    else:  # Linux
        candidates = [
            ('NotoCJK', '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'),
            ('NotoCJK', '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc'),
            ('WenQuanYi', '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc'),
            ('DroidFallback', '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf'),
        ]
        # 搜索额外路径
        for p in glob.glob('/usr/share/fonts/**/*CJK*', recursive=True):
            candidates.append(('LinuxCJK', p))

    for font_name, font_path in candidates:
        try:
            if os.path.exists(font_path):
                pdfmetrics.registerFont(TTFont(font_name, font_path))
                return font_name, font_path
        except Exception:
            continue

    return 'Helvetica', None


def process_name(name):
    """处理名字：两字名中间插入白色占位字符，使宽度与三字名一致"""
    clean_name = str(name).strip().replace(' ', '').replace('\u3000', '')
    if len(clean_name) == 2:
        return clean_name[0] + '\u5350' + clean_name[1]  # '卍' as invisible spacer
    return clean_name


def read_names_from_excel(excel_path):
    """读取Excel第一列的姓名，返回名字列表，失败返回None"""
    try:
        df = pd.read_excel(excel_path)
        col = df.columns[0]
        names = df[col].dropna().astype(str).tolist()
        # 跳过标题行
        if names and names[0].strip() == '姓名':
            names = names[1:]
        return [n.strip() for n in names if n.strip()]
    except Exception as e:
        return None


def generate_pdf(names, output_path, progress_callback=None):
    """生成桌牌PDF文件。

    Args:
        names: 姓名列表（原始名字，会自动处理两字名）
        output_path: PDF输出完整路径
        progress_callback: 回调函数(current, total, message)
    """
    page_width, page_height = A4
    card_width = 200 * mm
    card_height = 99 * mm
    crop_margin = 10 * mm
    cards_per_page = 3
    font_size = 110

    font_name, _ = find_chinese_font()

    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    c = canvas.Canvas(output_path, pagesize=A4)
    processed_names = [process_name(n) for n in names]
    total_pages = ((len(processed_names) - 1) // cards_per_page + 1) * 2

    page_idx = 0
    for i in range(0, len(processed_names), cards_per_page):
        page_names = processed_names[i:i + cards_per_page]

        # 正面
        if progress_callback:
            progress_callback(page_idx + 1, total_pages, '正在生成正面...')
        _draw_crop_lines(c, page_width, page_height, card_width)
        for j, name in enumerate(page_names):
            y = page_height - card_height - j * card_height
            _draw_name(c, name, 0, y, card_width, card_height, font_name, font_size, crop_margin=0)
        c.showPage()
        page_idx += 1

        # 反面
        if progress_callback:
            progress_callback(page_idx + 1, total_pages, '正在生成反面...')
        for j, name in enumerate(page_names):
            y = page_height - card_height - j * card_height
            _draw_name(c, name, crop_margin, y, card_width, card_height, font_name, font_size, crop_margin=crop_margin)
        c.showPage()
        page_idx += 1

    c.save()
    if progress_callback:
        progress_callback(total_pages, total_pages, '生成完成！')


def _draw_crop_lines(c, page_width, page_height, card_width):
    """绘制裁切线"""
    c.setStrokeColorRGB(0.5, 0.5, 0.5)
    c.setLineWidth(0.3)
    c.line(0, page_height - 99 * mm, card_width, page_height - 99 * mm)
    c.line(0, page_height - 198 * mm, card_width, page_height - 198 * mm)
    c.line(card_width, 0, card_width, 300 * mm)
    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(0.5)


def _draw_name(c, name, x, y, card_width, card_height, font_name, font_size, crop_margin=0):
    """绘制单个桌牌名字"""
    text_x = x + card_width / 2
    text_y = y + card_height / 2 - 10 * mm
    c.setFont(font_name, font_size)

    if len(name) == 3 and name[1] == '\u5350':
        # 两字名 + 白色占位符
        spacer_w = c.stringWidth('\u5350', font_name, font_size)
        c.setFillColorRGB(0, 0, 0)
        first_w = c.stringWidth(name[0], font_name, font_size)
        c.drawString(text_x - first_w - spacer_w / 2, text_y, name[0])
        c.setFillColorRGB(1, 1, 1)
        c.drawString(text_x - spacer_w / 2, text_y, '\u5350')
        c.setFillColorRGB(0, 0, 0)
        c.drawString(text_x + spacer_w / 2, text_y, name[2])
    else:
        c.setFillColorRGB(0, 0, 0)
        tw = c.stringWidth(name, font_name, font_size)
        c.drawString(text_x - tw / 2, text_y, name)

    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(0.5)


def create_template_excel(output_path):
    """创建模板Excel文件（第一行写「姓名」）"""
    df = pd.DataFrame({'姓名': ['张三', '李四', '王五']})
    df.to_excel(output_path, index=False)
