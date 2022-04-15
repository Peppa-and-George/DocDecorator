"""
修改文档元素样式模块:

Paragraph:
 - 修改段落字体字号颜色加粗倾斜
 - 修改段落样式
 - 修改段落对齐方式
 - 修改段落首行缩进
 - 设置段落行距

 Run:
 - 修改run字号字体颜色

 Table:
 - 修改表格样式
 - 设置表格是否自适应
 - 修改表格对齐方式
 - 修改表格内容对齐方式
 - 修改表格字体字号颜色

 Row:
 - 修改Row内容对齐方式
 - 修改Row字体字号颜色

 Column:
 - 修改Column内容对齐方式
 - 修改Column字体字号颜色

 Cell:
 - 修改Cell内容对齐方式
 - 修改Cell字体字号颜色
"""
__all__ = ['DocStyle']

import docx
from docx.document import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, RGBColor
from docx.table import Table, _Row, _Column, _Cell
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from docx.styles.style import (
    _ParagraphStyle,
    _TableStyle,
)
from docx.text.run import Run
from element import DOCXElement


class DocStyle(object):
    """
    修改文档元素的样式
    """

    # 段落对齐方式
    P_CENTER = WD_PARAGRAPH_ALIGNMENT.CENTER
    P_LEFT = WD_PARAGRAPH_ALIGNMENT.LEFT
    P_RIGHT = WD_PARAGRAPH_ALIGNMENT.RIGHT

    # 表格对齐方式
    T_CENTER = WD_TABLE_ALIGNMENT.CENTER
    T_LEFT = WD_TABLE_ALIGNMENT.LEFT
    T_RIGHT = WD_TABLE_ALIGNMENT.RIGHT

    def __init__(self, doc: str | Document):
        if isinstance(doc, str):
            self.doc = docx.Document(doc)
        else:
            self.doc = doc

    @classmethod
    def set_paragraph_attributes(cls, paragraph: Paragraph, en_font: str = 'w:eastAsia', zh_font: str = '微软雅黑',
                                 font_size: float = None, color: tuple | str = None,
                                 bold: bool = None, italic: bool = None) -> Paragraph:
        """
        设置段落的属性：字体，字号，颜色
        :param paragraph: 需要修改的段落对象
        :param en_font: 西文字体，默认w:eastAsia
        :param zh_font: 中文字体，默认
        :param font_size: 字号,单位 磅
        :param color: 颜色，颜色为三元组，格式为(r,g,b)
            或者 一个十六进制的颜色字符串，例如：3C2F80
        :param bold: 是否加粗
        :param italic: 是否倾斜
        :return:修改后的段落对象
        """
        for run in paragraph.runs:
            cls.set_run_attribute(run, en_font, zh_font, font_size, color, bold, italic)
        return paragraph

    @staticmethod
    def set_paragraph_style(paragraph: Paragraph, style: str | _ParagraphStyle) -> Paragraph:
        """
        为段落指定样式
        :param paragraph: 段落对象
        :param style: 样式名字或者样式对象
        :return: 修改后的段落对象
        """
        paragraph.style = style
        return paragraph

    @staticmethod
    def set_paragraph_alignment(paragraph: Paragraph, alignment: int) -> Paragraph:
        """
        设置段落的对齐方式
        :param paragraph: 段落对象
        :param alignment: 对齐方式，本类里面定义的对齐方式
        :return: 修改后的段落
        """
        paragraph.alignment = alignment
        return paragraph

    @staticmethod
    def set_paragraph_indent(paragraph: Paragraph, indent_num: int) -> Paragraph:
        """
        设置段落首行缩进
        :param paragraph: 短路对象
        :param indent_num: 所经的字符数
        :return: 修改后的段落
        """
        paragraph.paragraph_format.first_line_indent = \
            paragraph.style.font.size * indent_num
        return paragraph

    @staticmethod
    def set_paragraph_line_spacing(paragraph: Paragraph, before_space: float = None,
                                   after_space: float = None) -> Paragraph:
        """
        设置段落行间距
        :param paragraph: 短路对象
        :param before_space: 行前距，单位磅
        :param after_space: 行后距，单位磅
        :return: 修改后的段落对象
        """
        if before_space is not None:
            paragraph.paragraph_format.space_before = Pt(before_space)
        if after_space is not None:
            paragraph.paragraph_format.space_after = Pt(after_space)
        return paragraph

    @staticmethod
    def set_run_attribute(run: Run, en_font: str = 'w:eastAsia', zh_font: str = '微软雅黑',
                          font_size: float = None, color: tuple | str = None,
                          bold: bool = None, italic: bool = None) -> Run:
        """
        设置run对象的属性
        :param run: 需要修饰的run对象
        :param en_font: 西文字体
        :param zh_font: 中文字体
        :param font_size: 字体大小，单位磅
        :param color: 颜色，颜色为三元组，格式为(r,g,b)
            或者 一个十六进制的颜色字符串，例如：3C2F80
        :param bold: 是否加粗
        :param italic: 是否倾斜
        :return: 修改后的run对象
        """
        if en_font is not None or zh_font is not None:
            run.font.name = zh_font
            run._element.rPr.rFonts.set(qn(en_font), zh_font)
        if font_size is not None:
            run.font.size = Pt(font_size)
        if color is not None:
            if isinstance(color, tuple):
                r, g, b = color
                run.font.color = RGBColor(r, g, b)
            else:
                run.font.color = RGBColor.from_string(color)
        if bold is not None:
            run.bold = bold
        if italic is not None:
            run.italic = italic
        return run

    @staticmethod
    def set_table_style(table: Table, style: _TableStyle | str) -> Table:
        """
        设置表格的样式
        :param table: 表格对象
        :param style: 表格样式对象或者表格样式名称
        :return: 表格对象
        """
        table.style = style
        return table

    @staticmethod
    def set_table_autofix(table: Table, autofix: bool = None):
        """
        设置表格是否自适应
        :param table: 表格对象
        :param autofix: 是否自适应
        :return: 修改后的表格
        """
        table.autofit = autofix
        return table

    @staticmethod
    def set_table_alignment(table: Table, alignment: int) -> Table:
        """
        设置表格对齐方式,表格相对于文档的对齐方式
        :param table: 表格对象
        :param alignment: 表格对齐方式
        :return: 修改后的表格对象
        """
        table.alignment = alignment
        return table

    @classmethod
    def set_table_element_content_alignment(cls, element: Table | _Row | _Column | _Cell, alignment: int) -> Table:
        """
        设置表格内容的对齐方式
        :param element: 表格元素对象,包含 表格，行，列，单元格
        :param alignment: 对齐方式
        :return: 修改后的表格
        """
        paragraphs = DOCXElement.get_paragraphs_of_table(element)
        for paragraph in paragraphs:
            cls.set_paragraph_alignment(paragraph, alignment)
        return element

    @classmethod
    def set_table_element_content_attribute(cls, element: Table | _Row | _Column | _Cell, en_font: str = 'w:eastAsia',
                                            zh_font: str = '微软雅黑',
                                            font_size: float = None, color: tuple | str = None,
                                            bold: bool = None, italic: bool = None) -> Table:
        """
        修改表格内容的字体字号颜色是否加粗是否倾斜
        :param element: 需要修饰的表格元素对象,包含 表格，行，列，单元格
        :param en_font: 西文字体
        :param zh_font: 中文字体
        :param font_size: 字体大小，单位磅
        :param color: 颜色，颜色为三元组，格式为(r,g,b)
            或者 一个十六进制的颜色字符串，例如：3C2F80
        :param bold: 是否加粗
        :param italic: 是否倾斜
        :return: 修改后的表格对象
        """
        runs = DOCXElement.get_runs_in_table_element(element)
        for run in runs:
            cls.set_run_attribute(run, en_font, zh_font, font_size, color, bold, italic)
        return element


class StyleName(object):
    class ParagraphStyle(object):
        Normal = "Normal"
        Header = "Header"
        Footer = "Footer"
        Heading_1 = "Heading 1"
        Heading_2 = "Heading 2"
        Heading_3 = "Heading 3"
        Heading_4 = "Heading 4"
        Heading_5 = "Heading 5"
        Heading_6 = "Heading 6"
        Heading_7 = "Heading 7"
        Heading_8 = "Heading 8"
        Heading_9 = "Heading 9"
        Normal_Table = "Normal Table"
        No_Spacing = "No Spacing"
        Title = "Title"
        Subtitle = "Subtitle"
        List_Paragraph = "List Paragraph"
        Body_Text = "Body Text"
        Body_Text_2 = "Body Text 2"
        Body_Text_3 = "Body Text 3"
        List = "List"
        List_2 = "List 2"
        List_3 = "List 3"
        List_Bullet = "List Bullet"
        List_Bullet_2 = "List Bullet 2"
        List_Bullet_3 = "List Bullet 3"
        List_Number = "List Number"
        List_Number_2 = "List Number 2"
        List_Number_3 = "List Number 3"
        List_Continue = "List Continue"
        List_Continue_2 = "List Continue 2"
        List_Continue_3 = "List Continue 3"
        macro = "macro"
        Quote = "Quote"
        Caption = "Caption"
        Intense_Quote = "Intense Quote"
        TOC_Heading = "TOC Heading"
        Table_Grid = "Table Grid"
        Light_Shading = "Light Shading"
        Light_Shading_Accent_1 = "Light Shading Accent 1"
        Light_Shading_Accent_2 = "Light Shading Accent 2"
        Light_Shading_Accent_3 = "Light Shading Accent 3"
        Light_Shading_Accent_4 = "Light Shading Accent 4"
        Light_Shading_Accent_5 = "Light Shading Accent 5"
        Light_Shading_Accent_6 = "Light Shading Accent 6"
        Light_List = "Light List"
        Light_List_Accent_1 = "Light List Accent 1"
        Light_List_Accent_2 = "Light List Accent 2"
        Light_List_Accent_3 = "Light List Accent 3"
        Light_List_Accent_4 = "Light List Accent 4"
        Light_List_Accent_5 = "Light List Accent 5"
        Light_List_Accent_6 = "Light List Accent 6"
        Light_Grid = "Light Grid"
        Light_Grid_Accent_1 = "Light Grid Accent 1"
        Light_Grid_Accent_2 = "Light Grid Accent 2"
        Light_Grid_Accent_3 = "Light Grid Accent 3"
        Light_Grid_Accent_4 = "Light Grid Accent 4"
        Light_Grid_Accent_5 = "Light Grid Accent 5"
        Light_Grid_Accent_6 = "Light Grid Accent 6"
        Medium_Shading_1 = "Medium Shading 1"
        Medium_Shading_1_Accent_1 = "Medium Shading 1 Accent 1"
        Medium_Shading_1_Accent_2 = "Medium Shading 1 Accent 2"
        Medium_Shading_1_Accent_3 = "Medium Shading 1 Accent 3"
        Medium_Shading_1_Accent_4 = "Medium Shading 1 Accent 4"
        Medium_Shading_1_Accent_5 = "Medium Shading 1 Accent 5"
        Medium_Shading_1_Accent_6 = "Medium Shading 1 Accent 6"
        Medium_Shading_2 = "Medium Shading 2"
        Medium_Shading_2_Accent_1 = "Medium Shading 2 Accent 1"
        Medium_Shading_2_Accent_2 = "Medium Shading 2 Accent 2"
        Medium_Shading_2_Accent_3 = "Medium Shading 2 Accent 3"
        Medium_Shading_2_Accent_4 = "Medium Shading 2 Accent 4"
        Medium_Shading_2_Accent_5 = "Medium Shading 2 Accent 5"
        Medium_Shading_2_Accent_6 = "Medium Shading 2 Accent 6"
        Medium_List_1 = "Medium List 1"
        Medium_List_1_Accent_1 = "Medium List 1 Accent 1"
        Medium_List_1_Accent_2 = "Medium List 1 Accent 2"
        Medium_List_1_Accent_3 = "Medium List 1 Accent 3"
        Medium_List_1_Accent_4 = "Medium List 1 Accent 4"
        Medium_List_1_Accent_5 = "Medium List 1 Accent 5"
        Medium_List_1_Accent_6 = "Medium List 1 Accent 6"
        Medium_List_2 = "Medium List 2"
        Medium_List_2_Accent_1 = "Medium List 2 Accent 1"
        Medium_List_2_Accent_2 = "Medium List 2 Accent 2"
        Medium_List_2_Accent_3 = "Medium List 2 Accent 3"
        Medium_List_2_Accent_4 = "Medium List 2 Accent 4"
        Medium_List_2_Accent_5 = "Medium List 2 Accent 5"
        Medium_List_2_Accent_6 = "Medium List 2 Accent 6"
        Medium_Grid_1 = "Medium Grid 1"
        Medium_Grid_1_Accent_1 = "Medium Grid 1 Accent 1"
        Medium_Grid_1_Accent_2 = "Medium Grid 1 Accent 2"
        Medium_Grid_1_Accent_3 = "Medium Grid 1 Accent 3"
        Medium_Grid_1_Accent_4 = "Medium Grid 1 Accent 4"
        Medium_Grid_1_Accent_5 = "Medium Grid 1 Accent 5"
        Medium_Grid_1_Accent_6 = "Medium Grid 1 Accent 6"
        Medium_Grid_2 = "Medium Grid 2"
        Medium_Grid_2_Accent_1 = "Medium Grid 2 Accent 1"
        Medium_Grid_2_Accent_2 = "Medium Grid 2 Accent 2"
        Medium_Grid_2_Accent_3 = "Medium Grid 2 Accent 3"
        Medium_Grid_2_Accent_4 = "Medium Grid 2 Accent 4"
        Medium_Grid_2_Accent_5 = "Medium Grid 2 Accent 5"
        Medium_Grid_2_Accent_6 = "Medium Grid 2 Accent 6"
        Medium_Grid_3 = "Medium Grid 3"
        Medium_Grid_3_Accent_1 = "Medium Grid 3 Accent 1"
        Medium_Grid_3_Accent_2 = "Medium Grid 3 Accent 2"
        Medium_Grid_3_Accent_3 = "Medium Grid 3 Accent 3"
        Medium_Grid_3_Accent_4 = "Medium Grid 3 Accent 4"
        Medium_Grid_3_Accent_5 = "Medium Grid 3 Accent 5"
        Medium_Grid_3_Accent_6 = "Medium Grid 3 Accent 6"
        Dark_List = "Dark List"
        Dark_List_Accent_1 = "Dark List Accent 1"
        Dark_List_Accent_2 = "Dark List Accent 2"
        Dark_List_Accent_3 = "Dark List Accent 3"
        Dark_List_Accent_4 = "Dark List Accent 4"
        Dark_List_Accent_5 = "Dark List Accent 5"
        Dark_List_Accent_6 = "Dark List Accent 6"
        Colorful_Shading = "Colorful Shading"
        Colorful_Shading_Accent_1 = "Colorful Shading Accent 1"
        Colorful_Shading_Accent_2 = "Colorful Shading Accent 2"
        Colorful_Shading_Accent_3 = "Colorful Shading Accent 3"
        Colorful_Shading_Accent_4 = "Colorful Shading Accent 4"
        Colorful_Shading_Accent_5 = "Colorful Shading Accent 5"
        Colorful_Shading_Accent_6 = "Colorful Shading Accent 6"
        Colorful_List = "Colorful List"
        Colorful_List_Accent_1 = "Colorful List Accent 1"
        Colorful_List_Accent_2 = "Colorful List Accent 2"
        Colorful_List_Accent_3 = "Colorful List Accent 3"
        Colorful_List_Accent_4 = "Colorful List Accent 4"
        Colorful_List_Accent_5 = "Colorful List Accent 5"
        Colorful_List_Accent_6 = "Colorful List Accent 6"
        Colorful_Grid = "Colorful Grid"
        Colorful_Grid_Accent_1 = "Colorful Grid Accent 1"
        Colorful_Grid_Accent_2 = "Colorful Grid Accent 2"
        Colorful_Grid_Accent_3 = "Colorful Grid Accent 3"
        Colorful_Grid_Accent_4 = "Colorful Grid Accent 4"
        Colorful_Grid_Accent_5 = "Colorful Grid Accent 5"
        Colorful_Grid_Accent_6 = "Colorful Grid Accent 6"

    class TableStyle(object):
        Normal_Table = "Normal Table"
        Table_Grid = "Table Grid"
        Light_Shading = "Light Shading"
        Light_Shading_Accent_1 = "Light Shading Accent 1"
        Light_Shading_Accent_2 = "Light Shading Accent 2"
        Light_Shading_Accent_3 = "Light Shading Accent 3"
        Light_Shading_Accent_4 = "Light Shading Accent 4"
        Light_Shading_Accent_5 = "Light Shading Accent 5"
        Light_Shading_Accent_6 = "Light Shading Accent 6"
        Light_List = "Light List"
        Light_List_Accent_1 = "Light List Accent 1"
        Light_List_Accent_2 = "Light List Accent 2"
        Light_List_Accent_3 = "Light List Accent 3"
        Light_List_Accent_4 = "Light List Accent 4"
        Light_List_Accent_5 = "Light List Accent 5"
        Light_List_Accent_6 = "Light List Accent 6"
        Light_Grid = "Light Grid"
        Light_Grid_Accent_1 = "Light Grid Accent 1"
        Light_Grid_Accent_2 = "Light Grid Accent 2"
        Light_Grid_Accent_3 = "Light Grid Accent 3"
        Light_Grid_Accent_4 = "Light Grid Accent 4"
        Light_Grid_Accent_5 = "Light Grid Accent 5"
        Light_Grid_Accent_6 = "Light Grid Accent 6"
        Medium_Shading_1 = "Medium Shading 1"
        Medium_Shading_1_Accent_1 = "Medium Shading 1 Accent 1"
        Medium_Shading_1_Accent_2 = "Medium Shading 1 Accent 2"
        Medium_Shading_1_Accent_3 = "Medium Shading 1 Accent 3"
        Medium_Shading_1_Accent_4 = "Medium Shading 1 Accent 4"
        Medium_Shading_1_Accent_5 = "Medium Shading 1 Accent 5"
        Medium_Shading_1_Accent_6 = "Medium Shading 1 Accent 6"
        Medium_Shading_2 = "Medium Shading 2"
        Medium_Shading_2_Accent_1 = "Medium Shading 2 Accent 1"
        Medium_Shading_2_Accent_2 = "Medium Shading 2 Accent 2"
        Medium_Shading_2_Accent_3 = "Medium Shading 2 Accent 3"
        Medium_Shading_2_Accent_4 = "Medium Shading 2 Accent 4"
        Medium_Shading_2_Accent_5 = "Medium Shading 2 Accent 5"
        Medium_Shading_2_Accent_6 = "Medium Shading 2 Accent 6"
        Medium_List_1 = "Medium List 1"
        Medium_List_1_Accent_1 = "Medium List 1 Accent 1"
        Medium_List_1_Accent_2 = "Medium List 1 Accent 2"
        Medium_List_1_Accent_3 = "Medium List 1 Accent 3"
        Medium_List_1_Accent_4 = "Medium List 1 Accent 4"
        Medium_List_1_Accent_5 = "Medium List 1 Accent 5"
        Medium_List_1_Accent_6 = "Medium List 1 Accent 6"
        Medium_List_2 = "Medium List 2"
        Medium_List_2_Accent_1 = "Medium List 2 Accent 1"
        Medium_List_2_Accent_2 = "Medium List 2 Accent 2"
        Medium_List_2_Accent_3 = "Medium List 2 Accent 3"
        Medium_List_2_Accent_4 = "Medium List 2 Accent 4"
        Medium_List_2_Accent_5 = "Medium List 2 Accent 5"
        Medium_List_2_Accent_6 = "Medium List 2 Accent 6"
        Medium_Grid_1 = "Medium Grid 1"
        Medium_Grid_1_Accent_1 = "Medium Grid 1 Accent 1"
        Medium_Grid_1_Accent_2 = "Medium Grid 1 Accent 2"
        Medium_Grid_1_Accent_3 = "Medium Grid 1 Accent 3"
        Medium_Grid_1_Accent_4 = "Medium Grid 1 Accent 4"
        Medium_Grid_1_Accent_5 = "Medium Grid 1 Accent 5"
        Medium_Grid_1_Accent_6 = "Medium Grid 1 Accent 6"
        Medium_Grid_2 = "Medium Grid 2"
        Medium_Grid_2_Accent_1 = "Medium Grid 2 Accent 1"
        Medium_Grid_2_Accent_2 = "Medium Grid 2 Accent 2"
        Medium_Grid_2_Accent_3 = "Medium Grid 2 Accent 3"
        Medium_Grid_2_Accent_4 = "Medium Grid 2 Accent 4"
        Medium_Grid_2_Accent_5 = "Medium Grid 2 Accent 5"
        Medium_Grid_2_Accent_6 = "Medium Grid 2 Accent 6"
        Medium_Grid_3 = "Medium Grid 3"
        Medium_Grid_3_Accent_1 = "Medium Grid 3 Accent 1"
        Medium_Grid_3_Accent_2 = "Medium Grid 3 Accent 2"
        Medium_Grid_3_Accent_3 = "Medium Grid 3 Accent 3"
        Medium_Grid_3_Accent_4 = "Medium Grid 3 Accent 4"
        Medium_Grid_3_Accent_5 = "Medium Grid 3 Accent 5"
        Medium_Grid_3_Accent_6 = "Medium Grid 3 Accent 6"
        Dark_List = "Dark List"
        Dark_List_Accent_1 = "Dark List Accent 1"
        Dark_List_Accent_2 = "Dark List Accent 2"
        Dark_List_Accent_3 = "Dark List Accent 3"
        Dark_List_Accent_4 = "Dark List Accent 4"
        Dark_List_Accent_5 = "Dark List Accent 5"
        Dark_List_Accent_6 = "Dark List Accent 6"
        Colorful_Shading = "Colorful Shading"
        Colorful_Shading_Accent_1 = "Colorful Shading Accent 1"
        Colorful_Shading_Accent_2 = "Colorful Shading Accent 2"
        Colorful_Shading_Accent_3 = "Colorful Shading Accent 3"
        Colorful_Shading_Accent_4 = "Colorful Shading Accent 4"
        Colorful_Shading_Accent_5 = "Colorful Shading Accent 5"
        Colorful_Shading_Accent_6 = "Colorful Shading Accent 6"
        Colorful_List = "Colorful List"
        Colorful_List_Accent_1 = "Colorful List Accent 1"
        Colorful_List_Accent_2 = "Colorful List Accent 2"
        Colorful_List_Accent_3 = "Colorful List Accent 3"
        Colorful_List_Accent_4 = "Colorful List Accent 4"
        Colorful_List_Accent_5 = "Colorful List Accent 5"
        Colorful_List_Accent_6 = "Colorful List Accent 6"
        Colorful_Grid = "Colorful Grid"
        Colorful_Grid_Accent_1 = "Colorful Grid Accent 1"
        Colorful_Grid_Accent_2 = "Colorful Grid Accent 2"
        Colorful_Grid_Accent_3 = "Colorful Grid Accent 3"
        Colorful_Grid_Accent_4 = "Colorful Grid Accent 4"
        Colorful_Grid_Accent_5 = "Colorful Grid Accent 5"
        Colorful_Grid_Accent_6 = "Colorful Grid Accent 6"
