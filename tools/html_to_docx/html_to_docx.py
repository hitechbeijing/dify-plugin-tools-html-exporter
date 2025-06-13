import io
import logging
import re
from typing import Generator, Dict, Any, Optional, List, Tuple

from bs4 import BeautifulSoup, Tag, NavigableString
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.table import Table, _Cell

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from tools.html_to_docx.font_enum import DocxFontEnum
from tools.utils.file_utils import get_meta_data
from tools.utils.mimetype_utils import MimeType
from tools.utils.param_utils import get_html_text


class HtmlToDocxTool(Tool):
    logger = logging.getLogger(__name__)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._processed_tags = set()
        self._base_font_size = Pt(12)  # 基础字号
        self._current_paragraph = None  # 跟踪当前段落
        self._current_style_state = None  # 当前样式状态
        self._list_level = 0  # 列表嵌套层级
        self._div_stack = []  # 用于跟踪div嵌套

    def _invoke(self, tool_parameters: dict) -> Generator[ToolInvokeMessage, None, None]:
        html_text = get_html_text(tool_parameters)
        result_file_bytes = html_text.encode("utf-8")
        try:
            doc = self.create_document_with_styles()
            self.html_to_docx(doc, result_file_bytes)

            result_bytes_io = io.BytesIO()
            doc.save(result_bytes_io)
            result_file_bytes = result_bytes_io.getvalue()
        except Exception as e:
            self.logger.exception("Failed to convert HTML to DOCX")
            yield self.create_text_message(f"Failed to convert HTML text to DOCX file, error: {str(e)}")
            return

        yield self.create_blob_message(
            blob=result_file_bytes,
            meta=get_meta_data(
                mime_type=MimeType.DOCX,
                output_filename=tool_parameters.get("output_filename"),
            ),
        )

    def create_document_with_styles(self):
        doc = Document()

        # 设置默认段落样式
        style = doc.styles['Normal']
        font = style.font
        font.name = DocxFontEnum.TIMES_NEW_ROMAN
        font.size = self._base_font_size
        # 移除默认颜色设置，避免覆盖
        rPr = style.element.get_or_add_rPr()
        rPr.rFonts.set(qn('w:eastAsia'), DocxFontEnum.SONG_TI)

        # 设置段落格式
        para_format = style.paragraph_format
        para_format.space_before = Pt(6)
        para_format.space_after = Pt(6)
        para_format.line_spacing = 1.5

        return doc

    def html_to_docx(self, doc: Document, html: str):
        soup = BeautifulSoup(html, 'html.parser')
        body = soup.find('body') or soup
        
        # 重置处理状态
        self._processed_tags = set()
        self._current_paragraph = None
        # 移除默认颜色设置
        self._current_style_state = {
            'bold': False,
            'italic': False,
            'underline': False,
            'font_size': None,
            'color': None,  # 不再设置默认颜色
            'font_family': None,
            'line_height': None
        }
        self._list_level = 0
        self._div_stack = []
        
        # 处理所有子元素
        for child in body.children:
            if isinstance(child, Tag):
                self.process_tag(doc, child, self._current_style_state.copy())

    def process_tag(self, doc: Document, tag: Tag, parent_style: Dict[str, Any] = None):
        if self._is_processed(tag):
            return

        name = tag.name
        self._mark_as_processed(tag)

        # 继承父样式
        current_style = parent_style.copy() if parent_style else self._current_style_state.copy()
        
        # 处理div容器
        if name == 'div':
            # 保存当前状态
            self._div_stack.append({
                'current_paragraph': self._current_paragraph,
                'current_style': self._current_style_state.copy()
            })
            
            # 更新当前样式状态
            if tag.has_attr('style'):
                self._update_style_from_attributes(current_style, tag)
            
            # 设置当前样式状态为div的样式
            self._current_style_state = current_style.copy()
            
            # 处理div的所有子节点
            for child in tag.children:
                if isinstance(child, Tag):
                    self.process_tag(doc, child, current_style.copy())
                elif isinstance(child, NavigableString) and child.strip():
                    # 如果div内有直接文本，创建段落
                    if not self._current_paragraph:
                        self._current_paragraph = doc.add_paragraph()
                    run = self._current_paragraph.add_run(child.strip())
                    self._apply_run_style(run, current_style)
            
            # 恢复之前的样式状态
            if self._div_stack:
                prev_state = self._div_stack.pop()
                self._current_paragraph = prev_state['current_paragraph']
                self._current_style_state = prev_state['current_style']
            return

        # 处理标题标签
        if name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            # 创建普通段落而不是标题
            paragraph = doc.add_paragraph()
            self._current_paragraph = paragraph
            
            # 创建标题样式状态
            level = int(name[1])
            heading_style = current_style.copy()
            heading_style['bold'] = True
            # 设置标题字号，级别越高字号越大
            base_sizes = {1: 24, 2: 20, 3: 18, 4: 16, 5: 14, 6: 12}
            heading_style['font_size'] = Pt(base_sizes.get(level, 12))
            
            # 应用标题特定样式
            if tag.has_attr('style'):
                self._update_style_from_attributes(heading_style, tag)
            
            # 处理标题内容
            self._process_children(doc, paragraph, tag, heading_style)
            self._current_paragraph = None
            
            # 设置标题对齐方式
            self.set_paragraph_alignment(paragraph, tag)
            return

        # 处理段落
        if name == 'p':
            # 创建新段落
            paragraph = doc.add_paragraph()
            self._current_paragraph = paragraph

            # 设置段落对齐方式
            self.set_paragraph_alignment(paragraph, tag)

            # 应用段落样式
            if tag.has_attr('style'):
                # 先更新当前样式状态
                self._update_style_from_attributes(current_style, tag)
                # 然后应用段落样式
                self.apply_paragraph_styles(paragraph, tag['style'])
            
            # 处理内容
            self._process_children(doc, paragraph, tag, current_style)
            self._current_paragraph = None
            return

        # 处理表格
        if name == 'table':
            self.handle_table(doc, tag, current_style)
            return

        # 处理列表
        if name in ['ul', 'ol']:
            self._list_level += 1
            self.handle_list(doc, tag, current_style)
            self._list_level -= 1
            return

        # 处理列表项
        if name == 'li':
            # 列表项由handle_list处理，这里跳过
            return

        # 处理换行
        if name == 'br':
            if self._current_paragraph:
                self._current_paragraph.add_run('\n')
            else:
                doc.add_paragraph('\n')
            return

        # 处理其他块级元素
        if name not in ['span', 'b', 'strong', 'i', 'em', 'u', 'small', 'font']:
            text = tag.get_text().strip()
            if text:
                paragraph = doc.add_paragraph(text)
                self.set_paragraph_alignment(paragraph, tag)
                self._current_paragraph = None

    def _process_children(self, doc: Document, paragraph: Paragraph, parent_tag: Tag, parent_style: Dict[str, Any]):
        """递归处理子节点，维护样式状态"""
        for child in parent_tag.children:
            if isinstance(child, Tag):
                # 复制父样式作为当前基础样式
                current_style = parent_style.copy()
                
                # 根据标签类型更新样式
                tag_name = child.name
                if tag_name in ['b', 'strong']:
                    current_style['bold'] = True
                elif tag_name in ['i', 'em']:
                    current_style['italic'] = True
                elif tag_name == 'u':
                    current_style['underline'] = True
                elif tag_name == 'small':
                    # 在基础字号上减少2pt
                    if current_style['font_size'] is None:
                        current_style['font_size'] = self._base_font_size - Pt(2)
                    else:
                        current_style['font_size'] = current_style['font_size'] - Pt(2)
                elif tag_name in ['span', 'font']:
                    # 处理span和font的内联样式 - 无论是否有style属性都处理
                    self._update_style_from_attributes(current_style, child)
                
                # 处理块级标签（如嵌套的p/div）
                if tag_name in ['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'table', 'ul', 'ol']:
                    # 创建新段落处理块级元素
                    self.process_tag(doc, child, current_style)
                else:
                    # 递归处理内联元素
                    self._process_children(doc, paragraph, child, current_style)
                    
            elif isinstance(child, NavigableString):
                # 文本节点，应用当前样式
                text = str(child).replace('\xa0', ' ')  # 替换不间断空格
                if text.strip() or text == '\n':  # 保留换行符
                    run = paragraph.add_run(text)
                    self._apply_run_style(run, parent_style)

    def _update_style_from_attributes(self, style_state: Dict[str, Any], tag: Tag):
        """从标签属性更新样式状态 - 优化颜色处理"""
        # 处理style属性
        if tag.has_attr('style'):
            styles = self._parse_style_string(tag['style'])
            
            # 字体颜色 - 优先处理
            if "color" in styles:
                color = styles["color"].strip()
                parsed_color = self._parse_color(color)
                if parsed_color:
                    style_state['color'] = parsed_color
            
            # 字体粗细
            if "font-weight" in styles:
                weight = styles["font-weight"].strip()
                style_state['bold'] = weight in ["bold", "bolder", "700", "800", "900"]
            
            # 字体样式
            if "font-style" in styles:
                font_style = styles["font-style"].strip()
                style_state['italic'] = font_style == "italic"
            
            # 文本装饰
            if "text-decoration" in styles:
                decoration = styles["text-decoration"].strip()
                style_state['underline'] = "underline" in decoration
            
            # 字体大小
            if "font-size" in styles:
                size = styles["font-size"].strip()
                if size.endswith("pt"):
                    try:
                        pt_size = float(size[:-2])
                        style_state['font_size'] = Pt(pt_size)
                    except ValueError:
                        pass
                elif size.endswith("px"):
                    try:
                        px_size = float(size[:-2])
                        # 简单转换：1pt ≈ 1.33px
                        style_state['font_size'] = Pt(px_size / 1.33)
                    except ValueError:
                        pass
                elif size.endswith("em"):
                    try:
                        em_size = float(size[:-2])
                        # 1em ≈ 12pt (基础字号)
                        style_state['font_size'] = self._base_font_size * em_size
                    except ValueError:
                        pass
            
            # 字体类型
            if "font-family" in styles:
                font_family = styles["font-family"].split(',')[0].strip('"\' ')
                style_state['font_family'] = font_family
            
            # 行高
            if "line-height" in styles:
                line_height = styles["line-height"].strip()
                try:
                    if line_height.replace('.', '', 1).isdigit():
                        style_state['line_height'] = float(line_height)
                except ValueError:
                    pass
        
        # 处理直接属性（如font标签的color、face、size）
        if tag.has_attr('color'):
            color = tag['color'].strip()
            parsed_color = self._parse_color(color)
            if parsed_color:
                style_state['color'] = parsed_color
        
        if tag.has_attr('face'):
            font_family = tag['face'].split(',')[0].strip('"\' ')
            style_state['font_family'] = font_family
        
        if tag.has_attr('size'):
            size = tag['size'].strip()
            # 映射HTML字体大小到实际字号
            size_map = {
                '1': Pt(8),
                '2': Pt(10),
                '3': Pt(12),
                '4': Pt(14),
                '5': Pt(18),
                '6': Pt(24),
                '7': Pt(36)
            }
            # 处理相对大小（如+2, -1）
            if size.startswith('+') or size.startswith('-'):
                try:
                    # 获取当前字号
                    current_size = style_state.get('font_size', self._base_font_size)
                    if not isinstance(current_size, Pt):
                        current_size = self._base_font_size
                    
                    # 计算相对变化
                    change = int(size)
                    # 每级变化约20%
                    new_size = current_size.pt * (1.0 + 0.2 * change)
                    style_state['font_size'] = Pt(new_size)
                except ValueError:
                    pass
            elif size in size_map:
                style_state['font_size'] = size_map[size]
            else:
                try:
                    # 尝试直接解析为pt值
                    pt_size = float(size)
                    style_state['font_size'] = Pt(pt_size)
                except ValueError:
                    pass

    def _parse_style_string(self, style_str: str) -> Dict[str, str]:
        """解析样式字符串为字典"""
        styles = {}
        for declaration in style_str.split(';'):
            declaration = declaration.strip()
            if not declaration:
                continue
            if ':' in declaration:
                key, value = declaration.split(':', 1)
                styles[key.strip().lower()] = value.strip()
        return styles

    def _parse_color(self, color_str: str) -> Optional[RGBColor]:
        """解析颜色值，支持多种格式 - 优化十六进制颜色解析"""
        if not color_str:
            return None
            
        color_str = color_str.strip().lower()
        
        # 支持颜色名称
        color_names = {
            'black': RGBColor(0, 0, 0),
            'white': RGBColor(255, 255, 255),
            'red': RGBColor(255, 0, 0),
            'green': RGBColor(0, 128, 0),
            'blue': RGBColor(0, 0, 255),
            'yellow': RGBColor(255, 255, 0),
            'purple': RGBColor(128, 0, 128),
            'orange': RGBColor(255, 165, 0),
            'gray': RGBColor(128, 128, 128),
            'grey': RGBColor(128, 128, 128),
            'silver': RGBColor(192, 192, 192),
            'maroon': RGBColor(128, 0, 0),
            'olive': RGBColor(128, 128, 0),
            'lime': RGBColor(0, 255, 0),
            'aqua': RGBColor(0, 255, 255),
            'teal': RGBColor(0, 128, 128),
            'navy': RGBColor(0, 0, 128),
            'fuchsia': RGBColor(255, 0, 255),
        }
        if color_str in color_names:
            return color_names[color_str]
        
        # 处理带#前缀的十六进制颜色
        if color_str.startswith('#'):
            hex_str = color_str[1:]
            
            # 支持3位简写格式
            if len(hex_str) == 3:
                try:
                    r = int(hex_str[0]*2, 16)
                    g = int(hex_str[1]*2, 16)
                    b = int(hex_str[2]*2, 16)
                    return RGBColor(r, g, b)
                except ValueError:
                    return None
            
            # 处理6位格式
            if len(hex_str) == 6:
                try:
                    r = int(hex_str[0:2], 16)
                    g = int(hex_str[2:4], 16)
                    b = int(hex_str[4:6], 16)
                    return RGBColor(r, g, b)
                except ValueError:
                    return None
        
        # 支持rgb/rgba格式
        if color_str.startswith('rgb('):
            match = re.match(r'rgb\((\d+),\s*(\d+),\s*(\d+)\)', color_str)
            if match:
                return RGBColor(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        
        if color_str.startswith('rgba('):
            match = re.match(r'rgba\((\d+),\s*(\d+),\s*(\d+),\s*[\d.]+\)', color_str)
            if match:
                return RGBColor(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        
        return None

    def _apply_run_style(self, run: Run, style_state: Dict[str, Any]):
        """将样式状态应用到Run对象 - 优化颜色处理"""
        # 字体样式
        run.bold = style_state.get('bold', False)
        run.italic = style_state.get('italic', False)
        run.underline = style_state.get('underline', False)
        
        # 字体大小
        font_size = style_state.get('font_size')
        if font_size:
            run.font.size = font_size
        
        # 字体颜色 - 只在有设置时应用
        color = style_state.get('color')
        if color:
            run.font.color.rgb = color
        
        # 应用字体
        font_family = style_state.get('font_family')
        if font_family:
            run.font.name = font_family
            # 设置中文字体
            rPr = run._element.get_or_add_rPr()
            rPr.rFonts.set(qn('w:eastAsia'), font_family)
        else:
            # 应用默认字体（不设置颜色）
            self.apply_default_font(run)

    def apply_default_font(self, run: Run):
        """应用默认字体（不设置颜色）"""
        run.font.name = DocxFontEnum.TIMES_NEW_ROMAN
        rPr = run._element.get_or_add_rPr()
        rPr.rFonts.set(qn('w:eastAsia'), DocxFontEnum.SONG_TI)

    def apply_paragraph_styles(self, paragraph: Paragraph, style_str: str):
        """将 HTML 段落的 style 应用于 docx 段落"""
        styles = self._parse_style_string(style_str)

        # 行间距
        if "line-height" in styles:
            line_height = styles["line-height"].strip()
            try:
                if line_height.replace('.', '', 1).isdigit():
                    paragraph.paragraph_format.line_spacing = float(line_height)
            except ValueError:
                pass

        # 段前段后间距
        if "margin-top" in styles or "margin-bottom" in styles:
            if "margin-top" in styles:
                mt = styles["margin-top"].strip()
                if mt.endswith("pt"):
                    try:
                        paragraph.paragraph_format.space_before = Pt(float(mt[:-2]))
                    except ValueError:
                        pass
            if "margin-bottom" in styles:
                mb = styles["margin-bottom"].strip()
                if mb.endswith("pt"):
                    try:
                        paragraph.paragraph_format.space_after = Pt(float(mb[:-2]))
                    except ValueError:
                        pass

        # 文本对齐
        if "text-align" in styles:
            alignment = self.map_text_align_to_docx(styles["text-align"])
            paragraph.alignment = alignment

    def handle_table(self, doc: Document, table_tag: Tag, parent_style: Dict[str, Any]):
        rows = table_tag.find_all('tr')
        if not rows:
            return

        # 计算最大列数
        num_cols = max(len(row.find_all(['td', 'th'])) for row in rows)
        table = doc.add_table(rows=len(rows), cols=num_cols)
        
        # 应用表格样式
        if table_tag.has_attr('style'):
            self.apply_table_styles(table, table_tag['style'])
        else:
            table.style = 'Table Grid'

        for i, row in enumerate(rows):
            cells = row.find_all(['td', 'th'])
            for j, cell in enumerate(cells):
                if j >= num_cols:
                    continue

                doc_cell = table.cell(i, j)
                # 清除单元格默认段落
                for p in doc_cell.paragraphs:
                    p.clear()

                # 应用单元格样式
                if cell.has_attr('style'):
                    self.apply_cell_styles(doc_cell, cell['style'])
                
                # 创建段落
                paragraph = doc_cell.add_paragraph()
                self.set_paragraph_alignment(paragraph, cell)
                
                # 创建样式状态
                style_state = parent_style.copy()
                if cell.has_attr('style'):
                    self._update_style_from_attributes(style_state, cell)
                
                # 处理单元格内容
                self._process_children(doc, paragraph, cell, style_state)

                # 表头加粗
                if cell.name == 'th':
                    for run in paragraph.runs:
                        run.bold = True

    def apply_table_styles(self, table: Table, style_str: str):
        """应用表格样式"""
        styles = self._parse_style_string(style_str)
        
        # 设置表格宽度
        if "width" in styles:
            width = styles["width"].strip()
            if width.endswith("%"):
                try:
                    percent = float(width[:-1]) / 100.0
                    table.autofit = True
                    #table.width = Inches(8 * percent)
                except ValueError:
                    pass
        
        # 设置边框合并
        if "border-collapse" in styles and styles["border-collapse"] == "collapse":
            table.style = "Table Normal"

    def apply_cell_styles(self, cell: _Cell, style_str: str):
        """应用单元格样式"""
        styles = self._parse_style_string(style_str)
        
        # 设置单元格宽度
        if "width" in styles:
            width = styles["width"].strip()
            if width.endswith("%"):
                try:
                    percent = float(width[:-1]) / 100.0
                    cell.width = Inches(8 * percent) 
                except ValueError:
                    pass

    def handle_list(self, doc: Document, list_tag: Tag, parent_style: Dict[str, Any]):
        """处理 <ul> 或 <ol> 列表标签，支持嵌套和样式"""
        list_style = 'List Bullet' if list_tag.name == 'ul' else 'List Number'
        
        # 根据嵌套层级调整缩进
        indent_level = min(self._list_level, 5)  # 最大支持5级缩进

        for item in list_tag.find_all('li', recursive=False):
            paragraph = doc.add_paragraph(style=list_style)
            self._current_paragraph = paragraph

            # 设置缩进
            if indent_level > 1:
                paragraph.paragraph_format.left_indent = Inches(0.25 * indent_level)
                paragraph.paragraph_format.first_line_indent = Inches(-0.25)

            # 设置默认段落间距
            paragraph.paragraph_format.space_before = Pt(2)
            paragraph.paragraph_format.space_after = Pt(2)

            # 解析并应用 <li> 或 <ul> 的 style 样式
            if item.has_attr('style'):
                self.apply_paragraph_styles(paragraph, item['style'])
            elif list_tag.has_attr('style'):
                self.apply_paragraph_styles(paragraph, list_tag['style'])
            
            # 创建样式状态
            style_state = parent_style.copy()
            if item.has_attr('style'):
                self._update_style_from_attributes(style_state, item)
            
            # 解析内联内容
            self._process_children(doc, paragraph, item, style_state)
            self._current_paragraph = None

            # 处理嵌套列表
            nested_lists = item.find_all(['ul', 'ol'], recursive=False)
            for nested_list in nested_lists:
                self.handle_list(doc, nested_list, style_state)

    def set_paragraph_alignment(self, paragraph: Paragraph, tag: Tag):
        """设置段落对齐方式，支持align属性和style属性"""
        # 1. 检查align属性
        if tag.has_attr('align'):
            alignment = self.map_text_align_to_docx(tag['align'])
            paragraph.alignment = alignment
            return
        
        # 2. 检查style属性中的text-align
        if tag.has_attr('style'):
            styles = self._parse_style_string(tag['style'])
            if 'text-align' in styles:
                alignment = self.map_text_align_to_docx(styles['text-align'])
                paragraph.alignment = alignment
                return
        
        # 3. 检查父元素的样式
        parent = tag.parent
        while parent:
            if parent.has_attr('align'):
                alignment = self.map_text_align_to_docx(parent['align'])
                paragraph.alignment = alignment
                return
            if parent.has_attr('style'):
                styles = self._parse_style_string(parent['style'])
                if 'text-align' in styles:
                    alignment = self.map_text_align_to_docx(styles['text-align'])
                    paragraph.alignment = alignment
                    return
            parent = parent.parent

    @staticmethod
    def get_text_align_from_style(style_str):
        styles = dict([s.split(":", 1) for s in style_str.split(";") if ":" in s])
        return styles.get("text-align")

    @staticmethod
    def map_text_align_to_docx(text_align_value):
        if not text_align_value:
            return WD_ALIGN_PARAGRAPH.LEFT

        text_align_value = text_align_value.lower().strip()
        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        return align_map.get(text_align_value, WD_ALIGN_PARAGRAPH.LEFT)

    def _is_processed(self, tag: Tag) -> bool:
        return id(tag) in self._processed_tags

    def _mark_as_processed(self, tag: Tag):
        self._processed_tags.add(id(tag))
