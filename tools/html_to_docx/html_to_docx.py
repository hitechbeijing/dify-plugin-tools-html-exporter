import io
import logging
import re
from typing import Generator, Dict, Any, Optional

from bs4 import BeautifulSoup, Tag
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from tools.utils.md_utils import MarkdownUtils
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
        
        # 只处理直接子元素（块级元素）
        for child in body.children:
            if isinstance(child, Tag):
                self.process_tag(doc, child)

    def process_tag(self, doc: Document, tag: Tag):
        if self._is_processed(tag):
            return

        name = tag.name
        self._mark_as_processed(tag)

        if name in ['p', 'div']:
            # 创建新段落
            paragraph = doc.add_paragraph()

            # 设置段落对齐方式
            self.set_paragraph_alignment(paragraph, tag)

            # 默认段前段后间距
            paragraph.paragraph_format.space_before = Pt(12)
            paragraph.paragraph_format.space_after = Pt(12)

            # 处理 style 样式
            if tag.has_attr('style'):
                self.apply_paragraph_styles(paragraph, tag['style'])

            # 递归解析段落内容
            self.parse_inline_content(doc, paragraph, tag)

        elif name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            level = int(name[1])
            heading = doc.add_heading(level=level)
            for run in heading.runs:
                run.clear()
            self.parse_inline_content(doc, heading, tag)
            for run in heading.runs:
                run.bold = True
                run.font.size = Pt(18 - level * 2)

        elif name == 'table':
            self.handle_table(doc, tag)

        elif name in ['ul', 'ol']:
            self.handle_list(doc, tag)

        elif name == 'li':  # 新增对li标签的处理
            # 列表项由handle_list处理，这里跳过
            return

        elif name == 'br':
            doc.add_paragraph('\n')

        else:
            # 只处理有实际文本内容的非内联标签
            text = tag.get_text().strip()
            if text and name not in ['span', 'b', 'strong', 'i', 'em', 'u', 'small']:
                paragraph = doc.add_paragraph(text)
                self.set_paragraph_alignment(paragraph, tag)

    def parse_inline_content(self, doc: Document, paragraph: Paragraph, parent_tag: Tag):
        """使用样式状态机处理内联内容，支持嵌套标签"""
        # 初始化样式状态
        style_state = {
            'bold': False,
            'italic': False,
            'underline': False,
            'font_size': None,  # None表示使用基础字号
            'color': None
        }
        
        # 递归处理所有子节点
        self._process_children(doc, paragraph, parent_tag, style_state)

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
                elif tag_name == 'span' and child.has_attr('style'):
                    # 处理span的内联样式
                    self._update_style_from_attributes(current_style, child)
                
                # 处理块级标签（如嵌套的p/div）
                if tag_name in ['p', 'div']:
                    self.process_tag(doc, child)
                else:
                    # 递归处理子元素
                    self._process_children(doc, paragraph, child, current_style)
                    
            elif child.string:
                # 文本节点，应用当前样式
                text = child.string
                if text.strip() or text == '\n':  # 保留换行符
                    run = paragraph.add_run(text)
                    self._apply_run_style(run, parent_style)

    def _update_style_from_attributes(self, style_state: Dict[str, Any], tag: Tag):
        """从标签属性更新样式状态"""
        if tag.has_attr('style'):
            styles = dict([s.split(":", 1) for s in tag['style'].split(";") if ":" in s])
            
            # 字体颜色
            if "color" in styles:
                color = styles["color"].strip()
                style_state['color'] = self._parse_color(color)
            
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

    def _parse_color(self, color_str: str) -> Optional[RGBColor]:
        """解析16进制颜色值，支持3位和6位格式"""
        color_str = color_str.strip().lower()
        
        # 移除#前缀
        if color_str.startswith('#'):
            color_str = color_str[1:]
        
        # 支持rgb/rgba格式
        if color_str.startswith('rgb('):
            match = re.match(r'rgb\((\d+),\s*(\d+),\s*(\d+)\)', color_str)
            if match:
                return RGBColor(int(match.group(1)), int(match.group(2)), int(match.group(3)))
            return None
        
        # 支持3位简写格式
        if len(color_str) == 3:
            color_str = ''.join([c*2 for c in color_str])
        
        # 处理6位格式
        if len(color_str) == 6:
            try:
                r = int(color_str[0:2], 16)
                g = int(color_str[2:4], 16)
                b = int(color_str[4:6], 16)
                return RGBColor(r, g, b)
            except ValueError:
                return None
        
        return None

    def _apply_run_style(self, run: Run, style_state: Dict[str, Any]):
        """将样式状态应用到Run对象"""
        # 字体样式
        run.bold = style_state.get('bold', False)
        run.italic = style_state.get('italic', False)
        run.underline = style_state.get('underline', False)
        
        # 字体大小
        font_size = style_state.get('font_size')
        if font_size:
            run.font.size = font_size
        
        # 字体颜色
        color = style_state.get('color')
        if color:
            run.font.color.rgb = color
        
        # 应用默认字体
        self.apply_default_font(run)

    def apply_paragraph_styles(self, paragraph: Paragraph, style_str: str):
        """将 HTML 段落的 style 应用于 docx 段落"""
        styles = dict([s.split(":", 1) for s in style_str.split(";") if ":" in s])

        # 字体颜色
        if "color" in styles:
            color = self._parse_color(styles["color"])
            if color:
                for run in paragraph.runs:
                    run.font.color.rgb = color

        # 字体大小
        if "font-size" in styles:
            size = styles["font-size"].strip()
            if size.endswith("pt"):
                try:
                    pt_size = float(size[:-2])
                    for run in paragraph.runs:
                        run.font.size = Pt(pt_size)
                except ValueError:
                    pass

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

    def apply_default_font(self, run: Run):
        run.font.name = DocxFontEnum.TIMES_NEW_ROMAN
        rPr = run._element.get_or_add_rPr()
        rPr.rFonts.set(qn('w:eastAsia'), DocxFontEnum.SONG_TI)

    def handle_table(self, doc: Document, table_tag: Tag):
        rows = table_tag.find_all('tr')
        if not rows:
            return

        num_cols = max(len(row.find_all(['td', 'th'])) for row in rows)
        table = doc.add_table(rows=len(rows), cols=num_cols)
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

                paragraph = doc_cell.add_paragraph()
                self.set_paragraph_alignment(paragraph, cell)
                self.parse_inline_content(doc, paragraph, cell)

                # 表头加粗
                if cell.name == 'th':
                    for run in paragraph.runs:
                        run.bold = True

    def handle_list(self, doc: Document, list_tag: Tag):
        """处理 <ul> 或 <ol> 列表标签，支持嵌套和样式"""
        list_style = 'List Bullet' if list_tag.name == 'ul' else 'List Number'

        for item in list_tag.find_all('li', recursive=False):
            paragraph = doc.add_paragraph(style=list_style)

            # 设置默认段落间距
            paragraph.paragraph_format.space_before = Pt(4)
            paragraph.paragraph_format.space_after = Pt(4)

            # 解析并应用 <li> 或 <ul> 的 style 样式
            if item.has_attr('style'):
                self.apply_paragraph_styles(paragraph, item['style'])
            elif list_tag.has_attr('style'):
                self.apply_paragraph_styles(paragraph, list_tag['style'])

            # 解析内联内容
            self.parse_inline_content(doc, paragraph, item)

            # 处理嵌套列表
            nested_lists = item.find_all(['ul', 'ol'], recursive=True)
            for nested_list in nested_lists:
                self.handle_list(doc, nested_list)

    def set_paragraph_alignment(self, paragraph: Paragraph, tag: Tag):
        style_attr = tag.attrs.get("style", "")
        text_align = self.get_text_align_from_style(style_attr)

        # 检查父元素的样式
        parent = tag.parent
        while parent and not text_align:
            if parent.has_attr('style'):
                style_attr = parent.attrs.get("style", "")
                text_align = self.get_text_align_from_style(style_attr)
            parent = parent.parent

        if text_align:
            alignment = self.map_text_align_to_docx(text_align)
            paragraph.alignment = alignment

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