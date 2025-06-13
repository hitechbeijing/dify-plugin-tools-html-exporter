from io import StringIO
from logging import Logger

import markdown
import pandas as pd

from tools.utils.md_utils import MarkdownUtils


class TableParser:
    @staticmethod
    def parse_md_to_tables(logger: Logger, md_text: str) -> list[pd.DataFrame]:
        """
        Parse markdown text to tables
        :param logger: Logger
        :param md_text: markdown text
        :return: list of tables
        """
        try:
            md_text = MarkdownUtils.strip_markdown_wrapper(md_text)
            html_str = markdown.markdown(text=md_text, extensions=['tables'])
            tables = pd.read_html(StringIO(html_str))
            if not tables or len(tables) < 1:
                raise ValueError("No available tables parsed from markdown text")
            return tables
        except Exception as e:
            msg = f"Failed to parse markdown to tables, exception: {str(e)}"
            logger.exception(msg)
            raise ValueError(msg)
