from typing import Any

from dify_plugin import ToolProvider
from dify_plugin.errors.tool import ToolProviderCredentialValidationError

from tools.html_to_docx.html_to_docx import HtmlToDocxTool


class MdExporterProvider(ToolProvider):
    def _validate_credentials(self, credentials: dict[str, Any]) -> None:
        try:
            """
            IMPLEMENT YOUR VALIDATION HERE
            """
            tools = [
                HtmlToDocxTool,
            ]
            for tool in tools:
                tool.from_credentials({})
        except Exception as e:
            raise ToolProviderCredentialValidationError(str(e))
