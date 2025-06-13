import re
from typing import Any


THINK_TAG_REGEX = re.compile(r'<think>.*?</think>', flags=re.DOTALL)


def get_html_text(tool_parameters: dict[str, Any],
                  is_remove_think_tag: bool = True,
                  is_normalize_line_breaks: bool = True,
                  ) -> str:
    html_text = tool_parameters.get("html_text")
    html_text = html_text.strip() if html_text else None
    if not html_text:
        raise ValueError("Empty input html_text")

    # remove think tag
    if is_remove_think_tag:
        html_text = THINK_TAG_REGEX.sub('', html_text)

    # normalize line breaks
    if is_normalize_line_breaks and "\\n" in html_text:
        html_text = html_text.replace("\\n", "\n")

    return html_text

def get_param_value(tool_parameters: dict[str, Any], param_name: str, default_value: Any = None) -> Any:
    param_value = tool_parameters.get(param_name, default_value)
    if not param_value:
        raise ValueError(f"Empty input {param_name}")

    return param_value
