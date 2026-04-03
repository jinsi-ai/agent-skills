# deskclaw-pptx core: presentation, content, structural

from .presentation import (
    create_presentation,
    create_presentation_from_template,
    get_presentation_info,
    set_core_properties,
    get_template_file_info,
)
from .content import (
    add_slide,
    get_slide_info,
    extract_slide_text,
    extract_presentation_text,
    populate_placeholder,
    add_bullet_points,
)
from .structural import (
    add_table,
    add_shape,
    format_table_cell,
    add_chart,
)

__all__ = [
    "create_presentation",
    "create_presentation_from_template",
    "get_presentation_info",
    "set_core_properties",
    "get_template_file_info",
    "add_slide",
    "get_slide_info",
    "extract_slide_text",
    "extract_presentation_text",
    "populate_placeholder",
    "add_bullet_points",
    "add_table",
    "add_shape",
    "format_table_cell",
    "add_chart",
]
