# Copyright 2024 Liu Siyao
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

from __future__ import annotations

from enum import Enum
from pathlib import Path
from typing import List, Optional, Union, Any, Tuple
from dataclasses import dataclass, field

from pydantic import BaseModel


@dataclass
class ConversionConfig:
    """Configuration for PowerPoint to Markdown conversion."""

    pptx_path: Path
    """Path to the .pptx file to be converted"""

    output_path: Path
    """Path of the output file"""

    output_dir: Path
    """Where to put the output file(s)"""

    image_dir: Optional[Path]
    """Where to put images extracted"""

    title_path: Optional[Path] = None
    """Path to the custom title list file"""

    image_width: Optional[int] = None
    """Maximum image width in px"""

    image_height: Optional[int] = None
    """Maximum image height in px"""

    disable_image: bool = False
    """Disable image extraction"""

    disable_wmf: bool = False
    """Keep wmf formatted image untouched (avoid exceptions under linux)"""

    disable_color: bool = False
    """Do not add color HTML tags"""

    disable_escaping: bool = False
    """Do not attempt to escape special characters"""

    disable_notes: bool = False
    """Do not add presenter notes"""

    enable_slides: bool = False
    """Deliniate slides with `\n---\n`"""

    is_md: bool = False
    """Generate output as standard markdown"""

    is_wiki: bool = False
    """Generate output as wikitext (TiddlyWiki)"""

    is_mdk: bool = False
    """Generate output as madoko markdown"""

    is_qmd: bool = False
    """Generate output as quarto markdown presentation"""

    is_marp: bool = False
    """Generate output as marp markdown"""

    is_beamer: bool = False
    """Generate output as beamer tex"""

    is_json: bool = False
    """Generate output as the raw .pptx abstract syntax tree in JSON format"""

    min_block_size: int = 0
    """The minimum character number of a text block to be converted"""

    page: Optional[int] = None
    """Only convert the specified page"""

    custom_titles: List[Any] = field(default_factory=list)
    """Mapping of custom titles to their heading levels"""

    try_multi_column: bool = False
    """Try to detect multi-column slides"""

    keep_similar_titles: bool = False
    """Keep similar titles (allow for repeated slide titles - One or more - Add (cont.) to the title)"""

    slide_width_px: Optional[int] = None
    """Actual slide width in pixels"""

    slide_height_px: Optional[int] = None
    """Actual slide height in pixels"""

    apply_cropping_in_parser: bool = True
    """Apply cropping directly in the parser, modifying the saved image."""

    disable_captions: bool = True
    """Disable captions"""

    disable_image_wrapping: bool = False
    """Disable image wrapping"""

    # Thresholds for font size adjustments / density classes
    # These should ideally match or be derived from LINES_NORMAL_MAX etc. if those are global
    small_font_line_threshold: int = 8 # Example value
    smaller_font_line_threshold: int = 12 # Example value
    smallest_font_line_threshold: int = 18 # Example value

    # Thresholds for attempting to split content into columns
    marp_columns_line_length_threshold: int = 40
    beamer_columns_line_length_threshold: int = 40

    # Equivalents for density calculation in _get_slide_content_metrics
    image_density_line_equivalent: int = 3
    table_row_density_line_equivalent: int = 1

    def __post_init__(self):
        if isinstance(self.pptx_path, str):
            self.pptx_path = Path(self.pptx_path)
        if isinstance(self.output_path, str):
            self.output_path = Path(self.output_path)

class ElementType(str, Enum):
    Title = "Title"
    ListItem = "ListItem"
    Paragraph = "Paragraph"
    Image = "Image"
    Table = "Table"
    CodeBlock = "CodeBlock"
    Formula = "Formula"

class TextStyle(BaseModel):
    is_accent: bool = False
    is_strong: bool = False
    color_rgb: Optional[tuple[int, int, int]] = None
    hyperlink: Optional[str] = None
    is_code: bool = False
    is_math: bool = False


class TextRun(BaseModel):
    text: str
    style: TextStyle
    font_name: Optional[str] = None


class Position(BaseModel):
    left: float
    top: float
    width: float
    height: float


class BaseElement(BaseModel):
    type: ElementType
    position: Optional[Position] = None
    style: Optional[TextStyle] = None


class TitleElement(BaseElement):
    type: ElementType = ElementType.Title
    content: str
    level: int


class ListItemElement(BaseElement):
    type: ElementType = ElementType.ListItem
    content: List[TextRun]
    level: int = 1


class ParagraphElement(BaseElement):
    type: ElementType = ElementType.Paragraph
    content: List[TextRun]


class ImageElement(BaseElement):
    type: ElementType = ElementType.Image
    path: str
    width: Optional[int] = None
    original_ext: str = ""  # For tracking original file extension (e.g. wmf)
    alt_text: str = ""  # For accessibility
    display_width_px: Optional[int] = None
    display_height_px: Optional[int] = None
    original_width_px: Optional[int] = None
    original_height_px: Optional[int] = None
    original_filename: Optional[str] = None
    left_px: Optional[int] = None
    top_px: Optional[int] = None
    rotation: Optional[float] = None
    crop_left_pct: Optional[float] = None
    crop_right_pct: Optional[float] = None
    crop_top_pct: Optional[float] = None
    crop_bottom_pct: Optional[float] = None


class TableElement(BaseElement):
    type: ElementType = ElementType.Table
    content: List[List[List[TextRun]]]  # rows -> cols -> rich text


class CodeBlockElement(BaseElement):
    type: ElementType = ElementType.CodeBlock
    content: str
    language: Optional[str] = None
    position: Optional[Tuple[float, float]] = None


class FormulaElement(BaseElement):
    type: ElementType = ElementType.Formula
    content: str
    position: Optional[Tuple[float, float]] = None


SlideElement = Union[
    TitleElement, ParagraphElement, ListItemElement, ImageElement, TableElement, CodeBlockElement, FormulaElement
]


class SlideType(str, Enum):
    MultiColumn = "MultiColumn"
    General = "General"


class MultiColumnSlide(BaseModel):
    type: SlideType = SlideType.MultiColumn
    preface: List[SlideElement]
    columns: List[SlideElement]
    notes: List[str] = []


class GeneralSlide(BaseModel):
    type: SlideType = SlideType.General
    elements: List[SlideElement]
    notes: List[str] = []


Slide = Union[GeneralSlide, MultiColumnSlide]


class ParsedPresentation(BaseModel):
    slides: List[Slide]
