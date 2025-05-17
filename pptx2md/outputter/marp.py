# Copyright 2025 Oscar Pellicer
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

import re
import urllib.parse
from typing import List, Optional, Union

from rapidfuzz import fuzz

from pptx2md.outputter.base import Formatter, DEFAULT_SLIDE_WIDTH_PX, DEFAULT_SLIDE_HEIGHT_PX, MARP_TARGET_WIDTH_PX, MARP_TARGET_HEIGHT_PX
from pptx2md.types import ParsedPresentation, SlideElement, ElementType, SlideType, ImageElement, FormulaElement # TextRun is used via get_formatted_runs
from pptx2md.utils import rgb_to_hex

class MarpFormatter(Formatter):
    # write outputs to marp markdown
    def __init__(self, config):
        super().__init__(config)
        self.esc_re1 = re.compile(r'([\|\*`])') # Marp specific escapes (e.g. | for tables)
        self.esc_re2 = re.compile(r'(<[^>]+>)')

    def put_header(self):
        css_content = """
@import url('https://fonts.googleapis.com/css2?family=Cabin:ital,wght@0,400..700;1,400..700&family=Lato:ital,wght@0,100;0,300;0,400;0,700;0,900;1,100;1,300;1,400;1,700;1,900&family=Libre+Baskerville:ital,wght@0,400;0,700;1,400&family=Lora:ital,wght@0,400..700;1,400..700&family=Montserrat:ital,wght@0,100..900;1,100..900&family=Mulish:ital,wght@0,200..1000;1,200..1000&family=Open+Sans:ital,wght@0,300..800;1,300..800&family=Roboto:ital,wght@0,100;0,300;0,400;0,500;0,700;0,900;1,100;1,300;1,400;1,500;1,700;1,900&family=Ubuntu:ital,wght@0,300;0,400;0,500;0,700;1,300;1,400;1,500;1,700&display=swap');

/* --- FONT SELECTION --- */
/* Uncomment one of the following font-family lines to change the slide font */
/* Default Marp font is usually sans-serif based */
section { font-family: 'Open Sans', sans-serif; }
/* section { font-family: 'Roboto', sans-serif; } */
/* section { font-family: 'Montserrat', sans-serif; } */
/* section { font-family: 'Cabin', sans-serif; } */
/* section { font-family: 'Ubuntu', sans-serif; } */
/* section { font-family: 'Lato', sans-serif; } */
/* section { font-family: 'Mulish', sans-serif; } */
/* section { font-family: 'Lora', serif; } */
/* section { font-family: 'Libre Baskerville', serif; } */

h1, h2, h3, h4, h5, h6 {
  font-weight: 400; /* Lighter weight for titles */
  color:rgb(45, 45, 45); /* Black color for titles */
}

section.small {
  font-size: 24px;
}
section.smaller {
  font-size: 20px;
}
section.smallest {
  font-size: 18px;
}

/* CSS for absolutely positioned elements */
.abs-pos {
  position: absolute;
}

img[alt~="center"] {
  display: block;
  margin: 0 auto;
}
img[alt~="left"] {
  float: left;
  margin-right: 1em;
  margin-bottom: 0.5em;
}
img[alt~="right"] {
  float: right;
  margin-left: 1em;
  margin-bottom: 0.5em;
}

.columns {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 2em;
}

.columns > div {
  overflow: hidden;
}

.figure-container {
  margin-bottom: 1em;
}

.figure-container img {
  display: block;
  max-width: 100%;
  height: auto;
  margin-left: auto;
  margin-right: auto;
}

.figure-container.align-left {
  float: left;
  margin-right: 1em;
  margin-left: 0;
}
.figure-container.align-left img {
  margin-left: 0;
  margin-right: auto;
}

.figure-container.align-right {
  float: right;
  margin-left: 1em;
  margin-right: 0;
}
.figure-container.align-right img {
  margin-right: 0;
  margin-left: auto;
}

.figure-container.align-center {
  display: block;
  margin-left: auto;
  margin-right: auto;
}

.figure-container .figcaption,
.figure-container > em {
  display: block;
  font-size: 0.85em;
  color: #555;
  text-align: center;
  margin-top: 0.4em;
  line-height: 1.3;
  font-style: normal;
}
.figure-container > em {
  font-style: italic;
}
"""
        examples_comment_html = """<!--
  MANUAL LAYOUT USAGE EXAMPLES:

  Multi-column Layout:
  (Automatic for 'smaller'/'smallest' slides with short lines, title excluded)
  To manually create a two-column layout:

  <div class="columns">
  <div>

    * First column content
    * Can contain lists, paragraphs, etc.

  </div>
  <div>

    * Second column content
    * Each column is equally sized by default.

  </div>
  </div>

  Figures with Captions:
  Wrap an image and its caption in a 'figure-container'.
  Use 'align-left', 'align-right', or 'align-center' for positioning.
  Set width on the container for floated figures.
  Caption can be <p class="figcaption">...</p> or <em>...</em>.

  Example (right-floated):
  <div class="figure-container align-right" style="width: 252px;">
    <img src="path/to/your/image.png" alt="Alt text" width="252">
    <p class="figcaption">This is the caption.</p>
  </div>

  Example (centered):
  <div class="figure-container align-center" style="max-width: 500px;">
    <img src="path/to/your/image.png" alt="Alt text">
    <em>A simple italic caption.</em>
  </div>

  Absolute Positioning:
  Use a <div> with class="abs-pos" and inline styles for positioning.
  Coordinates are relative to the slide. (0,0) is top-left.
  Marp slide default is 1280x720px.

  Example (Image):
  <div class="abs-pos" style="left: 50px; top: 100px; width: 300px; height: 200px; z-index: 1;">
    <img src="image.png" alt="My absolutely positioned image" style="width: 100%; height: 100%; object-fit: cover;">        
  </div>

  Example (Text):
  <div class="abs-pos" style="left: 400px; top: 150px; width: 250px; padding: 10px; background-color: lightblue; z-index: 2;">
    This is some text placed absolutely.
  </div>

  Example (Text on top of an image):
  First, the image (lower z-index or default)
  <div class="abs-pos" style="left: 100px; top: 200px; width: 400px; z-index: 5;">
    <img src="background_image.jpg" alt="Background" style="width: 100%; height: auto;">
  </div>

  Then, the text (higher z-index)
  <div class="abs-pos" style="left: 120px; top: 220px; width: 360px; color: white; font-size: 24px; text-align: center; z-index: 10;">
    Text overlaying the image.
  </div>
-->"""

        self.write(f'''---
marp: true
theme: default
paginate: true
html: true
---

<style>
{css_content.strip()}
</style>

{examples_comment_html}

''')

    def _put_elements_on_slide(self, elements: List[SlideElement], is_continued_slide: bool = False):
        last_element_type: Optional[ElementType] = None
        for element_idx, element in enumerate(elements):
            current_content_str = ""
            if element.type in [ElementType.Title, ElementType.Paragraph, ElementType.ListItem]:
                if isinstance(element.content, list):
                    current_content_str = self.get_formatted_runs(element.content)
                elif isinstance(element.content, str):
                    current_content_str = self.get_escaped(element.content) # Marp needs escaping for its syntax

            match element.type:
                case ElementType.Title:
                    title_text = current_content_str.strip()
                    if title_text:
                        if not (is_continued_slide and element_idx == 0):
                            is_similar_to_last = False
                            if self.last_title_info and self.last_title_info[1] == element.level and \
                               fuzz.ratio(self.last_title_info[0], title_text, score_cutoff=92):
                                is_similar_to_last = True

                            if is_similar_to_last:
                                if self.config.keep_similar_titles:
                                    effective_title = f'{title_text}' # (cont.) removed for Marp simpler logic
                                    self.put_title(effective_title, element.level)
                                    self.last_title_info = (effective_title, element.level)
                            else:
                                self.put_title(title_text, element.level)
                                self.last_title_info = (title_text, element.level)
                case ElementType.ListItem:
                    if not (last_element_type == ElementType.ListItem):
                        self.put_list_header()
                    self.put_list(current_content_str, element.level)
                case ElementType.Paragraph:
                    self.put_para(current_content_str)
                case ElementType.Image:
                    if isinstance(element, ImageElement):
                        self.put_image(element) # Marp put_image expects ImageElement
                case ElementType.Table:
                    if element.content:
                        table_data = [[self.get_formatted_runs(cell) if isinstance(cell, list) else self.get_escaped(str(cell)) for cell in row] for row in element.content]
                        self.put_table(table_data)
                case ElementType.CodeBlock:
                    code_content = getattr(element, 'content', '')
                    code_lang = getattr(element, 'language', None)
                    self.put_code_block(code_content, code_lang)
                case ElementType.Formula:
                    if isinstance(element, FormulaElement):
                        self.put_formula(element) # Base formula for $$...$$

            last_element_type = element.type

        if last_element_type == ElementType.ListItem:
            self.put_list_footer()

    def output(self, presentation_data: ParsedPresentation):
        self.put_header() # Writes directly to file
        self.last_title_info = None # Reset for each presentation

        pres_original_slide_width_px = self.config.slide_width_px or DEFAULT_SLIDE_WIDTH_PX

        num_total_slides = len(presentation_data.slides)
        marp_slide_counter = 0

        for slide_idx, slide in enumerate(presentation_data.slides):
            marp_slide_counter += 1

            initial_elements_for_slide: List[SlideElement] = []
            if slide.type == SlideType.General:
                initial_elements_for_slide = slide.elements
            elif slide.type == SlideType.MultiColumn:
                # For Marp, flatten MultiColumn for now. Title/floats from preface, then other content.
                initial_elements_for_slide = slide.preface + [el for col in slide.columns for el in col] 

            if not initial_elements_for_slide: 
                 if marp_slide_counter < num_total_slides : 
                    self.write("\n---\n\n") # Writes directly to file
                 continue

            # Separate title, floated images, and other content elements using base method
            main_title_element, floated_image_elements, other_content_elements = \
                self._separate_slide_elements(
                    initial_elements_for_slide,
                    pres_original_slide_width_px,
                    MARP_TARGET_WIDTH_PX # Marp's target rendering width
                )

            # Calculate overall slide metrics based on all initial elements for density class
            line_count, char_count, max_img_w, max_img_h, text_lines_for_avg, text_chars_for_avg = \
                self._get_slide_content_metrics(initial_elements_for_slide) # Use all original elements for density
            current_slide_class = self._get_slide_density_class(line_count)

            # Determine if slide qualifies for column splitting based on 'other_content_elements'
            initial_split_qualification = False
            if current_slide_class in ["smaller", "smallest"]:
                 # Calculate text metrics specifically for content that would go into columns
                _, _, _, _, other_text_lines, other_text_chars = \
                    self._get_slide_content_metrics(other_content_elements) # Metrics from non-title, non-floated
                
                if other_text_lines > 0: 
                    avg_line_length = other_text_chars / other_text_lines
                    if avg_line_length < self.config.marp_columns_line_length_threshold: # Use a config threshold
                        initial_split_qualification = True
            
            contains_table_in_other_content = False
            if initial_split_qualification:
                for element in other_content_elements:
                    if element.type == ElementType.Table:
                        contains_table_in_other_content = True
                        break
            
            actually_split_columns = initial_split_qualification and \
                                     len(other_content_elements) >= 2 and \
                                     not contains_table_in_other_content

            effective_slide_class = current_slide_class
            if actually_split_columns:
                if current_slide_class in ["smaller", "smallest"]:
                    effective_slide_class = "small" # Content is distributed
            
            if effective_slide_class:
                self.write(f"<!-- _class: {effective_slide_class} -->\n\n")

            if main_title_element:
                self._put_elements_on_slide([main_title_element], is_continued_slide=False)

            if floated_image_elements:
                self._put_elements_on_slide(floated_image_elements, is_continued_slide=False)

            if actually_split_columns:
                num_in_first_col = (len(other_content_elements) + 1) // 2
                first_half_elements = other_content_elements[:num_in_first_col]
                second_half_elements = other_content_elements[num_in_first_col:]

                self.write('<div class="columns">\n<div>\n\n')
                self._put_elements_on_slide(first_half_elements, is_continued_slide=False)
                self.write('\n</div>\n<div>\n\n')
                self._put_elements_on_slide(second_half_elements, is_continued_slide=False)
                self.write('\n</div>\n</div>\n\n')
            else:
                if other_content_elements: 
                    self._put_elements_on_slide(other_content_elements, is_continued_slide=False)
            
            if not self.config.disable_notes and slide.notes:
                self.write("<!--\n")
                for note_line in slide.notes:
                    self.write(f"{note_line}\n")
                self.write("-->\n\n")

            # Add slide separator if not the very last conceptual slide
            is_last_original_slide = (slide_idx == num_total_slides - 1)
            if not (is_last_original_slide) : # Add --- if not the true end
                 self.write("\n---\n\n") # Writes directly

        self.close()

    def put_title(self, text, level):
        self.write('#' * level + ' ' + text + '\n\n') # Writes directly

    def put_list(self, text, level):
        self.write('  ' * level + '* ' + text.strip() + '\n') # Writes directly

    def put_para(self, text):
        self.write(text + '\n\n') # Writes directly

    def put_image(self, element: Union[ImageElement, FormulaElement]):
        alt = element.alt_text if element.alt_text else ""
        quoted_path = urllib.parse.quote(element.path)
        
        marp_alt_text_keywords = []
        
        # Use configured slide dimensions, falling back to defaults, for scaling calculations.
        original_slide_width_px = self.config.slide_width_px or DEFAULT_SLIDE_WIDTH_PX
        original_slide_height_px = self.config.slide_height_px or DEFAULT_SLIDE_HEIGHT_PX

        # Get image's display dimensions from PowerPoint.
        ppt_display_width = element.display_width_px
        ppt_display_height = element.display_height_px

        # If display width is not available from PPT, but a default image width is configured,
        # use it and calculate corresponding height maintaining aspect ratio (if available).
        if ppt_display_width is None and self.config.image_width is not None:
            ppt_display_width = self.config.image_width
            if element.original_width_px and element.original_height_px and element.original_width_px > 0:
                 aspect_ratio = element.original_height_px / element.original_width_px
                 ppt_display_height = int(round(ppt_display_width * aspect_ratio))

        scaled_marp_display_width = None
        scaled_marp_display_height = None

        # Scale image dimensions from original slide context to Marp target dimensions.
        # Prioritize scaling based on width, then height, maintaining aspect ratio if possible.
        if ppt_display_width is not None and original_slide_width_px > 0:
            width_scale_factor = MARP_TARGET_WIDTH_PX / original_slide_width_px
            scaled_marp_display_width = int(round(ppt_display_width * width_scale_factor))

            if element.original_width_px and element.original_height_px and \
               element.original_width_px > 0 and scaled_marp_display_width > 0:
                image_aspect_ratio = element.original_height_px / element.original_width_px
                scaled_marp_display_height = int(round(scaled_marp_display_width * image_aspect_ratio))
            elif ppt_display_height is not None: # If aspect ratio unknown, scale height by same factor.
                scaled_marp_display_height = int(round(ppt_display_height * width_scale_factor))
        elif ppt_display_height is not None and original_slide_height_px > 0 and \
             element.original_width_px and element.original_height_px and element.original_height_px > 0 :
            # Fallback to scaling based on height if width-based scaling wasn't possible/applicable.
            height_scale_factor = MARP_TARGET_HEIGHT_PX / original_slide_height_px
            scaled_marp_display_height = int(round(ppt_display_height * height_scale_factor))
            if element.original_width_px > 0 and element.original_height_px > 0 : 
                image_aspect_ratio_inv = element.original_width_px / element.original_height_px
                scaled_marp_display_width = int(round(scaled_marp_display_height * image_aspect_ratio_inv))

        current_display_width = scaled_marp_display_width
        current_display_height = scaled_marp_display_height
        
        # Add Marp sizing keywords (w:, h:) if dimensions are determined.
        if current_display_width is not None and current_display_width > 0:
            marp_alt_text_keywords.append(f'w:{current_display_width}px') 
        # if current_display_height is not None and current_display_height > 0:
        #     marp_alt_text_keywords.append(f'h:{current_display_height}px')

        # Determine position hint (left, center, right) based on scaled image position and size.
        slide_width_for_hinting = MARP_TARGET_WIDTH_PX
        position_hint = None
        
        scaled_left_px = None
        if element.left_px is not None and original_slide_width_px > 0:
            scaled_left_px = int(round(element.left_px * (MARP_TARGET_WIDTH_PX / original_slide_width_px)))

        if scaled_left_px is not None and current_display_width is not None:
            image_center_x = scaled_left_px + (current_display_width / 2)
            # slide_center_x = slide_width_for_hinting / 2
            # center_threshold = slide_width_for_hinting * 0.10 # 10% threshold for centering
            
            # Define boundaries for "left" and "right" thirds of the slide.
            left_third_boundary = slide_width_for_hinting / 3
            right_third_boundary = 2 * slide_width_for_hinting / 3

            if left_third_boundary < image_center_x < right_third_boundary:
                position_hint = "center" 
            elif image_center_x < left_third_boundary: 
                position_hint = "left"
            elif image_center_x > right_third_boundary: 
                position_hint = "right"

        # Use the calculated position_hint, or fallback to a hint provided on the element itself.
        effective_position_hint = position_hint or getattr(element, 'position_hint', None)
        
        if effective_position_hint:
            if effective_position_hint == "center":
                marp_alt_text_keywords.append("center") 
            elif effective_position_hint == "left":
                 marp_alt_text_keywords.append("left")
            elif effective_position_hint == "right":
                 marp_alt_text_keywords.append("right")

        # Construct the final alt text string for Marp.
        # Order is important: [bg/positioning] [original alt text] [w:/h: sizing keywords (if not bg)].
        ordered_alt_keywords = []
        
        # Add "bg" and its associated positioning keywords first.
        if "bg" in marp_alt_text_keywords: ordered_alt_keywords.append("bg")
        # Handle specific "bg left" and "bg right" by ensuring correct order.
        if "bg left" in " ".join(marp_alt_text_keywords): ordered_alt_keywords = ["bg", "left"] 
        elif "bg right" in " ".join(marp_alt_text_keywords): ordered_alt_keywords = ["bg", "right"]
        
        # Add non-background positioning keywords ("center", "left", "right").
        has_bg_keyword = False # Disable background images for now
        if not has_bg_keyword:
            if "center" in marp_alt_text_keywords and "center" not in ordered_alt_keywords: ordered_alt_keywords.append("center")
            if "left" in marp_alt_text_keywords and "left" not in ordered_alt_keywords: ordered_alt_keywords.append("left")
            if "right" in marp_alt_text_keywords and "right" not in ordered_alt_keywords: ordered_alt_keywords.append("right")

        if alt:
            ordered_alt_keywords.append(alt)
            
        # Add sizing keywords (w:, h:) last, unless it's a background image.
        if not has_bg_keyword:
            for kw in marp_alt_text_keywords:
                if (kw.startswith("w:") or kw.startswith("h:")) and kw not in ordered_alt_keywords:
                    ordered_alt_keywords.append(kw)
        
        final_marp_alt_string = " ".join(ordered_alt_keywords).strip()

        # Output the image using Marp's Markdown syntax.
        self.write(f'![{final_marp_alt_string}]({quoted_path})\n\n') # Writes directly

    def put_code_block(self, code: str, language: Optional[str]):
        lang_tag = language if language else ""
        self.write(f'```{lang_tag}\n{code.strip()}\n```\n\n') # Writes directly

    def get_inline_code(self, text: str) -> str:
        # First, escape Marp-specific characters within the code text itself.
        # This handles `|` -> `\|`, `*` -> `\*`, `` ` `` -> `\``, etc.
        escaped_text = self.get_escaped(text)
        # Then, wrap the Marp-escaped text in single backticks for inline code.
        return f'`{escaped_text}`'

    def get_accent(self, text):
        return self._format_text_with_delimiters(text, '*', '*')

    def get_strong(self, text):
        return self._format_text_with_delimiters(text, '**', '**')

    def get_colored(self, text, rgb):
        # Standard HTML for color, Marp should support it
        return '<span style="color:%s">%s</span>' % (rgb_to_hex(rgb), text)

    def get_hyperlink(self, text, url):
        return '[' + text + '](' + url + ')'

    def esc_repl(self, match):
        return '\\' + match.group(0)

    def get_escaped(self, text):
        if self.config.disable_escaping:
            return text
        # Replace problematic Unicode characters first
        text = text.replace('\u000B', ' ').replace('\u000C', ' ')
        text = re.sub(self.esc_re1, self.esc_repl, text)
        text = re.sub(self.esc_re2, self.esc_repl, text)
        return text