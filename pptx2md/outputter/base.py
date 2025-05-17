# Substantial modifications made by Oscar Pellicer, 2025
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

import os
import re
import urllib.parse
from typing import List, Tuple, Optional, Union
import io
import abc

from rapidfuzz import fuzz

from pptx2md.types import ConversionConfig, ElementType, ParsedPresentation, SlideElement, SlideType, TextRun, ImageElement, FormulaElement, TextStyle
from pptx2md.utils import rgb_to_hex


# Global variables
LINES_NORMAL_MAX = 8
LINES_SMALL_MAX = 12
LINES_SMALLER_MAX = 18
LINES_SPLIT_TRIGGER = 18

DEFAULT_SLIDE_WIDTH_PX = 1600
DEFAULT_SLIDE_HEIGHT_PX = 900

MARP_TARGET_WIDTH_PX = 1280
MARP_TARGET_HEIGHT_PX = 720

class Formatter(abc.ABC):

    def __init__(self, config: ConversionConfig):
        os.makedirs(config.output_path.parent, exist_ok=True)
        self.ofile = open(config.output_path, 'w', encoding='utf8')
        self.config = config
        # For BeamerFormatter and potentially others that buffer output
        self._buffer = io.StringIO()
        self.last_title_info: Optional[Tuple[str, int]] = None # Common for title similarity logic

    def write(self, text: str):
        # Default write to buffer. Formatters writing directly to file can override.
        self._buffer.write(text)

    def _format_text_with_delimiters(self, text: str, open_delimiter: str, close_delimiter: str) -> str:
        if not text: # Handle empty string input early
            return text # Return original empty string

        # 1. Find leading whitespace
        leading_whitespace_count = 0
        for char_idx, char_val in enumerate(text):
            if not char_val.isspace():
                leading_whitespace_count = char_idx
                break
        else: # String is all whitespace
            return text
        leading_whitespace = text[:leading_whitespace_count]
        text_without_leading = text[leading_whitespace_count:]

        # Handle case where original text was only leading whitespace correctly
        if not text_without_leading: 
            return text 

        # 2. Find trailing whitespace (from text_without_leading)
        trailing_whitespace_count = 0
        for char_idx, char_val in enumerate(reversed(text_without_leading)):
            if not char_val.isspace():
                trailing_whitespace_count = char_idx
                break
        # No 'else' needed here. If text_without_leading is all whitespace, 
        # trailing_whitespace_count will remain 0, and core_text will become empty.
        
        core_text_end_index = len(text_without_leading) - trailing_whitespace_count
        core_text = text_without_leading[:core_text_end_index]
        trailing_whitespace = text_without_leading[core_text_end_index:]

        if not core_text: # If, after stripping both ends, core is empty
                          # This implies the original text (after leading strip) was all whitespace,
                          # or the original text itself was effectively empty of non-whitespace.
            return text # Return original text to avoid "open_delimiterclose_delimiter" for "   "
        
        return f"{leading_whitespace}{open_delimiter}{core_text}{close_delimiter}{trailing_whitespace}"

    def _get_slide_content_metrics(self, elements_list: List[SlideElement]) -> Tuple[int, int, Optional[int], Optional[int], int, int]:
        """Calculates number of semantic lines, total characters, max image dimensions,
           and specific text line/char counts for avg line length heuristic."""
        line_count = 0
        char_count = 0
        max_image_width: Optional[int] = 0
        max_image_height: Optional[int] = 0
        
        text_lines_for_avg_heuristic = 0
        text_chars_for_avg_heuristic = 0

        for element in elements_list:
            element_text_content_for_avg = ""
            is_text_for_avg_heuristic = False

            if element.type == ElementType.Title:
                line_count += 1
                content = element.content.strip() if isinstance(element.content, str) else ""
                char_count += len(content)
            elif element.type == ElementType.ListItem:
                line_count += 1
                text_lines_for_avg_heuristic += 1
                is_text_for_avg_heuristic = True
                if isinstance(element.content, list): # List[TextRun]
                    item_text = "".join(run.text for run in element.content)
                    char_count += len(item_text)
                    element_text_content_for_avg = item_text
                elif isinstance(element.content, str): 
                    char_count += len(element.content)
                    element_text_content_for_avg = element.content

            elif element.type == ElementType.Paragraph:
                line_count += 1 
                text_lines_for_avg_heuristic += 1
                is_text_for_avg_heuristic = True
                if isinstance(element.content, list): # List[TextRun]
                    para_text = "".join(run.text for run in element.content)
                    char_count += len(para_text)
                    element_text_content_for_avg = para_text
                elif isinstance(element.content, str): 
                     char_count += len(element.content)
                     element_text_content_for_avg = element.content
            
            elif element.type == ElementType.CodeBlock:
                line_count += (element.content.count('\n') + 1) if element.content else 1
                char_count += len(element.content)
            
            elif element.type == ElementType.Table:
                if element.content: 
                    line_count += len(element.content) 
                    for row in element.content:
                        for cell_runs in row: # Assuming cell is List[TextRun]
                            if isinstance(cell_runs, list):
                                for run in cell_runs:
                                    char_count += len(run.text)
                            elif isinstance(cell_runs, str): # Fallback if cell is string
                                char_count += len(cell_runs)
            
            elif element.type == ElementType.Image:
                if hasattr(element, 'display_width_px') and element.display_width_px is not None:
                    max_image_width = max(max_image_width or 0, element.display_width_px)
                if hasattr(element, 'display_height_px') and element.display_height_px is not None:
                    max_image_height = max(max_image_height or 0, element.display_height_px)

            if is_text_for_avg_heuristic:
                text_chars_for_avg_heuristic += len(element_text_content_for_avg.strip())

        return line_count, char_count, max_image_width, max_image_height, text_lines_for_avg_heuristic, text_chars_for_avg_heuristic

    def _get_slide_density_class(self, line_count: int) -> Optional[str]:
        """Determines a density class based on line count."""
        if line_count > LINES_SMALLER_MAX: return "smallest"
        if line_count > LINES_SMALL_MAX: return "smaller"
        if line_count > LINES_NORMAL_MAX: return "small"
        return None

    def output(self, presentation_data: ParsedPresentation):
        self.put_header()

        last_element_type: Optional[ElementType] = None # Changed from last_element to track type
        # self.last_title_info is already in __init__

        for slide_idx, slide in enumerate(presentation_data.slides):
            all_elements: List[SlideElement] = []
            if slide.type == SlideType.General:
                all_elements = slide.elements
            elif slide.type == SlideType.MultiColumn:
                all_elements = slide.preface + [el for col in slide.columns for el in col]


            for element in all_elements:
                if last_element_type and last_element_type == ElementType.ListItem and element.type != ElementType.ListItem:
                    self.put_list_footer()
                
                current_content_str = "" # Initialize for elements with text content
                if element.type in [ElementType.Title, ElementType.Paragraph, ElementType.ListItem]:
                    if isinstance(element.content, list) and all(isinstance(run, TextRun) for run in element.content):
                        current_content_str = self.get_formatted_runs(element.content)
                    elif isinstance(element.content, str):
                        # For base formatter, escaping might be formatter-specific.
                        # Let derived formatters handle escaping if text is string.
                        # For now, pass raw string to put_para, put_title, put_list
                        current_content_str = element.content 
                    # else: content might be of unexpected type for these elements

                match element.type:
                    case ElementType.Title:
                        title_text = element.content.strip() if isinstance(element.content, str) else current_content_str.strip()
                        if title_text:
                            is_similar_to_last = False
                            if self.last_title_info and self.last_title_info[1] == element.level and \
                               fuzz.ratio(self.last_title_info[0], title_text, score_cutoff=92):
                                is_similar_to_last = True
                            
                            if is_similar_to_last:
                                if self.config.keep_similar_titles:
                                    effective_title = f'{title_text} (cont.)'
                                    self.put_title(effective_title, element.level)
                                    self.last_title_info = (effective_title, element.level)
                                # else skip
                            else:
                                self.put_title(title_text, element.level)
                                self.last_title_info = (title_text, element.level)
                    case ElementType.ListItem:
                        if not (last_element_type and last_element_type == ElementType.ListItem):
                            self.put_list_header()
                        self.put_list(current_content_str, element.level)
                    case ElementType.Paragraph:
                        self.put_para(current_content_str)
                    case ElementType.Image:
                        # Base Formatter might not know how to put_image.
                        # This will be overridden by specific formatters.
                        # Pass the whole element for rich data.
                        if hasattr(self, 'put_image') and callable(self.put_image):
                            self.put_image(element) 
                    case ElementType.Table:
                        # Similar to image, specific formatters will implement.
                        # Pass processed cell content.
                        if hasattr(self, 'put_table') and callable(self.put_table):
                            table_content = [[self.get_formatted_runs(cell) if isinstance(cell, list) else str(cell) for cell in row] for row in element.content]
                            self.put_table(table_content)
                    case ElementType.CodeBlock:
                        # Base Formatter might not know how to put_code_block.
                        if hasattr(self, 'put_code_block') and callable(self.put_code_block):
                            self.put_code_block(element.content, element.language)
                    case ElementType.Formula:
                        if isinstance(element, FormulaElement):
                            if hasattr(self, 'put_formula') and callable(self.put_formula):
                                self.put_formula(element)
                last_element_type = element.type

            if last_element_type == ElementType.ListItem:
                self.put_list_footer()

            if not self.config.disable_notes and slide.notes:
                # Notes handling can be generic if put_para is well-defined.
                self.put_para('---') # Markdown-like separator for notes block
                for note_line in slide.notes:
                    # Notes are typically raw strings, may need escaping by formatter's put_para.
                    self.put_para(note_line) 

            if slide_idx < len(presentation_data.slides) - 1 and self.config.enable_slides:
                # Slide separator, also potentially formatter-specific.
                # Defaulting to Markdown style.
                self.put_para("\n---\n")


        self.close() # Ensure derived classes call super().close() if they override it for flushing buffer

    def put_header(self):
        pass

    def put_list_header(self):
        """Placeholder for list header. Override in derived formatters."""
        # For formatters like basic Markdown, this might involve an empty line for spacing.
        # self.put_para('') # Example if spacing is needed.
        pass

    def put_list_footer(self):
        """Placeholder for list footer. Override in derived formatters."""
        # For formatters like basic Markdown, this might involve an empty line for spacing.
        # self.put_para('') # Example if spacing is needed.
        pass

    def _styles_are_compatible(self, style1: Optional[TextStyle], style2: Optional[TextStyle]) -> bool:
        if style1 is None or style2 is None:
            return False # Should not happen with proper initialization
        return (style1.is_code == style2.is_code and
                style1.is_accent == style2.is_accent and
                style1.is_strong == style2.is_strong and
                style1.is_math == style2.is_math and 
                style1.hyperlink == style2.hyperlink and
                style1.color_rgb == style2.color_rgb)

    def _format_single_merged_run(self, text: str, style: TextStyle) -> str:
        if not text and not style.is_code and not style.is_math: # Allow empty code/math runs potentially
            return ""

        formatted_text = text # Start with raw text for this segment

        if style.is_code:
            # self.get_inline_code is responsible for its own handling of text
            return self.get_inline_code(formatted_text)

        # For math, apply math formatting and return. Assume math is exclusive of other styling here.
        if style.is_math:
            # Escaping might not be needed or desired for math content itself.
            # get_inline_math will handle delimiters and whitespace.
            return self.get_inline_math(formatted_text)

        # Process non-code, non-math text
        if not self.config.disable_escaping:
            formatted_text = self.get_escaped(formatted_text)
        
        # Apply strong and accent (bold and italic)
        # This order will result in accent (e.g., italics) being the inner markup
        # if both are applied, e.g., **_text_** or __*text*__
        if style.is_strong:
            formatted_text = self.get_strong(formatted_text)
        if style.is_accent:
            formatted_text = self.get_accent(formatted_text)
        
        if style.color_rgb and not self.config.disable_color:
            formatted_text = self.get_colored(formatted_text, style.color_rgb)
        
        if style.hyperlink:
            formatted_text = self.get_hyperlink(formatted_text, style.hyperlink)
            
        return formatted_text

    def get_formatted_runs(self, runs: List[TextRun]):
        if not runs:
            return ""

        output_segments: List[str] = []
        
        def _normalize_whitespace_in_run_text(text: str) -> str:
            # Replace non-breaking space (U+00A0) with a regular space (U+0020).
            normalized_text = text.replace('\u00A0', ' ')
            # Narrow No-Break Space
            normalized_text = normalized_text.replace('\u202F', ' ')
            # Remove vertical tab (\x0B) for all non-code, non-math text
            normalized_text = normalized_text.replace('\x0B', '')
            return normalized_text

        # Initialize with the first run
        current_merged_text = _normalize_whitespace_in_run_text(runs[0].text)
        current_style = runs[0].style

        for i in range(1, len(runs)):
            next_run = runs[i]
            normalized_next_text = _normalize_whitespace_in_run_text(next_run.text)

            if self._styles_are_compatible(current_style, next_run.style):
                # Styles are compatible, merge text
                current_merged_text += normalized_next_text
            else:
                # Styles differ, format the accumulated text and add to segments
                # Process if text exists or if it's an intentionally empty code/math run
                if current_merged_text or (current_style and (current_style.is_code or current_style.is_math)):
                    # Only sanitize vertical tab for non-code, non-math
                    if not (current_style and (current_style.is_code or current_style.is_math)):
                        current_merged_text = current_merged_text.replace('\x0B', '')
                    output_segments.append(self._format_single_merged_run(current_merged_text, current_style))
                # Start a new merged run
                current_merged_text = normalized_next_text
                current_style = next_run.style
        # Format the last accumulated run
        if current_merged_text or (current_style and (current_style.is_code or current_style.is_math)):
            if not (current_style and (current_style.is_code or current_style.is_math)):
                current_merged_text = current_merged_text.replace('\x0B', '')
            output_segments.append(self._format_single_merged_run(current_merged_text, current_style))
        final_text = "".join(output_segments)
        return final_text.strip()

    def put_para(self, text):
        pass

    def put_image(self, path, max_width): # Base signature, specific formatters might take ImageElement
        pass

    def put_table(self, table: List[List[str]]): # table cells are already formatted strings
        if not table or not table[0]:
            return # Handle empty table
        
        gen_table_row = lambda row: '| ' + ' | '.join([
            c.replace('\n', '<br />') if '`' not in c else c.replace('\n', ' ') 
            for c in row
        ]) + ' |'
        
        self.write(gen_table_row(table[0]) + '\n') 
        self.write(gen_table_row([':-' for _ in table[0]]) + '\n') 
        self.write('\n'.join([gen_table_row(row) for row in table[1:]]) + '\n\n')

    def put_code_block(self, code: str, language: Optional[str]):
        lang_tag = language if language else ""
        self.write(f'```{lang_tag}\n{code.strip()}\n```\n\n')

    def put_formula(self, element: FormulaElement):
        formatted_content = self._format_text_with_delimiters(element.content, "$$", "$$")
        self.write(f'{formatted_content}\n\n')

    def get_inline_code(self, text: str) -> str:
        """Formats text as inline code. Does not strip or escape input text.
           Handles literal backticks within the text by using a longer fence.
        """
        if not text: # If the original run text was empty.
            return ""

        # Find the longest sequence of backticks in the text
        longest_backtick_sequence = 0
        current_backtick_sequence = 0
        for char in text:
            if char == '`':
                current_backtick_sequence += 1
            else:
                longest_backtick_sequence = max(longest_backtick_sequence, current_backtick_sequence)
                current_backtick_sequence = 0
        longest_backtick_sequence = max(longest_backtick_sequence, current_backtick_sequence)

        # The fence should be one longer than the longest sequence found
        fence_len = longest_backtick_sequence + 1
        fence = '`' * fence_len

        # If the text starts or ends with a backtick, or is all backticks,
        # and the chosen fence is just one backtick,
        # then a space is needed to disambiguate (CommonMark spec).
        # Example: ` `` ` vs `` ` ``. If text is '`a`' and fence is '`', then '` `a` `'
        # However, if fence_len > 1, this space padding is generally not needed.
        # For simplicity and robustness with `fence_len > 1`, just add fence.
        # If `text` is just '`', `fence_len` will be 2, result "`` ` ``".
        # If `text` is 'a`b', `fence_len` will be 2, result "``a`b``".
        
        # A common strategy for CommonMark compliance with content starting/ending with backticks
        # or being all backticks, when the fence is a single backtick:
        if fence_len == 1 and (text.startswith('`') or text.endswith('`') or text.isspace()):
             return f"{fence} {text} {fence}"
        
        return f"{fence}{text}{fence}"

    def get_accent(self, text):
        return self._format_text_with_delimiters(text, '_', '_')

    def get_strong(self, text):
        return self._format_text_with_delimiters(text, '__', '__')

    def get_inline_math(self, text: str) -> str:
        if not text:
            return ""

        # 1. Isolate Overall Whitespace from the input `text`
        original_len = len(text)
        text_lstripped = text.lstrip()
        
        if not text_lstripped: # Original text was all whitespace or empty
            return text
            
        overall_lw_len = original_len - len(text_lstripped)
        overall_lw = text[:overall_lw_len]

        core_run_text = text_lstripped.rstrip() # This is the central content of the run
        
        overall_tw_len = len(text_lstripped) - len(core_run_text)
        overall_tw = text_lstripped[len(core_run_text):] if overall_tw_len > 0 else ""

        # 2. Identify Formula Candidate Payload from core_run_text
        formula_candidate_payload: str
        if (core_run_text.startswith('$') and core_run_text.endswith('$') and
            len(core_run_text) >= 2 and
            not (core_run_text.startswith('$$') and core_run_text.endswith('$$'))):
            formula_candidate_payload = core_run_text[1:-1]
        else:
            formula_candidate_payload = core_run_text
        
        # 3. Separate True Math Symbols from any trailing text within the candidate payload
        #    Example: if formula_candidate_payload is "x_s ", actual_math_symbols = "x_s", internal_trailing_text = " "
        actual_math_symbols = formula_candidate_payload.rstrip()
        internal_trailing_text = formula_candidate_payload[len(actual_math_symbols):]

        # 4. Clean the actual math symbols (e.g., strip leading spaces from them)
        #    Example: if actual_math_symbols was "  x_s", cleaned_math_symbols = "x_s"
        cleaned_math_symbols = actual_math_symbols.strip()

        # 5. Format the cleaned math symbols
        formatted_math: str
        if not cleaned_math_symbols:
            formatted_math = "$ $"  # Handle case of empty or all-space math content
        else:
            formatted_math = f"${cleaned_math_symbols}$"
            
        # 6. Reconstruct the string
        return f"{overall_lw}{formatted_math}{internal_trailing_text}{overall_tw}"

    def get_colored(self, text, rgb):
        # Standard HTML for color, Marp should support it
        return '<span style="color:%s">%s</span>' % (rgb_to_hex(rgb), text)

    def get_hyperlink(self, text, url):
        return '[' + text + '](' + url + ')'

    def get_escaped(self, text):
        return text

    def flush(self):
        # If writing directly to file, self.ofile.flush()
        # If using self._buffer, this might not be needed until close(),
        # or could write buffer to ofile and clear buffer.
        if hasattr(self.ofile, 'flush'):
             self.ofile.flush()


    def close(self):
        # Default implementation: write buffer to file, then close file.
        # Formatters directly writing to self.ofile should override self.write()
        # or ensure they don't use self._buffer.
        # BeamerFormatter example correctly uses self._buffer and writes it here.
        
        # If _buffer was used (e.g. by BeamerFormatter's self.write())
        buffered_content = self._buffer.getvalue()
        if buffered_content:
            # Perform any final sanitization on buffered_content if needed
            # Example: sanitized_output = buffered_content.replace('\x0B', ' ')
            # For now, assume content is fine or handled by specific formatter before writing to buffer
            self.ofile.write(buffered_content)
        
        if self.ofile:
            self.ofile.close()
            self.ofile = None # type: ignore

    def _get_scaled_image_width_for_hinting(
        self, 
        element: ImageElement, 
        original_slide_width_px: float,
        target_slide_width_px: float
    ) -> Optional[int]:
        ppt_display_width = element.display_width_px

        if ppt_display_width is None and self.config.image_width is not None: # Default width from config
            ppt_display_width = self.config.image_width
        
        if ppt_display_width is not None and original_slide_width_px > 0:
            width_scale_factor = target_slide_width_px / original_slide_width_px
            return int(round(ppt_display_width * width_scale_factor))
        
        return None

    def _get_image_effective_position_hint(
        self, 
        element: ImageElement, 
        original_slide_width_px: float,
        target_slide_width_px: float
    ) -> Optional[str]:
        scaled_display_width = self._get_scaled_image_width_for_hinting(
            element, original_slide_width_px, target_slide_width_px
        )
        
        calculated_hint = None
        
        if element.left_px is not None and original_slide_width_px > 0 and \
           scaled_display_width is not None and scaled_display_width > 0:
            
            # Scale left_px to the target coordinate system
            if original_slide_width_px > 0 :
                 scaled_left_px = int(round(element.left_px * (target_slide_width_px / original_slide_width_px)))
            else: # Should not happen, but as a fallback
                 scaled_left_px = element.left_px


            image_center_x = scaled_left_px + (scaled_display_width / 2)
            slide_width_for_hinting = target_slide_width_px # Use target width for boundaries
            
            left_third_boundary = slide_width_for_hinting / 3
            right_third_boundary = 2 * slide_width_for_hinting / 3

            # Allow a small tolerance for centering, e.g., 5-10% of slide width around the true center.
            # center_tolerance = slide_width_for_hinting * 0.05 
            # slide_center = slide_width_for_hinting / 2

            # if abs(image_center_x - slide_center) < center_tolerance:
            #     calculated_hint = "center"
            if left_third_boundary < image_center_x < right_third_boundary:
                calculated_hint = "center" 
            elif image_center_x < left_third_boundary: 
                calculated_hint = "left"
            elif image_center_x > right_third_boundary: 
                calculated_hint = "right"
        
        # Prioritize explicit hint if available and valid
        explicit_hint = getattr(element, 'position_hint', None)
        if explicit_hint in ["left", "right", "center"]:
            return explicit_hint
        
        return calculated_hint

    def _separate_slide_elements(
        self,
        initial_elements: List[SlideElement],
        original_slide_width_px: float,
        target_slide_width_px: float
    ) -> Tuple[Optional[SlideElement], List[ImageElement], List[SlideElement]]:
        main_title_element: Optional[SlideElement] = None
        floated_image_elements: List[ImageElement] = []
        other_content_elements: List[SlideElement] = []
        
        temp_content_pool = list(initial_elements) # Work on a copy

        if temp_content_pool and temp_content_pool[0].type == ElementType.Title:
            main_title_element = temp_content_pool.pop(0)

        for element in temp_content_pool:
            if isinstance(element, ImageElement):
                hint = self._get_image_effective_position_hint(
                    element, 
                    original_slide_width_px,
                    target_slide_width_px
                )
                if hint in ["left", "right"]:
                    floated_image_elements.append(element)
                else:
                    other_content_elements.append(element)
            else:
                other_content_elements.append(element)
                
        return main_title_element, floated_image_elements, other_content_elements