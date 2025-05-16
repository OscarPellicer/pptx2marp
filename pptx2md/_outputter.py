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

class Formatter:

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
                     output_segments.append(self._format_single_merged_run(current_merged_text, current_style))
                
                # Start a new merged run
                current_merged_text = normalized_next_text
                current_style = next_run.style
            
        # Format the last accumulated run
        if current_merged_text or (current_style and (current_style.is_code or current_style.is_math)):
            output_segments.append(self._format_single_merged_run(current_merged_text, current_style))
            
        final_text = "".join(output_segments)
        # Removed .strip() from here to allow precise whitespace control by formatters if needed at paragraph level.
        # Individual formatters can strip if they desire.
        # However, for get_formatted_runs which is about content *within* an element,
        # stripping the final combined string is often desirable to avoid leading/trailing spaces
        # from the combination process itself, unless specifically intended.
        # Let's keep strip() for now, as it's generally safer for run concatenation.
        return final_text.strip()

    def put_para(self, text):
        pass

    def put_image(self, path, max_width):
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


class MarkdownFormatter(Formatter):
    # write outputs to markdown
    def __init__(self, config: ConversionConfig):
        super().__init__(config)
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)#\+-\.])')  
        self.esc_re2 = re.compile(r'(<[^>]+>)')

    def put_title(self, text, level):
        self.write('#' * level + ' ' + text + '\n\n')

    def put_list(self, text, level):
        self.write('  ' * level + '* ' + text.strip() + '\n')

    def put_para(self, text):
        self.write(text + '\n\n')

    def put_image(self, path, max_width=None):
        if max_width is None:
            self.write(f'![]({urllib.parse.quote(path)})\n\n')
        else:
            self.write(f'<img src="{path}" style="max-width:{max_width}px;" />\n\n')

    def put_table(self, table: List[List[str]]):
        if not table or not table[0]: return
        gen_table_row = lambda row: '| ' + ' | '.join([
            c.replace('\n', '<br />') if '`' not in c else c.replace('\n', ' ')
            for c in row
        ]) + ' |'
        self.write(gen_table_row(table[0]) + '\n')
        self.write(gen_table_row([':-:' for _ in table[0]]) + '\n') # Centered for Markdown
        self.write('\n'.join([gen_table_row(row) for row in table[1:]]) + '\n\n')

    def put_code_block(self, code: str, language: Optional[str]):
        lang_tag = language if language else ""
        self.write(f'```{lang_tag}\n{code.strip()}\n```\n\n')

    def get_accent(self, text):
        return self._format_text_with_delimiters(text, '_', '_')

    def get_strong(self, text):
        return self._format_text_with_delimiters(text, '__', '__')

    def get_colored(self, text, rgb):
        return ' <span style="color:%s">%s</span> ' % (rgb_to_hex(rgb), text)

    def get_hyperlink(self, text, url):
        return '[' + text + '](' + url + ')'

    def esc_repl(self, match):
        return '\\' + match.group(0)

    def get_escaped(self, text):
        text = re.sub(self.esc_re1, self.esc_repl, text)
        text = re.sub(self.esc_re2, self.esc_repl, text)
        return text


class WikiFormatter(Formatter):
    # write outputs to wikitext
    def __init__(self, config: ConversionConfig):
        super().__init__(config)
        self.esc_re = re.compile(r'<([^>]+)>')

    def put_title(self, text, level):
        self.write('!' * level + ' ' + text + '\n\n')

    def put_list(self, text, level):
        self.write('*' * (level + 1) + ' ' + text.strip() + '\n')

    def put_para(self, text):
        self.write(text + '\n\n')

    def put_image(self, path, max_width):
        if max_width is None:
            self.write(f'<img src="{path}" />\n\n')
        else:
            self.write(f'<img src="{path}" width={max_width}px />\n\n')

    def put_code_block(self, code: str, language: Optional[str]):
        lang_tag = language if language else ""
        self.write(f'```{lang_tag}\n{code.strip()}\n```\n\n')

    def get_accent(self, text):
        return self._format_text_with_delimiters(text, "__", "__") # As it was before for italic emphasis

    def get_strong(self, text):
        return self._format_text_with_delimiters(text, "''", "''") # As it was before for strong emphasis

    def get_colored(self, text, rgb):
        return ' @@color:%s; %s @@ ' % (rgb_to_hex(rgb), text)

    def get_hyperlink(self, text, url):
        return '[[' + text + '|' + url + ']]'

    def esc_repl(self, match):
        return "''''" + match.group(0)

    def get_escaped(self, text):
        text = re.sub(self.esc_re, self.esc_repl, text)
        return text


class MadokoFormatter(Formatter):
    # write outputs to madoko markdown
    def __init__(self, config: ConversionConfig):
        super().__init__(config)
        self.write('[TOC]\n\n') # Use self.write for TOC
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)#\+-\.])')
        self.esc_re2 = re.compile(r'(<[^>]+>)')

    def put_title(self, text, level):
        self.write('#' * level + ' ' + text + '\n\n')

    def put_list(self, text, level):
        self.write('  ' * level + '* ' + text.strip() + '\n')

    def put_para(self, text):
        self.write(text + '\n\n')

    def put_image(self, path, max_width):
        if max_width is None:
            self.write(f'<img src="{path}" />\n\n')
        elif max_width < 500:
            self.write(f'<img src="{path}" width={max_width}px />\n\n')
        else:
            self.write('~ Figure {caption: image caption}\n')
            self.write('![](%s){width:%spx;}\n' % (path, max_width))
            self.write('~\n\n')

    def put_code_block(self, code: str, language: Optional[str]):
        lang_tag = language if language else ""
        self.write(f'```{lang_tag}\n{code.strip()}\n```\n\n')

    def get_accent(self, text):
        return self._format_text_with_delimiters(text, '_', '_')

    def get_strong(self, text):
        return self._format_text_with_delimiters(text, '__', '__')

    def get_colored(self, text, rgb):
        return ' <span style="color:%s">%s</span> ' % (rgb_to_hex(rgb), text)

    def get_hyperlink(self, text, url):
        return '[' + text + '](' + url + ')'

    def esc_repl(self, match):
        return '\\' + match.group(0)

    def get_escaped(self, text):
        text = re.sub(self.esc_re1, self.esc_repl, text)
        text = re.sub(self.esc_re2, self.esc_repl, text)
        return text


class QuartoFormatter(Formatter):
    # write outputs to quarto markdown - reveal js
    def __init__(self, config: ConversionConfig):
        super().__init__(config)
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)#\+-\.])')
        self.esc_re2 = re.compile(r'(<[^>]+>)')

    def output(self, presentation_data: ParsedPresentation):
        self.put_header() # Uses self.write -> buffer

        last_title = None # Quarto specific title handling within its output

        def put_elements(elements: List[SlideElement]):
            nonlocal last_title # last_title is for Quarto's specific duplicate check here

            last_element_type: Optional[ElementType] = None # Changed from last_element
            for element in elements:
                if last_element_type and last_element_type == ElementType.ListItem and element.type != ElementType.ListItem:
                    self.put_list_footer() # Uses Quarto's put_list_footer if overridden, else base.

                current_content_str = ""
                if element.type in [ElementType.Title, ElementType.Paragraph, ElementType.ListItem]:
                    if isinstance(element.content, list): current_content_str = self.get_formatted_runs(element.content)
                    elif isinstance(element.content, str): current_content_str = self.get_escaped(element.content)


                match element.type:
                    case ElementType.Title:
                        title_text = current_content_str.strip() # Use already processed content
                        if title_text:
                            is_similar_to_last = False
                            # Quarto's fuzz.ratio logic for titles
                            if last_title and last_title.level == element.level and fuzz.ratio(
                                    last_title.content, title_text, score_cutoff=92): # title_text is now formatted
                                is_similar_to_last = True
                            
                            if is_similar_to_last:
                                if self.config.keep_similar_titles:
                                    self.put_title(f'{title_text} (cont.)', element.level) 
                            else:
                                self.put_title(title_text, element.level)
                            # Update last_title for Quarto's specific tracking.
                            # Need to decide if last_title stores raw or formatted string.
                            # For fuzz.ratio, it might be better to compare raw content if possible,
                            # or ensure comparison is consistent.
                            # For simplicity here, we'll assume element.content holds the string for comparison.
                            temp_last_title_obj = type('TempTitle', (), {'content': title_text, 'level': element.level})()
                            last_title = temp_last_title_obj

                    case ElementType.ListItem:
                        if not (last_element_type and last_element_type == ElementType.ListItem):
                            self.put_list_header() # Uses Quarto's or base
                        self.put_list(current_content_str, element.level)
                    case ElementType.Paragraph:
                        self.put_para(current_content_str)
                    case ElementType.Image:
                        self.put_image(element.path, element.width) # Assuming element.width is max_width
                    case ElementType.Table:
                        self.put_table([[self.get_formatted_runs(cell) for cell in row] for row in element.content])
                    case ElementType.CodeBlock:
                        code_content = getattr(element, 'content', '')
                        code_lang = getattr(element, 'language', None)
                        self.put_code_block(code_content, code_lang)
                    case ElementType.Formula:
                        if isinstance(element, FormulaElement):
                            self.put_formula(element)
                last_element_type = element.type
            
            if last_element_type == ElementType.ListItem:
                self.put_list_footer()


        for slide_idx, slide in enumerate(presentation_data.slides):
            if slide.type == SlideType.General:
                put_elements(slide.elements)
            elif slide.type == SlideType.MultiColumn:
                put_elements(slide.preface)
                if len(slide.columns) == 2:
                    width = '50%'
                elif len(slide.columns) == 3:
                    width = '33%'
                else: # Should ideally not happen if parser validates
                    width = f'{100/len(slide.columns):.0f}%' if slide.columns else '100%'


                self.put_para(':::: {.columns}') # Uses buffered write
                for column_elements in slide.columns: # Iterate over List[SlideElement] which is a column
                    self.put_para(f'::: {{.column width="{width}"}}') # Uses buffered write
                    put_elements(column_elements) # Process elements within this column
                    self.put_para(':::') # Uses buffered write
                self.put_para('::::') # Uses buffered write

            if not self.config.disable_notes and slide.notes:
                self.put_para("::: {.notes}") # Uses buffered write
                for note in slide.notes:
                    self.put_para(note) # Assumes note is already a string, put_para will handle escaping via get_formatted_runs
                self.put_para(":::") # Uses buffered write

            if slide_idx < len(presentation_data.slides) - 1 and self.config.enable_slides:
                self.put_para("\n---\n") # Uses buffered write

        self.close() # Calls base Formatter.close()

    def put_header(self):
        self.write('''---
title: "Presentation Title"
author: "Author"
format: 
  revealjs:
    slide-number: c/t
    width: 1600
    height: 900
    logo: img/logo.png
    footer: "Organization"
    incremental: true
    theme: [simple]
---
''') # Uses buffered write

    def put_title(self, text, level):
        self.write('#' * level + ' ' + text + '\n\n') # Uses buffered write

    def put_list(self, text, level):
        self.write('  ' * level + '* ' + text.strip() + '\n') # Uses buffered write

    def put_para(self, text):
        self.write(text + '\n\n') # Uses buffered write

    def put_image(self, path, max_width=None): # Signature matches MarkdownFormatter
        # path here is element.path, max_width is element.width
        # This is slightly different from other put_image that take the full element.
        # For consistency, it might be better to make all put_image take ImageElement.
        # For now, adapting to existing signature.
        quoted_path = urllib.parse.quote(str(path)) # Ensure path is string
        if max_width is None:
            self.write(f'![]({quoted_path})\n\n') # Uses buffered write
        else:
            self.write(f'<img src="{quoted_path}" style="max-width:{max_width}px;" />\n\n') # Uses buffered write

    def put_table(self, table):
        gen_table_row = lambda row: '| ' + ' | '.join([c.replace('\n', '<br />') for c in row]) + ' |'
        self.write(gen_table_row(table[0]) + '\n') # Uses buffered write
        self.write(gen_table_row([':-:' for _ in table[0]]) + '\n') # Uses buffered write
        self.write('\n'.join([gen_table_row(row) for row in table[1:]]) + '\n\n') # Uses buffered write

    def put_code_block(self, code: str, language: Optional[str]):
        lang_tag = language if language else ""
        self.write(f'```{lang_tag}\n{code.strip()}\n```\n\n') # Uses buffered write

    def put_formula(self, element: FormulaElement): # Added for Quarto
        # Quarto uses $...$ for inline and $$...$$ for display math, similar to Markdown.
        # The _format_text_with_delimiters in base class handles $$...$$ for block formulas.
        # FormulaElement content should be raw math.
        formatted_content = self._format_text_with_delimiters(element.content, "$$", "$$")
        self.write(f'{formatted_content}\n\n') # Uses buffered write

    def get_accent(self, text):
        return self._format_text_with_delimiters(text, '_', '_')

    def get_strong(self, text):
        return self._format_text_with_delimiters(text, '**', '**')

    def get_colored(self, text, rgb):
        return ' <span style="color:%s">%s</span> ' % (rgb_to_hex(rgb), text)

    def get_hyperlink(self, text, url):
        return '[' + text + '](' + url + ')'

    def esc_repl(self, match):
        return '\\' + match.group(0)

    def get_escaped(self, text):
        text = re.sub(self.esc_re1, self.esc_repl, text)
        text = re.sub(self.esc_re2, self.esc_repl, text)
        return text

class MarpFormatter(Formatter):
    # write outputs to marp markdown
    def __init__(self, config: ConversionConfig):
        super().__init__(config) # This now sets up self._buffer and self.last_title_info
        self.esc_re1 = re.compile(r'([\|\*`])')
        self.esc_re2 = re.compile(r'(<[^>]+>)')
        # self.last_title_info: Optional[Tuple[str, int]] = None # Moved to base

    def write(self, text: str): # Override to write directly to file
        self.ofile.write(text)

    def put_header(self):
        # CSS content, now to be embedded
        css_content = """
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
  /* Default width/height can be auto or set via style if needed by content */
  /* Ensure z-index is used if overlap control is needed */
}

img[alt~="center"] {
  display: block;
  margin: 0 auto;
}
img[alt~="left"] {
  float: left;
  margin-right: 1em;
  margin-bottom: 0.5em; /* Optional: consistent with previous .img-float-left */
}
img[alt~="right"] {
  float: right;
  margin-left: 1em;
  margin-bottom: 0.5em; /* Optional: consistent with previous .img-float-right */
}
/* For Marp background images: ![bg right:30% 200%](image.jpg) */
/* For Marp image sizing: ![alt text w:300px](image.png) */

.columns {
  display: grid;
  grid-template-columns: repeat(2, 1fr); /* Creates two equal-width columns */
  gap: 2em; /* Adjust the gap between columns as needed */
}

.columns > div {
  /* Optional: You can add styling for individual columns here if needed */
  /* For example, to ensure lists render correctly within columns */
  overflow: hidden; /* Helps with list rendering inside flex/grid items */
}

/* Styles for images with captions (figure container) */
.figure-container {
  margin-bottom: 1em; /* Space below the figure block */
  /* Consider clear: both; if flow issues arise after floated figures,
     though Marp sections usually handle this. */
}

.figure-container img {
  display: block; /* Image as a block element within its container */
  max-width: 100%; /* Responsive: won't overflow container */
  height: auto;
  /* Center image if container is wider than image (e.g., for align-center) */
  margin-left: auto;
  margin-right: auto;
}

.figure-container.align-left {
  float: left;
  margin-right: 1em; /* Spacing from content to its right */
  margin-left: 0; /* Override auto margins for float */
}
.figure-container.align-left img {
  margin-left: 0; /* Align image to the left of its container */
  margin-right: auto; /* Allow centering if container is somehow wider */
}

.figure-container.align-right {
  float: right;
  margin-left: 1em; /* Spacing from content to its left */
  margin-right: 0; /* Override auto margins for float */
}
.figure-container.align-right img {
  margin-right: 0; /* Align image to the right of its container */
  margin-left: auto; /* Allow centering if container is somehow wider */
}

.figure-container.align-center {
  display: block; /* Container is block, centered by its own margins */
  margin-left: auto;
  margin-right: auto;
  /* The img inside will be centered due to its own auto margins */
}

.figure-container .figcaption,
.figure-container > em { /* Supports <p class="figcaption"> or direct <em> */
  display: block; /* Ensures caption is block for consistent styling */
  font-size: 0.85em;
  color: #555; /* Muted color for caption text */
  text-align: center; /* Captions are centered by default */
  margin-top: 0.4em; /* Space between image and caption */
  line-height: 1.3;
  font-style: normal; /* Override em's italic if p.figcaption is used; em will be italic */
}
.figure-container > em {
  font-style: italic; /* Ensure em tag remains italic */
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
        # self.write() used above calls MarpFormatter's direct file write.

    def _put_elements_on_slide(self, elements: List[SlideElement], is_continued_slide: bool = False):
        """Helper to output a list of elements. `self.last_title_info` is from base class."""
        last_element_type: Optional[ElementType] = None
        for element_idx, element in enumerate(elements):
            if last_element_type and last_element_type == ElementType.ListItem and element.type != ElementType.ListItem:
                self.put_list_footer()

            # Special handling for the first title on a continued slide part
            if is_continued_slide and element_idx == 0 and element.type == ElementType.Title:
                # The "(Continued)" title is already printed by the caller, so skip this one if it's the original.
                # This assumes the first element of the second half of a split is the original title.
                # A more robust way would be to pass the original slide's main title text and level.
                # For now, if it's a continued slide, we assume the title is handled by the main output loop for now.
                # However, if this element IS the explicit "(Continued)" title, it should be printed.
                # This logic needs to be careful. Let's assume caller prints the continued title.
                pass # Title for continued slide is handled by the main output loop for now.
            
            current_content_str = ""
            if element.type in [ElementType.Title, ElementType.Paragraph, ElementType.ListItem]:
                if isinstance(element.content, list): current_content_str = self.get_formatted_runs(element.content)
                elif isinstance(element.content, str): current_content_str = self.get_escaped(element.content)


            match element.type:
                case ElementType.Title:
                    title_text = element.content.strip() if isinstance(element.content, str) else ""
                    if title_text:
                        # If this is the start of a continued part of a slide, the (Continued) title is manually added by output()
                        # So, we only process regular titles here.
                        if not (is_continued_slide and element_idx == 0): # Avoid re-printing main title on continued part
                            is_similar_to_last = False
                            if self.last_title_info and self.last_title_info[1] == element.level and \
                               fuzz.ratio(self.last_title_info[0], title_text, score_cutoff=92):
                                is_similar_to_last = True
                            
                            if is_similar_to_last:
                                if self.config.keep_similar_titles:
                                    effective_title = f'{title_text}'# (cont.)
                                    self.put_title(effective_title, element.level)
                                    self.last_title_info = (effective_title, element.level)
                                # else skip
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
                    # Pass the whole element to put_image
                    self.put_image(element)
                case ElementType.Table:
                    self.put_table([[self.get_formatted_runs(cell) for cell in row] for row in element.content])
                case ElementType.CodeBlock:
                    code_content = getattr(element, 'content', '')
                    code_lang = getattr(element, 'language', None)
                    self.put_code_block(code_content, code_lang)
                case ElementType.Formula:
                    if isinstance(element, FormulaElement):
                        self.put_formula(element)
            
            last_element_type = element.type
        
        if last_element_type == ElementType.ListItem: # Ensure list footer if slide ends with list
            self.put_list_footer()

    def output(self, presentation_data: ParsedPresentation):
        self.put_header() # Writes directly to file
        self.last_title_info = None # Reset for each presentation

        num_total_slides = len(presentation_data.slides)
        marp_slide_counter = 0

        for slide_idx, slide in enumerate(presentation_data.slides):
            marp_slide_counter += 1

            all_elements: List[SlideElement] = []
            if slide.type == SlideType.General:
                all_elements = slide.elements
            elif slide.type == SlideType.MultiColumn:
                # For Marp, flatten MultiColumn for now, title separation will handle preface.
                all_elements = slide.preface + [el for col in slide.columns for el in col] 

            if not all_elements: 
                 if marp_slide_counter < num_total_slides : 
                    self.write("\n---\n\n") # Writes directly to file
                 continue

            # USE THE BASE CLASS METHODS
            line_count, char_count, max_img_w, max_img_h, text_lines_for_avg, text_chars_for_avg = \
                self._get_slide_content_metrics(all_elements)
            current_slide_class = self._get_slide_density_class(line_count)

            # Determine if slide qualifies for splitting based on its overall content metrics
            initial_split_qualification = False
            if current_slide_class in ["smaller", "smallest"]:
                if text_lines_for_avg > 0:
                    avg_line_length = text_chars_for_avg / text_lines_for_avg
                    if avg_line_length < 40:
                        initial_split_qualification = True
            
            # Identify title and content that would go into columns
            main_title_element: Optional[SlideElement] = None
            content_for_columns: List[SlideElement] = all_elements

            if all_elements and all_elements[0].type == ElementType.Title:
                main_title_element = all_elements[0]
                content_for_columns = all_elements[1:] # Elements after the title

            # Final decision: must qualify AND have enough elements left for columns
            # AND not contain a table in the content intended for columns.
            contains_table_in_content_for_columns = False
            if initial_split_qualification: # Only check for tables if it might split
                for element in content_for_columns:
                    if element.type == ElementType.Table:
                        contains_table_in_content_for_columns = True
                        break
            
            actually_split_columns = initial_split_qualification and \
                                     len(content_for_columns) >= 2 and \
                                     not contains_table_in_content_for_columns

            # Determine effective slide class
            effective_slide_class = current_slide_class
            if actually_split_columns:
                # If it was going to be 'smaller' or 'smallest' and we are splitting,
                # make it 'small' as content is now distributed.
                if current_slide_class in ["smaller", "smallest"]:
                    effective_slide_class = "small"
            
            # Output class directive (if any)
            if effective_slide_class:
                self.write(f"<!-- _class: {effective_slide_class} -->\n\n") # Writes directly

            # Output the main title (if it was identified and separated)
            if main_title_element:
                self._put_elements_on_slide([main_title_element], is_continued_slide=False)

            # Output the remaining content, either in columns or as a single block
            if actually_split_columns:
                # Split content_for_columns and output in two divs
                num_in_first_col = (len(content_for_columns) + 1) // 2
                first_half_elements = content_for_columns[:num_in_first_col]
                second_half_elements = content_for_columns[num_in_first_col:]

                self.write('<div class="columns">\n<div>\n\n') # Writes directly
                self._put_elements_on_slide(first_half_elements, is_continued_slide=False)
                self.write('\n</div>\n<div>\n\n') # Writes directly
                self._put_elements_on_slide(second_half_elements, is_continued_slide=False)
                self.write('\n</div>\n</div>\n\n') # Writes directly
            else:
                # Not splitting columns (either didn't qualify or not enough content after title).
                # Output content_for_columns as a single block.
                # This list contains all elements if no title was found at the start,
                # or elements after the title if a title was found and already printed.
                if content_for_columns: # Only print if there's content remaining
                    self._put_elements_on_slide(content_for_columns, is_continued_slide=False)
            
            if not self.config.disable_notes and slide.notes:
                self.write("<!--\n") # Writes directly
                for note_line in slide.notes:
                    self.write(f"{note_line}\n") # Writes directly
                self.write("-->\n\n") # Writes directly

            # Add slide separator if not the very last conceptual slide
            is_last_original_slide = (slide_idx == num_total_slides - 1)
            if not (is_last_original_slide) : # Add --- if not the true end
                 self.write("\n---\n\n") # Writes directly

        self.close() # MarpFormatter should have its own close or rely on base if it were using buffer

    def close(self): # MarpFormatter specific close, as it writes directly
        if self.ofile:
            self.ofile.close()
            self.ofile = None # type: ignore

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
        
        has_bg_keyword = False
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
        text = re.sub(self.esc_re1, self.esc_repl, text)
        text = re.sub(self.esc_re2, self.esc_repl, text)
        return text

class BeamerFormatter(Formatter):
    # write outputs to LaTeX Beamer
    def __init__(self, config: ConversionConfig):
        super().__init__(config) # This now sets up self._buffer and self.last_title_info
        # self._buffer is used by BeamerFormatter's self.write() (inherited from Formatter)
        # LaTeX specific escaping. Order can be important.
        self.esc_map = {
            '\\': r'\textbackslash{}',
            '{': r'\{',
            '}': r'\}',
            '&': r'\&',
            '%': r'\%',
            '$': r'\$',
            '#': r'\#',
            '_': r'\_',
            '^': r'\textasciicircum{}',
            '~': r'\textasciitilde{}',
            '<': r'\textless{}',
            '>': r'\textgreater{}',
            '|': r'\textbar{}',
            '"': r"''", # Using typographic quotes for `"`
            '\u2019': r"'", # Typographic right single quote (U+2019) to apostrophe
            # Add other common typographic chars if needed, using their Unicode escapes:
            '\u2018': r"`",   # Typographic left single quote (U+2018) to grave accent
            '\u201C': r"``",  # Typographic left double quote (U+201C) to double grave
            '\u201D': r"''",  # Typographic right double quote (U+201D) to double apostrophe
            '\u2013': r"--",  # en-dash (U+2013) to TeX en-dash
            '\u2014': r"---", # em-dash (U+2014) to TeX em-dash
            # Non-breaking space is handled in get_formatted_runs, but for completeness:
            '\u00A0': r"~", # Non-breaking space to TeX tie (~) OR " " if preferred
        }
        self.esc_re = re.compile('|'.join(re.escape(key) for key in self.esc_map.keys()))
        self.in_frame = False
        self.current_list_level = 0 # 0 = no list active, 1 = first level itemize, etc.
        # REMOVED: self._buffer = io.StringIO() # Base class handles this

    def write(self, text: str): # Now uses inherited Formatter.write()
        self._buffer.write(text)

    def _put_elements_on_slide(self, elements: List[SlideElement]):
        """Helper to output a list of elements within the current Beamer frame."""
        last_element_type: Optional[ElementType] = None
        for element in elements:
            if last_element_type and last_element_type == ElementType.ListItem and element.type != ElementType.ListItem:
                self.put_list_footer()

            current_content_str = ""
            if element.type in [ElementType.Title, ElementType.Paragraph, ElementType.ListItem]:
                if isinstance(element.content, list) and all(isinstance(run, TextRun) for run in element.content):
                    current_content_str = self.get_formatted_runs(element.content)
                elif isinstance(element.content, str):
                    # For Beamer, non-run string content should also be escaped.
                    current_content_str = self.get_escaped(element.content)
            
            match element.type:
                case ElementType.Title:
                    # This is for titles *within* a frame, not the \frametitle
                    # current_content_str already contains the formatted/escaped title
                    if current_content_str.strip(): # Ensure there's content
                        self.put_title(current_content_str.strip(), element.level)
                case ElementType.ListItem:
                    if not (last_element_type and last_element_type == ElementType.ListItem):
                        self.put_list_header()
                    self.put_list(current_content_str, element.level) # current_content_str is formatted/escaped
                case ElementType.Paragraph:
                    self.put_para(current_content_str) # current_content_str is formatted/escaped
                case ElementType.Image:
                    if isinstance(element, ImageElement):
                        self.put_image(element) 
                case ElementType.Table:
                    if element.content: # Check if table has content
                        # Table cell content needs to be formatted (runs to string, escaped)
                        table_content_processed = []
                        for row in element.content:
                            processed_row = []
                            for cell_runs_or_str in row:
                                if isinstance(cell_runs_or_str, list): # List[TextRun]
                                    processed_row.append(self.get_formatted_runs(cell_runs_or_str))
                                elif isinstance(cell_runs_or_str, str):
                                    processed_row.append(self.get_escaped(cell_runs_or_str))
                                else:
                                    processed_row.append('') # Should not happen with valid AST
                            table_content_processed.append(processed_row)
                        self.put_table(table_content_processed)
                case ElementType.CodeBlock:
                    # Content is raw code string, language is optional string
                    self.put_code_block(element.content, element.language)
                case ElementType.Formula:
                    if isinstance(element, FormulaElement):
                        self.put_formula(element) # put_formula handles its own escaping/formatting for LaTeX
            
            last_element_type = element.type
        
        if last_element_type == ElementType.ListItem: # Ensure list footer if elements end with list
            self.put_list_footer()

    def put_header(self):
        # Default Beamer page size is 128mm x 96mm (4:3)
        # For 16:9 aspect ratio (like 1280x720), use aspectratio=169
        # Marp default: 1280x720px.
        self.write(r'''\documentclass[aspectratio=169]{beamer}
\usetheme{default} % Or any other theme

\usepackage[utf8]{inputenc}
\usepackage{graphicx} % For images
\usepackage{booktabs} % For tables (toprule, midrule, bottomrule)
\usepackage{xcolor}   % For colors
\usepackage{hyperref} % For hyperlinks
\usepackage{amsmath}  % For math
\usepackage{amssymb}  % For math symbols
\usepackage{wrapfig}  % For text wrapping around figures
\usepackage{listings} % For code blocks (optional, more advanced)
% \usepackage{minted} % For code blocks (optional, powerful, needs shell-escape)

% Beamer settings
\beamertemplatenavigationsymbolsempty % Disable navigation symbols
% \setbeamertemplate{footline}[frame number] % Optionally show frame number

% \title{Presentation Title} % Removed: No automatic title page
% \author{Author Name}     % Removed
% \date{\today}            % Removed

\begin{document}

% \maketitle % Removed: No automatic title page

''') # self.write() now uses inherited buffered write

    def output(self, presentation_data: ParsedPresentation):
        self.put_header()
        self.last_title_info = None 

        for slide_idx, slide in enumerate(presentation_data.slides):
            slide_elements_for_processing: List[SlideElement] = []
            is_multicolumn_slide_type = False
            original_columns_data: Optional[List[List[SlideElement]]] = None

            if slide.type == SlideType.General:
                slide_elements_for_processing = slide.elements
            elif slide.type == SlideType.MultiColumn:
                is_multicolumn_slide_type = True
                slide_elements_for_processing = slide.preface
                original_columns_data = slide.columns

            if not slide_elements_for_processing and not (is_multicolumn_slide_type and original_columns_data):
                if slide_idx < len(presentation_data.slides) - 1:
                    self.write(r'\begin{frame}{}\end{frame}' + '\n\n') 
                continue

            initial_all_text_elements = slide.preface + [el for col in (original_columns_data or []) for el in col] if is_multicolumn_slide_type else slide_elements_for_processing
            line_count, _, _, _, text_lines_for_avg, text_chars_for_avg = \
                self._get_slide_content_metrics(initial_all_text_elements)
            density_class = self._get_slide_density_class(line_count)
            
            self.write(r'\begin{frame}') # Open the frame
            self.in_frame = True
            
            main_title_element: Optional[SlideElement] = None
            content_after_title: List[SlideElement] = [] 

            if slide_elements_for_processing and slide_elements_for_processing[0].type == ElementType.Title:
                main_title_element = slide_elements_for_processing[0]
                content_after_title = slide_elements_for_processing[1:]
                
                title_text_runs = main_title_element.content if isinstance(main_title_element.content, list) else None
                title_text_str = main_title_element.content if isinstance(main_title_element.content, str) else None
                formatted_title = ""
                if title_text_runs: formatted_title = self.get_formatted_runs(title_text_runs)
                elif title_text_str: formatted_title = self.get_escaped(title_text_str.strip())
                
                if formatted_title:
                    # Correctly place \frametitle
                    self.write(f'\n\\frametitle{{{formatted_title}}}\n') 
                    self.last_title_info = (formatted_title, main_title_element.level)
            else:
                # No title, so all elements are content after (non-existent) title
                content_after_title = slide_elements_for_processing
            
            current_font_scale_opened = False
            if density_class == "small": 
                self.write(r'{\small' + "\n")
                current_font_scale_opened = True
            elif density_class == "smaller": 
                self.write(r'{\footnotesize' + "\n")
                current_font_scale_opened = True
            elif density_class == "smallest": 
                self.write(r'{\scriptsize' + "\n")
                current_font_scale_opened = True

            # --- Column Splitting Logic ---
            # This heuristic applies if the slide was NOT originally SlideType.MultiColumn,
            # OR if it was, but we want to re-evaluate the `content_after_title` from preface.
            # For now, let's simplify: if slide.type was MultiColumn, we use its structure.
            # Otherwise, we apply the heuristic to `content_after_title`.

            actually_split_columns_heuristic = False
            if not is_multicolumn_slide_type and content_after_title: # Apply heuristic only to General slides
                # Use metrics from content_after_title for column split decision
                _, _, _, _, ca_text_lines, ca_text_chars = self._get_slide_content_metrics(content_after_title)
                # The density_class for font size is already determined from the whole slide.
                # This is just for column splitting decision.
                current_content_density_class = self._get_slide_density_class(ca_text_lines)

                initial_split_qualification = False
                if current_content_density_class in ["smaller", "smallest"]: # or use line_count from content_after_title
                    if ca_text_lines > 0:
                        avg_line_length = ca_text_chars / ca_text_lines
                        if avg_line_length < 40: # Marp's threshold
                            initial_split_qualification = True
                
                contains_table_in_content = any(el.type == ElementType.Table for el in content_after_title)
                
                actually_split_columns_heuristic = initial_split_qualification and \
                                         len(content_after_title) >= 2 and \
                                         not contains_table_in_content

            # --- Output content ---
            if is_multicolumn_slide_type and original_columns_data:
                # Handle preface elements (if any remain after title extraction)
                if main_title_element and slide_elements_for_processing == slide.preface: # content_after_title is rest of preface
                     self._put_elements_on_slide(content_after_title) 
                elif not main_title_element and slide.preface: # No title in preface, output all preface
                     self._put_elements_on_slide(slide.preface)

                # Output Beamer columns from original_columns_data
                num_cols = len(original_columns_data)
                if num_cols > 0:
                    self.write(r'\begin{columns}[T]' + '\n') # [T] for top alignment
                    col_width = f'{1/num_cols:.2f}' # e.g., 0.50, 0.33
                    for column_data_list in original_columns_data:
                        self.write(r'  \column{' + col_width + r'\textwidth}' + '\n')
                        self._put_elements_on_slide(column_data_list)
                    self.write(r'\end{columns}' + '\n')

            elif actually_split_columns_heuristic:
                # Heuristically split content_after_title into two Beamer columns
                num_in_first_col = (len(content_after_title) + 1) // 2
                first_half_elements = content_after_title[:num_in_first_col]
                second_half_elements = content_after_title[num_in_first_col:]

                self.write(r'\begin{columns}[T]' + '\n')
                self.write(r'  \column{0.48\textwidth}' + '\n') # Slight gap with 0.48+0.48
                self._put_elements_on_slide(first_half_elements)
                self.write(r'  \column{0.48\textwidth}' + '\n')
                self._put_elements_on_slide(second_half_elements)
                self.write(r'\end{columns}' + '\n')
            else:
                # Output content_after_title as a single block (no columns)
                self._put_elements_on_slide(content_after_title)
            
            # Notes and end of frame
            if not self.config.disable_notes and slide.notes:
                # Escape notes content
                escaped_notes = [self.get_escaped(note) for note in slide.notes]
                self.write(r'\note{' + '\n'.join(escaped_notes) + '}\n')

            if current_font_scale_opened: 
                self.write("\n}\n") # Close font scaling group
            
            self.write(r'\end{frame}' + '\n\n')
            self.in_frame = False

        self.write(r'\end{document}' + '\n')
        self.close()

    def put_title(self, text: str, level: int):
        # For titles within a frame (e.g., if not the frametitle)
        # text is already escaped and formatted by get_formatted_runs
        if level == 1: # Beamer's \section* ? or \block?
            self.write(r'\begin{block}{' + text + '}\n\\end{block}\n\n') 
        elif level == 2: # \subsection*?
            self.write(r'\textbf{' + text + '}\par\n\n') # Ensure paragraph break after
        else:
            self.write(r'\textit{' + text + '}\par\n\n') # Ensure paragraph break after


    def put_list(self, text: str, level: int):
        # text is already escaped and formatted by get_formatted_runs
        # level is 0-indexed from the parser
        target_latex_nest_level = level + 1 # 1-indexed for LaTeX environment nesting

        # Open new environments if nesting deeper
        while self.current_list_level < target_latex_nest_level:
            indent_str = '  ' * self.current_list_level
            self.write(indent_str + r'\begin{itemize}' + '\n')
            self.current_list_level += 1
        
        # Close environments if un-nesting (going to a shallower level)
        while self.current_list_level > target_latex_nest_level:
            self.current_list_level -= 1
            indent_str = '  ' * self.current_list_level
            self.write(indent_str + r'\end{itemize}' + '\n')

        # Write the item at the current (now target) nest level
        # Indentation for the \item itself is based on the original 0-indexed level
        item_indent_str = '  ' * level 
        self.write(item_indent_str + r'\item ' + text.strip() + '\n')

    def put_list_header(self):
        # This is called by the base Formatter's output loop when a list sequence starts.
        # With the new put_list logic, this can be a no-op, as put_list will handle
        # opening the necessary environments based on self.current_list_level.
        pass

    def put_list_footer(self):
        # This is called by the base Formatter's output loop when a list sequence ends.
        # Close any remaining open itemize environments.
        while self.current_list_level > 0:
            self.current_list_level -= 1
            indent_str = '  ' * self.current_list_level
            self.write(indent_str + r'\end{itemize}' + '\n')
        # Ensure current_list_level is reset to 0, indicating no list is active.
        self.current_list_level = 0


    def put_para(self, text: str):
        # text is already escaped and formatted by get_formatted_runs
        self.write(text + '\n\n')

    def put_image(self, element: ImageElement):
        image_path_latex = element.path 

        # Do not include captions automatically, let users handle it
        caption_text = None
        # caption_text = self.get_escaped(element.alt_text) if element.alt_text else None

        position_hint = "center" 
        wrapfig_char_placement = None 

        if element.left_px is not None and element.display_width_px is not None and \
           self.config.slide_width_px and self.config.slide_width_px > 0:
            ppt_slide_w = self.config.slide_width_px
            image_center_ppt = element.left_px + (element.display_width_px / 2)
            if image_center_ppt < ppt_slide_w / 3.0:
                position_hint = "left"
                wrapfig_char_placement = "l" 
            elif image_center_ppt > ppt_slide_w * (1 - 1/3.0):
                position_hint = "right"
                wrapfig_char_placement = "r" 
        
        effective_position_hint = getattr(element, 'position_hint', position_hint)
        if effective_position_hint == "left":
            wrapfig_char_placement = "l"
        elif effective_position_hint == "right":
            wrapfig_char_placement = "r"
        
        # Default wrapfigure width fraction if calculation is not possible or yields too small/large values
        # This is the fraction of \linewidth that the wrapfigure environment will occupy.
        wf_width_frac = 0.4 # Default width for wrapfigure: 40% of linewidth

        if element.display_width_px and self.config.slide_width_px and self.config.slide_width_px > 0:
            ppt_img_frac_of_slide = element.display_width_px / self.config.slide_width_px
            # Let's make the wrapfigure occupy a space proportional to the image's width on the slide,
            # but cap it to prevent it from being too dominant or too small.
            # e.g., if image was 60% of PPT slide, maybe wrapfig takes 50-60% of textwidth.
            # If image was 10% of PPT slide, maybe wrapfig takes 20-25% of textwidth.
            wf_width_frac = min(max(0.25, ppt_img_frac_of_slide), 0.6) # Cap between 25% and 60%

        # Options for \includegraphics
        # Inside wrapfigure, we want the image to fill the wrapfigure's width while keeping aspect ratio.
        # For centered figures, we can use a similar logic for its width.
        includegraphics_opts_str = ""
        if wrapfig_char_placement: # For wrapfigure
            includegraphics_opts_str = "width=\\linewidth,keepaspectratio"
        else: # For centered figure
            # Calculate width for centered image (similar to wf_width_frac but can be larger)
            center_img_width_frac = 0.7 # Default for centered image
            if element.display_width_px and self.config.slide_width_px and self.config.slide_width_px > 0:
                 ppt_img_frac_of_slide = element.display_width_px / self.config.slide_width_px
                 center_img_width_frac = min(max(0.2, ppt_img_frac_of_slide), 0.85) # Cap 20%-85% for centered
            includegraphics_opts_str = f"width={center_img_width_frac:.2f}\\textwidth,keepaspectratio"


        if wrapfig_char_placement and (effective_position_hint == "left" or effective_position_hint == "right"):
            self.write(f'\\begin{{wrapfigure}}{{{wrapfig_char_placement}}}{{{wf_width_frac:.2f}\\linewidth}}\n')
            self.write(r'  \centering' + '\n') 
            self.write(f'  \\includegraphics[{includegraphics_opts_str}]{{{image_path_latex}}}\n')
            if caption_text: # Though you requested no captions, keeping the if block
                self.write(f'  \\caption{{{caption_text}}}\n')
            self.write(r'\end{wrapfigure}' + '\n')
        else:
            self.write(r'\begin{figure}' + '\n')
            self.write(r'  \centering' + '\n')
            self.write(f'  \\includegraphics[{includegraphics_opts_str}]{{{image_path_latex}}}\n')
            if caption_text:
                self.write(f'  \\caption{{{caption_text}}}\n')
            self.write(r'\end{figure}' + '\n\n')

    def put_code_block(self, code: str, language: Optional[str]):
        lang_opt = f'[language={language}]' if language and self.config.use_listings else ''
        self.write(f'\\begin{{lstlisting}}{lang_opt}\n{self.get_escaped(code.strip(), verbatim_like=True)}\n\\end{{lstlisting}}\n\n')

    def put_formula(self, element: FormulaElement):
        content = element.content.strip()
        # Check if it's already delimited for display math, or should be
        if content.startswith('$$') and content.endswith('$$'):
            math_content = content[2:-2].strip()
            self.write(f'\\[\n{math_content}\n\\]\n\n')
        elif content.startswith('$') and content.endswith('$') and not content.startswith('$$'):
             # Inline math, already formatted by get_formatted_runs potentially
            self.write(f'{content}\n\n') # If it reached here as a block
        else:
            # Assume it's a block of math to be displayed
            self.write(f'\\[\n{content}\n\\]\n\n')


    def get_inline_code(self, text: str) -> str:
        # For inline code, \texttt{} or a custom command with listings/minted
        return r'\texttt{' + self.get_escaped(text, verbatim_like=True) + '}'

    def get_accent(self, text):
        return self._format_text_with_delimiters(text, r'\textit{', '}')

    def get_strong(self, text):
        return self._format_text_with_delimiters(text, r'\textbf{', '}')

    def get_colored(self, text, rgb):
        # Convert RGB to 0-1 scale for xcolor if needed, or use rbg model
        # xcolor \definecolor{mycolor}{RGB}{r,g,b} then \textcolor{mycolor}{text}
        # Or directly: \textcolor[RGB]{r,g,b}{text}
        r, g, b = rgb
        # Text itself should be properly escaped if it contains special LaTeX characters
        # The _format_single_merged_run method already handles escaping non-code/non-math text
        # before calling get_colored. So 'text' here is assumed to be "final" form for its content.
        return f'\\textcolor[RGB]{{{r},{g},{b}}}{{{text}}}'

    def get_hyperlink(self, text, url):
        # text is the display text, url is the target
        # text is assumed to be already formatted/escaped by get_formatted_runs
        escaped_url = self.get_escaped(url, is_url=True)
        # The text part for \href typically doesn't need further escaping IF it's already processed.
        # If text contains e.g. an already escaped underscore like \_ this is fine.
        return r'\href{' + escaped_url + '}{' + text + '}'

    def esc_repl(self, match, verbatim_like=False, is_url=False):
        char = match.group(0)
        if verbatim_like: # Inside \texttt{} or verbatim, less escaping is needed/different rules
            if char == '{': return r'\{'
            if char == '}': return r'\}'
            if char == '\\': return r'\textbackslash{}'
            # Other chars are usually fine in \texttt
            return char
        if is_url:
            # URLs have specific characters that are problematic for TeX: %, #, &, _, ~
            # hyperref usually handles many, but explicit escaping can be safer.
            if char == '%': return r'\%'
            if char == '#': return r'\#'
            if char == '&': return r'\&'
            if char == '_': return r'\_'
            if char == '~': return r'\textasciitilde{}' # or let hyperref handle
            # Other chars in URLs are typically fine.
            return char
        return self.esc_map.get(char, char)

    def get_escaped(self, text, verbatim_like=False, is_url=False):
        if self.config.disable_escaping:
            return text
        # When calling re.sub with a function, the function gets a match object
        return self.esc_re.sub(lambda m: self.esc_repl(m, verbatim_like, is_url), text)

    def put_list_header(self):
        # This is called by the base Formatter's output loop when a list sequence starts.
        # With the new put_list logic, this can be a no-op, as put_list will handle
        # opening the necessary environments based on self.current_list_level.
        pass

    def put_list_footer(self):
        # This is called by the base Formatter's output loop when a list sequence ends.
        # Close any remaining open itemize environments.
        while self.current_list_level > 0:
            self.current_list_level -= 1
            indent_str = '  ' * self.current_list_level
            self.write(indent_str + r'\end{itemize}' + '\n')
        # Ensure current_list_level is reset to 0, indicating no list is active.
        self.current_list_level = 0
