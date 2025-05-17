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
# import urllib.parse # Not obviously used directly
from typing import List, Optional, Tuple # Union not obviously used directly
# import io # Not obviously used directly

# from rapidfuzz import fuzz # Not obviously used directly in BeamerFormatter specific logic

from pptx2md.outputter.base import Formatter, DEFAULT_SLIDE_WIDTH_PX # Import base items
from pptx2md.types import ParsedPresentation, SlideElement, ElementType, SlideType, TextRun, ImageElement, FormulaElement # TextStyle not obviously used directly
# from pptx2md.utils import rgb_to_hex # Not directly used, get_colored is overridden

class BeamerFormatter(Formatter):
    # write outputs to LaTeX Beamer
    def __init__(self, config):
        super().__init__(config)
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
            '"': r"''", 
            '\u2019': r"'",
            '\u2018': r"`",
            '\u201C': r"``",
            '\u201D': r"''",
            '\u2013': r"--",
            '\u2014': r"---",
            '\u00A0': r"~",
            '\u000B': r' ', # Vertical Tab -> space
            '\u000C': r' ', # Form Feed -> space
        }
        self.esc_re = re.compile('|'.join(re.escape(key) for key in self.esc_map.keys()))
        self.in_frame = False
        self.current_list_level = 0

    # write is inherited from Formatter base, uses self._buffer

    def _put_elements_on_slide(self, elements: List[SlideElement]):
        last_element_type: Optional[ElementType] = None
        for element in elements:
            if last_element_type and last_element_type == ElementType.ListItem and element.type != ElementType.ListItem:
                self.put_list_footer()

            current_content_str = ""
            if element.type in [ElementType.Title, ElementType.Paragraph, ElementType.ListItem]:
                if isinstance(element.content, list) and all(isinstance(run, TextRun) for run in element.content):
                    current_content_str = self.get_formatted_runs(element.content)
                elif isinstance(element.content, str):
                    current_content_str = self.get_escaped(element.content)
            
            match element.type:
                case ElementType.Title:
                    if current_content_str.strip():
                        self.put_title(current_content_str.strip(), element.level)
                case ElementType.ListItem:
                    if not (last_element_type and last_element_type == ElementType.ListItem):
                        self.put_list_header()
                    self.put_list(current_content_str, element.level)
                case ElementType.Paragraph:
                    self.put_para(current_content_str)
                case ElementType.Image:
                    if isinstance(element, ImageElement):
                        self.put_image(element) 
                case ElementType.Table:
                    if element.content:
                        table_content_processed = []
                        for row in element.content:
                            processed_row = []
                            for cell_runs_or_str in row:
                                if isinstance(cell_runs_or_str, list):
                                    processed_row.append(self.get_formatted_runs(cell_runs_or_str))
                                elif isinstance(cell_runs_or_str, str):
                                    processed_row.append(self.get_escaped(cell_runs_or_str))
                                else:
                                    processed_row.append('')
                            table_content_processed.append(processed_row)
                        self.put_table(table_content_processed)
                case ElementType.CodeBlock:
                    self.put_code_block(element.content, element.language)
                case ElementType.Formula:
                    if isinstance(element, FormulaElement):
                        self.put_formula(element)
            
            last_element_type = element.type
        
        if last_element_type == ElementType.ListItem:
            self.put_list_footer()

    def put_header(self):
        self.write(
            r'\documentclass[aspectratio=169]{beamer}' + '\n'
            r'\usetheme{default}' + '\n\n'
            r'\usepackage[utf8]{inputenc}' + '\n'
            r'\usepackage{graphicx}' + '\n'
            r'\usepackage{booktabs}' + '\n'
            r'\usepackage{xcolor}' + '\n'
            r'\usepackage{hyperref}' + '\n'
            r'\usepackage{amsmath}' + '\n'
            r'\usepackage{amssymb}' + '\n'
            r'\usepackage{esint}' + '\n' 
            r'\usepackage{wrapfig}' + '\n'
            r'\usepackage{listings}' + '\n'
            r'% \usepackage{minted}' + '\n\n'
            r'\beamertemplatenavigationsymbolsempty' + '\n'
            r'% \setbeamertemplate{footline}[frame number]' + '\n\n'
            r'% \title{Presentation Title}' + '\n'
            r'% \author{Author Name}' + '\n'
            r'% \date{\today}' + '\n\n'
            r'\begin{document}' + '\n\n'
            r'% \maketitle' + '\n\n'
        )

    def output(self, presentation_data: ParsedPresentation):
        self.put_header()
        self.last_title_info: Optional[Tuple[str, int]] = None 
        pres_original_slide_width_px = self.config.slide_width_px or DEFAULT_SLIDE_WIDTH_PX

        for slide_idx, slide in enumerate(presentation_data.slides):
            # Elements for density calculation (overall slide content)
            initial_all_elements_for_density: List[SlideElement] = []
            # Elements that will be separated (title, floated images, other initial content)
            elements_to_separate: List[SlideElement] = []
            
            is_multicolumn_slide_type = slide.type == SlideType.MultiColumn
            original_columns_data: Optional[List[List[SlideElement]]] = None

            if is_multicolumn_slide_type:
                initial_all_elements_for_density = slide.preface + [el for col in (slide.columns or []) for el in col]
                elements_to_separate = slide.preface # Separate only from preface for multicol
                original_columns_data = slide.columns
            else: # General slide type
                initial_all_elements_for_density = slide.elements
                elements_to_separate = slide.elements

            if not initial_all_elements_for_density : # If truly empty after considering all parts
                if slide_idx < len(presentation_data.slides) - 1:
                    self.write(r'\begin{frame}{}\end{frame}' + '\n\n') 
                continue

            # Separate title, floated images, and other content from 'elements_to_separate'
            main_title_element, floated_elements, other_preface_or_general_content = \
                self._separate_slide_elements(
                    elements_to_separate,
                    pres_original_slide_width_px,
                    pres_original_slide_width_px # For Beamer, target hint width is original width
                )

            line_count, _, _, _, _, _ = \
                self._get_slide_content_metrics(initial_all_elements_for_density)
            density_class = self._get_slide_density_class(line_count)
            
            self.write(r'\begin{frame}')
            self.in_frame = True
            
            if main_title_element:
                title_text_runs = main_title_element.content if isinstance(main_title_element.content, list) else None
                title_text_str = main_title_element.content if isinstance(main_title_element.content, str) else None
                formatted_title = ""
                if title_text_runs: formatted_title = self.get_formatted_runs(title_text_runs)
                elif title_text_str: formatted_title = self.get_escaped(title_text_str.strip())
                
                if formatted_title:
                    self.write(f'\n\\frametitle{{{formatted_title}}}\n') 
                    if isinstance(main_title_element.content, str):
                         self.last_title_info = (main_title_element.content.strip(), main_title_element.level)
                    else: 
                         self.last_title_info = (formatted_title, main_title_element.level) # Or derive from runs
            
            current_font_scale_opened = False
            font_scale_prefix = ""
            font_scale_suffix = ""
            if density_class == "small": 
                font_scale_prefix = r'{\small' + "\n"
                font_scale_suffix = "\n}\n"
                current_font_scale_opened = True
            elif density_class == "smaller": 
                font_scale_prefix = r'{\footnotesize' + "\n"
                font_scale_suffix = "\n}\n"
                current_font_scale_opened = True
            elif density_class == "smallest": 
                font_scale_prefix = r'{\scriptsize' + "\n"
                font_scale_suffix = "\n}\n"
                current_font_scale_opened = True
            
            if current_font_scale_opened: self.write(font_scale_prefix)

            # Output floated images first
            if floated_elements:
                self._put_elements_on_slide(floated_elements)

            # Now handle the remaining content (other_preface_or_general_content) and original_columns_data
            
            if is_multicolumn_slide_type:
                # Output remaining preface content (non-title, non-floated from preface)
                if other_preface_or_general_content:
                    self._put_elements_on_slide(other_preface_or_general_content)
                
                # Then process the actual columns
                if original_columns_data:
                    num_cols = len(original_columns_data)
                    if num_cols > 0:
                        self.write(r'\begin{columns}[T]' + '\n')
                        # Ensure col_width calculation is robust, e.g. handles num_cols=0 if it could occur
                        col_width_val = (1.0 / num_cols) if num_cols > 0 else 1.0
                        col_width = f'{col_width_val:.2f}'

                        for column_data_list in original_columns_data:
                            self.write(r'  \column{' + col_width + r'\textwidth}' + '\n')
                            self._put_elements_on_slide(column_data_list)
                        self.write(r'\end{columns}' + '\n')
            else: # General slide type, apply heuristic column splitting if needed
                actually_split_columns_heuristic = False
                if other_preface_or_general_content: # Check based on remaining content
                    _, _, _, _, ca_text_lines, ca_text_chars = self._get_slide_content_metrics(other_preface_or_general_content)
                    # Use a density class based on this remaining content for splitting decision
                    # Or use the overall slide_density_class (density_class variable)
                    # Let's use a specific density check for column content:
                    content_density_for_cols_class = self._get_slide_density_class(ca_text_lines)

                    initial_split_qualification = False
                    if content_density_for_cols_class in ["smaller", "smallest"]: # or density_class
                        if ca_text_lines > 0:
                            avg_line_length = ca_text_chars / ca_text_lines
                            if avg_line_length < self.config.beamer_columns_line_length_threshold: # Configurable threshold
                                initial_split_qualification = True
                    
                    contains_table_in_content = any(el.type == ElementType.Table for el in other_preface_or_general_content)
                    
                    actually_split_columns_heuristic = initial_split_qualification and \
                                             len(other_preface_or_general_content) >= 2 and \
                                             not contains_table_in_content
                
                if actually_split_columns_heuristic:
                    num_in_first_col = (len(other_preface_or_general_content) + 1) // 2
                    first_half_elements = other_preface_or_general_content[:num_in_first_col]
                    second_half_elements = other_preface_or_general_content[num_in_first_col:]

                    self.write(r'\begin{columns}[T]' + '\n')
                    self.write(r'  \column{0.48\textwidth}' + '\n') # Default split
                    self._put_elements_on_slide(first_half_elements)
                    self.write(r'  \column{0.48\textwidth}' + '\n')
                    self._put_elements_on_slide(second_half_elements)
                    self.write(r'\end{columns}' + '\n')
                elif other_preface_or_general_content: # Not splitting, print as single block
                    self._put_elements_on_slide(other_preface_or_general_content)

            if not self.config.disable_notes and slide.notes:
                escaped_notes = [self.get_escaped(note) for note in slide.notes]
                self.write(r'\note{' + '\n'.join(escaped_notes) + '}\n')

            if current_font_scale_opened: 
                self.write(font_scale_suffix)
            
            self.write(r'\end{frame}' + '\n\n')
            self.in_frame = False

        self.write(r'\end{document}' + '\n')
        self.close() # Uses base Formatter.close to write self._buffer to file

    def put_title(self, text: str, level: int):
        if level == 1:
            self.write(f'\\begin{{block}}{{{text}}}\n\\end{{block}}\n\n') 
        elif level == 2:
            self.write(f'\\textbf{{{text}}}\\par\n\n')
        else:
            self.write(f'\\textit{{{text}}}\\par\n\n')

    def put_list(self, text: str, level: int):
        MAX_LATEX_LIST_LEVEL = 3 # Standard LaTeX/Beamer itemize depth is 3 levels

        # input `level` is 0-indexed from parser.
        # Clamp the effective level for LaTeX generation to avoid exceeding MAX_LATEX_LIST_LEVEL.
        clamped_parser_level = min(level, MAX_LATEX_LIST_LEVEL - 1) # 0-indexed, capped (0 to 3)
        target_latex_nest_level = clamped_parser_level + 1          # 1-indexed, capped (1 to 4)
        
        while self.current_list_level < target_latex_nest_level:
            indent_str = '  ' * self.current_list_level
            self.write(indent_str + r'\begin{itemize}' + '\n')
            self.current_list_level += 1
        
        while self.current_list_level > target_latex_nest_level:
            self.current_list_level -= 1
            indent_str = '  ' * self.current_list_level
            self.write(indent_str + r'\end{itemize}' + '\n')

        # Indent the item based on its (clamped) LaTeX nesting level
        item_indent_str = '  ' * clamped_parser_level 
        self.write(item_indent_str + r'\item ' + text.strip() + '\n')

    def put_para(self, text: str):
        self.write(text + '\n\n')

    def put_image(self, element: ImageElement):
        # Convert to forward slashes first, then escape for LaTeX URL/path context
        path_with_forward_slashes = str(element.path).replace('\\', '/')
        image_path_latex = self.get_escaped(path_with_forward_slashes, is_url=True) 
        caption_text = self.get_escaped(element.alt_text) if element.alt_text else None

        position_hint = "center" 
        wrapfig_char_placement = None 

        if element.left_px is not None and element.display_width_px is not None and \
           self.config.slide_width_px and self.config.slide_width_px > 0:
            ppt_slide_w = self.config.slide_width_px
            image_center_ppt = element.left_px + (element.display_width_px / 2)
            if image_center_ppt < ppt_slide_w / 3.0: position_hint = "left"; wrapfig_char_placement = "l" 
            elif image_center_ppt > ppt_slide_w * (2/3.0): position_hint = "right"; wrapfig_char_placement = "r" 
        
        effective_position_hint = getattr(element, 'position_hint', position_hint)
        if effective_position_hint == "left": wrapfig_char_placement = "l"
        elif effective_position_hint == "right": wrapfig_char_placement = "r"
        
        wf_width_frac = 0.4 
        if element.display_width_px and self.config.slide_width_px and self.config.slide_width_px > 0:
            ppt_img_frac_of_slide = element.display_width_px / self.config.slide_width_px
            wf_width_frac = min(max(0.25, ppt_img_frac_of_slide), 0.6)

        includegraphics_opts_str = ""
        if wrapfig_char_placement:
            includegraphics_opts_str = r"width=\linewidth,keepaspectratio"
        else:
            center_img_width_frac = 0.7
            if element.display_width_px and self.config.slide_width_px and self.config.slide_width_px > 0:
                 ppt_img_frac_of_slide = element.display_width_px / self.config.slide_width_px
                 center_img_width_frac = min(max(0.2, ppt_img_frac_of_slide), 0.85)
            includegraphics_opts_str = f"width={center_img_width_frac:.2f}\\textwidth,keepaspectratio"

        if wrapfig_char_placement and (effective_position_hint == "left" or effective_position_hint == "right") and not self.config.disable_image_wrapping:
            self.write(f'\\begin{{wrapfigure}}{{{wrapfig_char_placement}}}{{{wf_width_frac:.2f}\\linewidth}}\n')
            self.write(r'  \centering' + '\n') 
            self.write(f'  \\includegraphics[{includegraphics_opts_str}]{{{image_path_latex}}}\n')
            if caption_text and not self.config.disable_captions:
                self.write(f'  \\caption{{{caption_text}}}\n')
            self.write(r'\end{wrapfigure}' + '\n')
        else:
            self.write(r'\begin{figure}[H]' + '\n') 
            self.write(r'  \centering' + '\n')
            self.write(f'  \\includegraphics[{includegraphics_opts_str}]{{{image_path_latex}}}\n')
            if caption_text and not self.config.disable_captions:
                self.write(f'  \\caption{{{caption_text}}}\n')
            self.write(r'\end{figure}' + '\n\n')

    def put_table(self, table: List[List[str]]):
        if not table or not table[0]: return
        num_cols = len(table[0])
        col_spec = 'l' * num_cols
        
        self.write(r'\begin{table}[H]' + '\n')
        self.write(r'  \centering' + '\n')
        self.write(f'  \\begin{{tabular}}{{{col_spec}}}\n')
        self.write(r'    \toprule' + '\n')
        header_row_latex = ' & '.join([cell for cell in table[0]]) + r' \\' + '\n'
        self.write('    ' + header_row_latex)
        self.write(r'    \midrule' + '\n')
        for row_data in table[1:]:
            row_latex = ' & '.join([cell for cell in row_data]) + r' \\' + '\n'
            self.write('    ' + row_latex)
        self.write(r'    \bottomrule' + '\n')
        self.write(r'  \end{tabular}' + '\n')
        self.write(r'\end{table}' + '\n\n')

    def put_code_block(self, code: str, language: Optional[str]):
        lines = code.splitlines() # Split into a list of lines
        if not lines and code.strip() == "": # Handle truly empty or whitespace-only code block gracefully
            # If the original code was just whitespace, it might result in empty `lines`
            # but we might still want a visual indication of an attempted code block, e.g., a small space.
            # For now, let's just ensure a paragraph break if it was meant to be a block.
            # If there were no lines at all (empty string input), this will also do nothing if write buffer is empty.
            self.write('\n') 
            return
        
        # Ensure there's some vertical separation before the code block if it's not the first element.
        # self.write('\medskip\noindent') # Optional: add some space and prevent indentation

        for line in lines:
            # self.get_inline_code handles escaping for \texttt and wraps it.
            texttt_line = self.get_inline_code(line.rstrip('\r')) # rstrip to remove potential \r from \r\n
            # Using \par for a paragraph break after each line.
            # Add an explicit newline in the .tex source for readability.
            self.write(f'{texttt_line}\par\n')
        
        # Ensure separation after the block too, if desired.
        # self.write('\medskip\n')

    def put_formula(self, element: FormulaElement):
        content = element.content.strip()
        
        # Replace newlines within the formula content with LaTeX math newlines ' \\ '.
        # The added spaces around \\ are for robustness, and a newline afterwards in the source for readability.
        processed_formula_text = content.replace('\n', ' \\\\ ')

        if content.startswith('$$') and content.endswith('$$'):
            # Original content had $$, we extract the inner part for \[\]
            math_content_inner = content[2:-2].strip()
            processed_math_content_inner = math_content_inner.replace('\n', ' \\\\ ')
            self.write(f'\\[\n{processed_math_content_inner}\n\\]\n\n')
        elif content.startswith('$') and content.endswith('$'):
            # For inline math $...$, replace internal newlines and write as is.
            # It's unusual for inline math to have newlines, but handle defensively.
            self.write(f'{processed_formula_text}\n\n') 
        else:
            # Assume content is display math needing \[ ... \] but without $$ delimiters originally.
            # Use the processed_formula_text which has newlines replaced.
            self.write(f'\\[\n{processed_formula_text}\n\\]\n\n')

    def get_inline_code(self, text: str) -> str:
        return r'\texttt{' + self.get_escaped(text, verbatim_like=True) + r'}'

    def get_accent(self, text):
        return self._format_text_with_delimiters(text, r'\textit{', '}')

    def get_strong(self, text):
        return self._format_text_with_delimiters(text, r'\textbf{', '}')

    def get_colored(self, text, rgb):
        r_val, g_val, b_val = rgb
        return f'\\textcolor[RGB]{{{r_val},{g_val},{b_val}}}{{{text}}}'

    def get_hyperlink(self, text, url):
        # Convert to forward slashes first for URLs, then escape for LaTeX URL context
        url_with_forward_slashes = str(url).replace('\\', '/')
        escaped_url = self.get_escaped(url_with_forward_slashes, is_url=True)
        return r'\href{' + escaped_url + r'}{' + text + r'}'

    def esc_repl(self, match, verbatim_like=False, is_url=False):
        char = match.group(0)
        if verbatim_like:
            # For \texttt{} and similar contexts:
            # \, {, } need specific escapes.
            # Other special characters from esc_map (like |, _, $, %) also need escaping.
            if char == '\\': return r'\textbackslash{}'
            if char == '{': return r'\{'
            if char == '}': return r'\}'
            # Fall through to the main escape map for other characters.
            # This ensures robust escaping for _ , ^, &, %, $, #, and crucially | -> \textbar{}
            return self.esc_map.get(char, char)

        if is_url:
            # For URLs/paths, common problematic characters need escaping for LaTeX.
            # Backslash should NOT become \textbackslash{}. It's converted to / before this.
            # If a raw backslash somehow reaches here in a URL context (it shouldn't if pre-converted to /),
            # it's tricky. For now, assume it has been converted to '/'.
            if char == '%': return r'\%'
            if char == '#': return r'\#'
            if char == '&': return r'\&'
            if char == '_': return r'\_'
            # '~', '^', '<', '>' are less common in file paths but can be in URLs.
            # Keep their specific LaTeX escapes if needed for general URL text.
            # For file paths, they are usually not special unless the filename itself contains them.
            if char == '~': return r'\textasciitilde{}'
            if char == '^': return r'\textasciicircum{}'
            # '<' and '>' are already handled by the main esc_map if they are part of self.esc_re
            # For simple file paths, most other characters are fine or handled by graphicx.
            # The default self.esc_map.get(char, char) will handle other special chars like $, {, }
            # if they are in self.esc_re.
            return self.esc_map.get(char, char) # Fallback to main map for other chars in URL
        
        # Default LaTeX escaping for non-URL, non-verbatim contexts
        return self.esc_map.get(char, char)

    def get_escaped(self, text, verbatim_like=False, is_url=False):
        if self.config.disable_escaping:
            return text
        # Ensure text is string
        text_str = str(text)
        return self.esc_re.sub(lambda m: self.esc_repl(m, verbatim_like, is_url), text_str)

    def put_list_header(self):
        pass

    def put_list_footer(self):
        while self.current_list_level > 0:
            self.current_list_level -= 1
            indent_str = '  ' * self.current_list_level
            self.write(indent_str + r'\end{itemize}' + '\n')
        self.current_list_level = 0 

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

    def _get_slide_content_metrics(self, elements: List[SlideElement]) -> Tuple[int, int, Optional[int], Optional[int], int, int]:
        line_count = 0
        char_count = 0
        max_img_w: Optional[int] = None
        max_img_h: Optional[int] = None
        text_lines_for_avg = 0 # For calculating average line length, excluding titles
        text_chars_for_avg = 0 # For calculating average line length, excluding titles

        for element in elements:
            content_str = ""
            is_title = element.type == ElementType.Title

            if element.type in [ElementType.Title, ElementType.Paragraph, ElementType.ListItem]:
                if isinstance(element.content, list): # List of TextRuns
                    for run in element.content:
                        content_str += run.text
                elif isinstance(element.content, str):
                    content_str = element.content
                
                lines_in_element = content_str.count('\n') + 1 if content_str else 0
                line_count += lines_in_element
                char_count += len(content_str)
                if not is_title:
                    text_lines_for_avg += lines_in_element
                    text_chars_for_avg += len(content_str)

            elif isinstance(element, ImageElement):
                if element.display_width_px is not None:
                    max_img_w = max(max_img_w or 0, element.display_width_px)
                if element.display_height_px is not None:
                    max_img_h = max(max_img_h or 0, element.display_height_px)
                # Add a nominal line count for an image to contribute to density
                line_count += self.config.image_density_line_equivalent
                if not is_title: # Unlikely for an image to be a title, but for consistency
                    text_lines_for_avg += self.config.image_density_line_equivalent


            elif element.type == ElementType.Table:
                if element.content: # List of lists (rows of cells)
                    num_rows = len(element.content)
                    line_count += num_rows * self.config.table_row_density_line_equivalent
                    if not is_title:
                         text_lines_for_avg += num_rows * self.config.table_row_density_line_equivalent
                    for row in element.content:
                        for cell in row:
                            cell_str = ""
                            if isinstance(cell, list): # List of TextRuns
                                for run in cell: cell_str += run.text
                            elif isinstance(cell, str): cell_str = cell
                            char_count += len(cell_str)
                            if not is_title: text_chars_for_avg += len(cell_str)
            
            elif element.type == ElementType.CodeBlock and isinstance(element.content, str):
                lines_in_code = element.content.count('\n') + 1 if element.content else 0
                line_count += lines_in_code
                char_count += len(element.content)
                if not is_title:
                    text_lines_for_avg += lines_in_code
                    text_chars_for_avg += len(element.content)

        return line_count, char_count, max_img_w, max_img_h, text_lines_for_avg, text_chars_for_avg

    def _get_slide_density_class(self, line_count: int) -> Optional[str]:
        if line_count >= self.config.smallest_font_line_threshold:
            return "smallest"
        elif line_count >= self.config.smaller_font_line_threshold:
            return "smaller"
        elif line_count >= self.config.small_font_line_threshold:
            return "small"
        return None

    def _format_text_with_delimiters(self, text: str, start_delim: str, end_delim: str) -> str:
        # Helper for simple formatting like bold or italic
        # Recursively apply to handle nested TextRuns if text is a list
        if isinstance(text, list) and all(isinstance(run, TextRun) for run in text):
            return start_delim + self.get_formatted_runs(text) + end_delim
        return start_delim + str(text) + end_delim
        
    def get_formatted_runs(self, runs: List[TextRun]) -> str:
        if not runs:
            return ""
        
        formatted_texts: List[str] = []
        for run in runs:
            text_content = self.get_escaped(run.text)
            
            if run.style.is_code:
                text_content = self.get_inline_code(run.text) # Use raw run.text for code
            else:
                if run.style.is_accent and run.style.is_strong:
                    # Apply strong first, then accent for combined effect (e.g., **_text_**)
                    # Or let specific formatters decide combined style if they have one.
                    # This simple nesting might not always be ideal for all markdown flavors.
                    text_content = self.get_accent(self.get_strong(text_content))
                elif run.style.is_strong:
                    text_content = self.get_strong(text_content)
                elif run.style.is_accent:
                    text_content = self.get_accent(text_content)

                if run.style.color_rgb:
                    text_content = self.get_colored(text_content, run.style.color_rgb)
            
            if run.style.hyperlink:
                # For hyperlinks, the text_content is already formatted (bold, italic, color).
                # The raw run.text should be used if the link text itself shouldn't inherit formatting.
                # However, usually the displayed text part of a link can be styled.
                text_content = self.get_hyperlink(text_content, run.style.hyperlink)

            formatted_texts.append(text_content)
        return "".join(formatted_texts) 