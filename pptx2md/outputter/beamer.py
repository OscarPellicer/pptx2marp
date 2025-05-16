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

from pptx2md.outputter.base import Formatter
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
            
            self.write(r'\begin{frame}')
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
                    self.write(f'\n\\frametitle{{{formatted_title}}}\n') 
                    if isinstance(main_title_element.content, str):
                         self.last_title_info = (main_title_element.content.strip(), main_title_element.level)
                    else: 
                         self.last_title_info = (formatted_title, main_title_element.level)
            else:
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

            actually_split_columns_heuristic = False
            if not is_multicolumn_slide_type and content_after_title:
                _, _, _, _, ca_text_lines, ca_text_chars = self._get_slide_content_metrics(content_after_title)
                current_content_density_class = self._get_slide_density_class(ca_text_lines)

                initial_split_qualification = False
                if current_content_density_class in ["smaller", "smallest"]:
                    if ca_text_lines > 0:
                        avg_line_length = ca_text_chars / ca_text_lines
                        if avg_line_length < 40:
                            initial_split_qualification = True
                
                contains_table_in_content = any(el.type == ElementType.Table for el in content_after_title)
                
                actually_split_columns_heuristic = initial_split_qualification and \
                                         len(content_after_title) >= 2 and \
                                         not contains_table_in_content

            if is_multicolumn_slide_type and original_columns_data:
                if main_title_element and slide_elements_for_processing == slide.preface:
                     self._put_elements_on_slide(content_after_title) 
                elif not main_title_element and slide.preface:
                     self._put_elements_on_slide(slide.preface)

                num_cols = len(original_columns_data)
                if num_cols > 0:
                    self.write(r'\begin{columns}[T]' + '\n')
                    col_width = f'{1/num_cols:.2f}'
                    for column_data_list in original_columns_data:
                        self.write(r'  \column{' + col_width + r'\textwidth}' + '\n')
                        self._put_elements_on_slide(column_data_list)
                    self.write(r'\end{columns}' + '\n')

            elif actually_split_columns_heuristic:
                num_in_first_col = (len(content_after_title) + 1) // 2
                first_half_elements = content_after_title[:num_in_first_col]
                second_half_elements = content_after_title[num_in_first_col:]

                self.write(r'\begin{columns}[T]' + '\n')
                self.write(r'  \column{0.48\textwidth}' + '\n')
                self._put_elements_on_slide(first_half_elements)
                self.write(r'  \column{0.48\textwidth}' + '\n')
                self._put_elements_on_slide(second_half_elements)
                self.write(r'\end{columns}' + '\n')
            else:
                self._put_elements_on_slide(content_after_title)
            
            if not self.config.disable_notes and slide.notes:
                escaped_notes = [self.get_escaped(note) for note in slide.notes]
                self.write(r'\note{' + '\n'.join(escaped_notes) + '}\n')

            if current_font_scale_opened: 
                self.write("\n}\n")
            
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
        MAX_LATEX_LIST_LEVEL = 4 # Standard LaTeX/Beamer itemize depth is 4 levels

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
        processed_code = code.strip('\n') # Strip leading/trailing newlines only
        lang_opt = f',language={language}' if language and getattr(self.config, 'use_listings', False) else ''
        
        if lang_opt: 
            # lstlisting handles most content well; avoid get_escaped.
            self.write(f'\\begin{{lstlisting}}[basicstyle=\\ttfamily\footnotesize{lang_opt}]\n{processed_code}\n\\end{{lstlisting}}\n\n')
        else:
            # verbatim environment prints content as is.
            self.write(f'\\begin{{verbatim}}\n{processed_code}\n\\end{{verbatim}}\n\n')

    def put_formula(self, element: FormulaElement):
        content = element.content.strip()
        if content.startswith('$$') and content.endswith('$$'):
            math_content = content[2:-2].strip()
            self.write(f'\\[\n{math_content}\n\\]\n\n')
        elif content.startswith('$') and content.endswith('$'):
            self.write(f'{content}\n\n') # Output $...$ math directly
        else:
            # Assume content is display math needing \[ ... \]
            self.write(f'\\[\n{content}\n\\]\n\n')

    def get_formatted_runs(self, runs: List[TextRun]) -> str:
        result_parts = []
        for run in runs:
            current_run_text = run.text if run.text is not None else "" # Ensure text is not None
            
            is_math_style = getattr(run.style, 'is_math', False)
            
            processed_text = ""
            is_run_code_style = getattr(run.style, 'is_code', False)

            if is_math_style:
                processed_text = current_run_text # Math text is used directly
            elif is_run_code_style:
                processed_text = self.get_inline_code(current_run_text)
            else:
                processed_text = self.get_escaped(current_run_text)

            if run.style.is_strong:
                processed_text = self.get_strong(processed_text)
            if run.style.is_accent:
                processed_text = self.get_accent(processed_text)
            if run.style.color_rgb:
                processed_text = self.get_colored(processed_text, run.style.color_rgb)
            if run.style.hyperlink:
                processed_text = self.get_hyperlink(processed_text, run.style.hyperlink)
            
            result_parts.append(processed_text)
        return "".join(result_parts)

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
            if char == '{': return r'\{'
            if char == '}': return r'\}'
            if char == '\\': return r'\textbackslash{}' # Keep for verbatim
            return char
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