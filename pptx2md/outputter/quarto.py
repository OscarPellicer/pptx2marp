# Modified by Oscar Pellicer, 2025
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

import re
import urllib.parse
from typing import List, Optional

from rapidfuzz import fuzz

from pptx2md.outputter.base import Formatter
from pptx2md.types import ParsedPresentation, SlideElement, ElementType, SlideType, FormulaElement, ImageElement # ImageElement is used via put_image
from pptx2md.utils import rgb_to_hex

class QuartoFormatter(Formatter):
    # write outputs to quarto markdown - reveal js
    def __init__(self, config):
        super().__init__(config)
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)\#\+-\.\|])') # Added | for tables
        self.esc_re2 = re.compile(r'(<[^>]+>)')

    def output(self, presentation_data: ParsedPresentation):
        self.put_header() # Uses self.write -> buffer

        last_title_tracker = { # Quarto specific title tracking within its output method
            'content': None,
            'level': -1
        }

        def put_elements(elements: List[SlideElement]):
            nonlocal last_title_tracker
            last_element_type: Optional[ElementType] = None
            for element in elements:
                if last_element_type and last_element_type == ElementType.ListItem and element.type != ElementType.ListItem:
                    self.put_list_footer()

                current_content_str = ""
                if element.type in [ElementType.Title, ElementType.Paragraph, ElementType.ListItem]:
                    if isinstance(element.content, list):
                        current_content_str = self.get_formatted_runs(element.content)
                    elif isinstance(element.content, str):
                        # For Quarto, non-run string content should also be escaped if it might contain Markdown specials
                        current_content_str = self.get_escaped(element.content)

                match element.type:
                    case ElementType.Title:
                        title_text = current_content_str.strip()
                        if title_text:
                            is_similar_to_last = False
                            if last_title_tracker['content'] and last_title_tracker['level'] == element.level and \
                               fuzz.ratio(last_title_tracker['content'], title_text, score_cutoff=92):
                                is_similar_to_last = True
                            
                            if is_similar_to_last:
                                if self.config.keep_similar_titles:
                                    self.put_title(f'{title_text} (cont.)', element.level) 
                                # else skip
                            else:
                                self.put_title(title_text, element.level)
                            
                            last_title_tracker['content'] = title_text # Store formatted for consistent comparison
                            last_title_tracker['level'] = element.level

                    case ElementType.ListItem:
                        if not (last_element_type and last_element_type == ElementType.ListItem):
                            self.put_list_header()
                        self.put_list(current_content_str, element.level)
                    case ElementType.Paragraph:
                        self.put_para(current_content_str)
                    case ElementType.Image:
                        # Assuming element is ImageElement based on usage
                        self.put_image(element)
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
                            self.put_formula(element)
                last_element_type = element.type
            
            if last_element_type == ElementType.ListItem:
                self.put_list_footer()


        for slide_idx, slide in enumerate(presentation_data.slides):
            if slide.type == SlideType.General:
                put_elements(slide.elements)
            elif slide.type == SlideType.MultiColumn:
                put_elements(slide.preface)
                if slide.columns: # Ensure there are columns
                    num_cols = len(slide.columns)
                    if num_cols == 2: width_val = '50%'
                    elif num_cols == 3: width_val = '33%'
                    else: width_val = f'{100/num_cols:.0f}%' if num_cols > 0 else '100%'

                    self.put_para(':::: {.columns}')
                    for column_elements in slide.columns:
                        self.put_para(f'::: {{.column width="{width_val}"}}')
                        put_elements(column_elements)
                        self.put_para(':::')
                    self.put_para('::::')

            if not self.config.disable_notes and slide.notes:
                self.put_para("::: {.notes}")
                for note in slide.notes: # notes are List[str]
                    self.put_para(self.get_escaped(note))
                self.put_para(":::")

            if slide_idx < len(presentation_data.slides) - 1 and self.config.enable_slides:
                self.put_para("\n---\n")

        self.close() # Calls base Formatter.close()

    def put_header(self):
        # Default Quarto header, user can customize the qmd file later
        self.write('''---
title: "Presentation Title"
author: "Author"
format: 
  revealjs:
    slide-number: c/t
    width: 1600
    height: 900
    # logo: img/logo.png # User should provide their own logo
    # footer: "Organization" # User can add footer
    incremental: true
    theme: [simple] # A basic theme
---
''')

    def put_title(self, text, level):
        self.write('#' * level + ' ' + text + '\n\n')

    def put_list(self, text, level):
        self.write('  ' * level + '* ' + text.strip() + '\n')

    def put_para(self, text):
        self.write(text + '\n\n')

    def put_image(self, element: ImageElement):
        # Convert to forward slashes for URLs
        path_with_forward_slashes = str(element.path).replace('\\', '/')
        quoted_path = urllib.parse.quote(path_with_forward_slashes)
        
        # Use alt_text if available, otherwise use a default
        alt_text = element.alt_text if element.alt_text else "Image"
        
        if element.display_width_px is None:
            self.write(f'![{alt_text}]({quoted_path})\n\n')
        else:
            # Quarto specific width: ![alt text](image.png){width="80%" height="300px"}
            self.write(f'![{alt_text}]({quoted_path}){{width="{element.display_width_px}px"}}\n\n')

    def put_table(self, table: List[List[str]]):
        # Quarto uses standard Pandoc Markdown tables, centered by default
        # Base Formatter.put_table provides left-aligned, let's make it centered for Quarto
        if not table or not table[0]: return
        gen_table_row = lambda row: '| ' + ' | '.join([
            c.replace('\n', '<br />') if '`' not in c else c.replace('\n', ' ') 
            for c in row
        ]) + ' |'
        self.write(gen_table_row(table[0]) + '\n')
        self.write(gen_table_row([':-:' for _ in table[0]]) + '\n') # Centered for Quarto
        self.write('\n'.join([gen_table_row(row) for row in table[1:]]) + '\n\n')

    def put_code_block(self, code: str, language: Optional[str]):
        lang_tag = language if language else ""
        # Quarto can use ```{language} or ```language
        self.write(f'```{lang_tag}\n{code.strip()}\n```\n\n')

    # put_formula inherited from base Formatter for $$...$$

    # get_accent, get_strong inherited (Markdown style)
    def get_strong(self, text):
        return self._format_text_with_delimiters(text, '**', '**') # Quarto uses **

    def get_colored(self, text, rgb):
        # Text is already escaped by _format_single_merged_run
        return ' <span style="color:%s">%s</span> ' % (rgb_to_hex(rgb), text)

    # get_hyperlink inherited (Markdown style)

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