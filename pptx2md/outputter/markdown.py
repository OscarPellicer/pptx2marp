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
from typing import List, Optional # For type hinting

from pptx2md.outputter.base import Formatter
from pptx2md.utils import rgb_to_hex # For get_colored
from pptx2md.types import ImageElement
#ConversionConfig, ElementType, ParsedPresentation, SlideElement, SlideType, TextRun, ImageElement, FormulaElement, TextStyle

class MarkdownFormatter(Formatter):
    # write outputs to markdown
    def __init__(self, config):
        super().__init__(config)
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)\#\+-\.\|])') # Added | to esc_re1 for tables
        self.esc_re2 = re.compile(r'(<[^>]+>)')

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
            self.write(f'<img src="{quoted_path}" alt="{alt_text}" style="max-width:{element.display_width_px}px;" />\n\n')

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

    # get_accent, get_strong inherited from base use '_' and '__' which is fine for Markdown

    def get_colored(self, text, rgb):
        # Text is already escaped by _format_single_merged_run
        return ' <span style="color:%s">%s</span> ' % (rgb_to_hex(rgb), text)

    # get_hyperlink inherited from base is fine: [text](url)

    def esc_repl(self, match):
        return '\\' + match.group(0)

    def get_escaped(self, text):
        if self.config.disable_escaping:
            return text
        text = re.sub(self.esc_re1, self.esc_repl, text)
        text = re.sub(self.esc_re2, self.esc_repl, text)
        return text 