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
from typing import List, Optional # For type hinting

from pptx2md.outputter.base import Formatter
from pptx2md.utils import rgb_to_hex # For get_colored
from pptx2md.types import ImageElement # Add this import

class WikiFormatter(Formatter):
    # write outputs to wikitext
    def __init__(self, config):
        super().__init__(config)
        self.esc_re = re.compile(r'<([^>]+)>_FIX_ME_WIKI_ESCAPING_<') # Placeholder, review wiki escaping needs
        # Original was: self.esc_re = re.compile(r'<([^>]+)>')
        # This regex seems to target HTML tags for escaping. 
        # Wiki escaping is usually about characters like [, ], |, =, etc.
        # For now, I will use a more common wiki char escaping approach.
        self.wiki_esc_map = {
            '[': '&#91;',
            ']': '&#93;',
            '|': '&#124;',
            '{': '&#123;',
            '}': '&#125;',
            '=': '&#61;', # Often problematic in wiki
            '*': '&#42;',
            '#': '&#35;',
            "'": '&#39;', # For 'strong' and 'accent'
            '<': '&lt;',
            '>': '&gt;',
            '&': '&amp;'
        }
        self.wiki_esc_re = re.compile('|'.join(re.escape(key) for key in self.wiki_esc_map.keys()))

    def put_title(self, text, level):
        self.write('=' * (level + 1) + ' ' + text + ' ' + '=' * (level + 1) + '\n\n') # Common wiki title

    def put_list(self, text, level):
        self.write('*' * (level + 1) + ' ' + text.strip() + '\n')

    def put_para(self, text):
        self.write(text + '\n\n')

    def put_image(self, element: ImageElement): # Changed signature
        # element is an ImageElement object
        img_path_escaped = self.get_escaped(str(element.path))
        options = []
        # Use element.display_width_px as the max_width, if available
        # or element.original_width_px as a fallback.
        # The choice depends on desired behavior. display_width_px is often what's intended.
        max_width_val = element.display_width_px
        if max_width_val:
            options.append(f'{int(max_width_val)}px') # Ensure it's an int if it's float

        img_str = f'[[File:{img_path_escaped}' 
        if options:
            img_str += '|' + '|'.join(options)
        
        # Consider adding alt text as caption if available and desired
        # if element.alt_text:
        #     if not options: img_str += '|' # Separator for options if none yet
        #     img_str += f'|{self.get_escaped(element.alt_text)}' # Basic caption

        img_str += ']]\n\n'
        self.write(img_str)

    def put_code_block(self, code: str, language: Optional[str]):
        # Wiki usually uses <syntaxhighlight lang="language">
        lang_attr = f' lang="{language}"' if language else ""
        self.write(f'<syntaxhighlight{lang_attr}>\n{self.get_escaped(code.strip())}\n</syntaxhighlight>\n\n')

    def put_table(self, table: List[List[str]]):
        if not table or not table[0]: return
        self.write('{| class="wikitable"\n')
        replace_newline = lambda x: x.replace("\n", "<br />")
        # Header
        header_cells = [f'! {self.get_escaped(replace_newline(cell))}' for cell in table[0]]
        self.write(' '.join(header_cells) + '\n')
        # Body rows
        for row_data in table[1:]:
            self.write('|-\n')
            row_cells = [f'| {self.get_escaped(replace_newline(cell))}' for cell in row_data]
            self.write('\n'.join(row_cells) + '\n') # One cell per line in this common format
        self.write('|}\n\n')

    def get_accent(self, text):
        # Wiki typically uses '' for italics
        return self._format_text_with_delimiters(text, "''", "''")

    def get_strong(self, text):
        # Wiki typically uses ''' for bold
        return self._format_text_with_delimiters(text, "'''", "'''")

    def get_colored(self, text, rgb):
        # Wiki might support <span style="color:..."> or specific templates
        # Using HTML version for broader compatibility if allowed by wiki
        # Text is already escaped by _format_single_merged_run
        return ' <span style="color:%s">%s</span> ' % (rgb_to_hex(rgb), text)

    def get_hyperlink(self, text, url):
        # Wiki: [url text] or [[Page Name|text]]
        # Assuming external link here
        # Text and URL should be escaped for wiki syntax if they contain special chars
        return '[' + self.get_escaped(url) + ' ' + self.get_escaped(text) + ']'

    def _wiki_esc_repl(self, match):
        char = match.group(0)
        return self.wiki_esc_map.get(char, char)

    def get_escaped(self, text):
        if self.config.disable_escaping:
            return text
        # First, handle general XML/HTML-like escapes if any text might be HTML itself
        # text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;') # Basic HTML escape
        # Then apply specific Wiki character escaping
        return self.wiki_esc_re.sub(self._wiki_esc_repl, text)

    # Original esc_repl and esc_re seem for HTML tags, which might be too aggressive or not what is needed for general wiki text.
    # The base get_escaped returns text as is, which is then overridden here. 