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
from pptx2md.utils import rgb_to_hex
from pptx2md.types import ImageElement

class MadokoFormatter(Formatter):
    # write outputs to madoko markdown
    def __init__(self, config):
        super().__init__(config)
        self.write('[TOC]\n\n') # Use self.write for TOC
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)\#\+-\.\|])') # Added | for tables
        self.esc_re2 = re.compile(r'(<[^>]+>)')

    def put_title(self, text, level):
        self.write('#' * level + ' ' + text + '\n\n')

    def put_list(self, text, level):
        self.write('  ' * level + '* ' + text.strip() + '\n')

    def put_para(self, text):
        self.write(text + '\n\n')

    def put_image(self, element: ImageElement): # Changed signature
        # element is an ImageElement object
        
        # Paths in Markdown should typically use forward slashes.
        # Ensure element.path is a string and correctly formatted.
        img_path = str(element.path).replace('\\', '/') 
        
        alt_text = self.get_escaped(element.alt_text if element.alt_text else "")
        
        attributes = []
        # Use element.display_width_px for the width if available
        max_width_val = element.display_width_px
        if max_width_val:
            attributes.append(f'width="{int(max_width_val)}px"') # Ensure width is integer

        # Madoko might also support height, title, etc. Add them to attributes if needed.
        # e.g., if element.title: attributes.append(f'title="{self.get_escaped(element.title)}"')

        attr_str = ""
        if attributes:
            attr_str = '{' + ' '.join(attributes) + '}'
            
        # Basic Markdown image syntax: ![alt_text](path)
        # With attributes: ![alt_text](path){attributes}
        self.write(f'![{alt_text}]({img_path}){attr_str}\n\n')

    # put_table will be inherited from base Formatter (Markdown-like, left-aligned)
    # If Madoko has specific table needs, it can be overridden here.

    def put_code_block(self, code: str, language: Optional[str]):
        lang_tag = language if language else ""
        self.write(f'```{lang_tag}\n{code.strip()}\n```\n\n')

    # get_accent, get_strong inherited (Markdown _ and __)

    def get_colored(self, text, rgb):
        # Text is already escaped by _format_single_merged_run
        return ' <span style="color:%s">%s</span> ' % (rgb_to_hex(rgb), text)

    # get_hyperlink inherited (Markdown [text](url))

    def esc_repl(self, match):
        return '\\' + match.group(0)

    def get_escaped(self, text):
        if self.config.disable_escaping:
            return text
        text = re.sub(self.esc_re1, self.esc_repl, text)
        text = re.sub(self.esc_re2, self.esc_repl, text)
        return text 