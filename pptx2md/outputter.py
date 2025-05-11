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
from typing import List, Tuple, Optional

from rapidfuzz import fuzz

from pptx2md.types import ConversionConfig, ElementType, ParsedPresentation, SlideElement, SlideType, TextRun, ImageElement
from pptx2md.utils import rgb_to_hex


class Formatter:

    def __init__(self, config: ConversionConfig):
        os.makedirs(config.output_path.parent, exist_ok=True)
        self.ofile = open(config.output_path, 'w', encoding='utf8')
        self.config = config

    def output(self, presentation_data: ParsedPresentation):
        self.put_header()

        last_element = None
        last_title = None
        for slide_idx, slide in enumerate(presentation_data.slides):
            all_elements = []
            if slide.type == SlideType.General:
                all_elements = slide.elements
            elif slide.type == SlideType.MultiColumn:
                all_elements = slide.preface + slide.columns

            for element in all_elements:
                if last_element and last_element.type == ElementType.ListItem and element.type != ElementType.ListItem:
                    self.put_list_footer()

                match element.type:
                    case ElementType.Title:
                        element.content = element.content.strip()
                        if element.content:
                            if last_title and last_title.level == element.level and fuzz.ratio(
                                    last_title.content, element.content, score_cutoff=92):
                                # skip if the title is the same as the last one
                                # Allow for repeated slide titles - One or more - Add (cont.) to the title
                                if self.config.keep_similar_titles:
                                    self.put_title(f'{element.content} (cont.)', element.level)
                            else:
                                self.put_title(element.content, element.level)
                            last_title = element
                    case ElementType.ListItem:
                        if not (last_element and last_element.type == ElementType.ListItem):
                            self.put_list_header()
                        self.put_list(self.get_formatted_runs(element.content), element.level)
                    case ElementType.Paragraph:
                        self.put_para(self.get_formatted_runs(element.content))
                    case ElementType.Image:
                        self.put_image(element.path, element.width)
                    case ElementType.Table:
                        self.put_table([[self.get_formatted_runs(cell) for cell in row] for row in element.content])
                last_element = element

            if not self.config.disable_notes and slide.notes:
                self.put_para('---')
                for note in slide.notes:
                    self.put_para(note)

            if slide_idx < len(presentation_data.slides) - 1 and self.config.enable_slides:
                self.put_para("\n---\n")

        self.close()

    def put_header(self):
        pass

    def put_title(self, text, level):
        pass

    def put_list(self, text, level):
        pass

    def put_list_header(self):
        self.put_para('')

    def put_list_footer(self):
        self.put_para('')

    def get_formatted_runs(self, runs: List[TextRun]):
        res = ''
        for run in runs:
            text = run.text
            if text == '':
                continue

            if not self.config.disable_escaping:
                text = self.get_escaped(text)

            if run.style.hyperlink:
                text = self.get_hyperlink(text, run.style.hyperlink)
            if run.style.is_accent:
                text = self.get_accent(text)
            elif run.style.is_strong:
                text = self.get_strong(text)
            if run.style.color_rgb and not self.config.disable_color:
                text = self.get_colored(text, run.style.color_rgb)

            res += text
        return res.strip()

    def put_para(self, text):
        pass

    def put_image(self, path, max_width):
        pass

    def put_table(self, table):
        pass

    def get_accent(self, text):
        pass

    def get_strong(self, text):
        pass

    def get_colored(self, text, rgb):
        pass

    def get_hyperlink(self, text, url):
        pass

    def get_escaped(self, text):
        pass

    def write(self, text):
        self.ofile.write(text)

    def flush(self):
        self.ofile.flush()

    def close(self):
        self.ofile.close()


class MarkdownFormatter(Formatter):
    # write outputs to markdown
    def __init__(self, config: ConversionConfig):
        super().__init__(config)
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)#\+-\.])')
        self.esc_re2 = re.compile(r'(<[^>]+>)')

    def put_title(self, text, level):
        self.ofile.write('#' * level + ' ' + text + '\n\n')

    def put_list(self, text, level):
        self.ofile.write('  ' * level + '* ' + text.strip() + '\n')

    def put_para(self, text):
        self.ofile.write(text + '\n\n')

    def put_image(self, path, max_width=None):
        if max_width is None:
            self.ofile.write(f'![]({urllib.parse.quote(path)})\n\n')
        else:
            self.ofile.write(f'<img src="{path}" style="max-width:{max_width}px;" />\n\n')

    def put_table(self, table):
        gen_table_row = lambda row: '| ' + ' | '.join([c.replace('\n', '<br />') for c in row]) + ' |'
        self.ofile.write(gen_table_row(table[0]) + '\n')
        self.ofile.write(gen_table_row([':-:' for _ in table[0]]) + '\n')
        self.ofile.write('\n'.join([gen_table_row(row) for row in table[1:]]) + '\n\n')

    def get_accent(self, text):
        return ' _' + text + '_ '

    def get_strong(self, text):
        return ' __' + text + '__ '

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
        self.ofile.write('!' * level + ' ' + text + '\n\n')

    def put_list(self, text, level):
        self.ofile.write('*' * (level + 1) + ' ' + text.strip() + '\n')

    def put_para(self, text):
        self.ofile.write(text + '\n\n')

    def put_image(self, path, max_width):
        if max_width is None:
            self.ofile.write(f'<img src="{path}" />\n\n')
        else:
            self.ofile.write(f'<img src="{path}" width={max_width}px />\n\n')

    def get_accent(self, text):
        return ' __' + text + '__ '

    def get_strong(self, text):
        return ' \'\'' + text + '\'\' '

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
        self.ofile.write('[TOC]\n\n')
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)#\+-\.])')
        self.esc_re2 = re.compile(r'(<[^>]+>)')

    def put_title(self, text, level):
        self.ofile.write('#' * level + ' ' + text + '\n\n')

    def put_list(self, text, level):
        self.ofile.write('  ' * level + '* ' + text.strip() + '\n')

    def put_para(self, text):
        self.ofile.write(text + '\n\n')

    def put_image(self, path, max_width):
        if max_width is None:
            self.ofile.write(f'<img src="{path}" />\n\n')
        elif max_width < 500:
            self.ofile.write(f'<img src="{path}" width={max_width}px />\n\n')
        else:
            self.ofile.write('~ Figure {caption: image caption}\n')
            self.ofile.write('![](%s){width:%spx;}\n' % (path, max_width))
            self.ofile.write('~\n\n')

    def get_accent(self, text):
        return ' _' + text + '_ '

    def get_strong(self, text):
        return ' __' + text + '__ '

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
        self.put_header()

        last_title = None

        def put_elements(elements: List[SlideElement]):
            nonlocal last_title

            last_element = None
            for element in elements:
                if last_element and last_element.type == ElementType.ListItem and element.type != ElementType.ListItem:
                    self.put_list_footer()

                match element.type:
                    case ElementType.Title:
                        element.content = element.content.strip()
                        if element.content:
                            if last_title and last_title.level == element.level and fuzz.ratio(
                                    last_title.content, element.content, score_cutoff=92):
                                # skip if the title is the same as the last one
                                # Allow for repeated slide titles - One or more - Add (cont.) to the title
                                if self.config.keep_similar_titles:
                                    self.put_title(f'{element.content} (cont.)', element.level)
                            else:
                                self.put_title(element.content, element.level)
                            last_title = element
                    case ElementType.ListItem:
                        if not (last_element and last_element.type == ElementType.ListItem):
                            self.put_list_header()
                        self.put_list(self.get_formatted_runs(element.content), element.level)
                    case ElementType.Paragraph:
                        self.put_para(self.get_formatted_runs(element.content))
                    case ElementType.Image:
                        self.put_image(element.path, element.width)
                    case ElementType.Table:
                        self.put_table([[self.get_formatted_runs(cell) for cell in row] for row in element.content])
                last_element = element

        for slide_idx, slide in enumerate(presentation_data.slides):
            if slide.type == SlideType.General:
                put_elements(slide.elements)
            elif slide.type == SlideType.MultiColumn:
                put_elements(slide.preface)
                if len(slide.columns) == 2:
                    width = '50%'
                elif len(slide.columns) == 3:
                    width = '33%'
                else:
                    raise ValueError(f'Unsupported number of columns: {len(slide.columns)}')

                self.put_para(':::: {.columns}')
                for column in slide.columns:
                    self.put_para(f'::: {{.column width="{width}"}}')
                    put_elements(column)
                    self.put_para(':::')
                self.put_para('::::')

            if not self.config.disable_notes and slide.notes:
                self.put_para("::: {.notes}")
                for note in slide.notes:
                    self.put_para(note)
                self.put_para(":::")

            if slide_idx < len(presentation_data.slides) - 1 and self.config.enable_slides:
                self.put_para("\n---\n")

        self.close()

    def put_header(self):
        self.ofile.write('''---
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
''')

    def put_title(self, text, level):
        self.ofile.write('#' * level + ' ' + text + '\n\n')

    def put_list(self, text, level):
        self.ofile.write('  ' * level + '* ' + text.strip() + '\n')

    def put_para(self, text):
        self.ofile.write(text + '\n\n')

    def put_image(self, path, max_width=None):
        if max_width is None:
            self.ofile.write(f'![]({urllib.parse.quote(path)})\n\n')
        else:
            self.ofile.write(f'<img src="{path}" style="max-width:{max_width}px;" />\n\n')

    def put_table(self, table):
        gen_table_row = lambda row: '| ' + ' | '.join([c.replace('\n', '<br />') for c in row]) + ' |'
        self.ofile.write(gen_table_row(table[0]) + '\n')
        self.ofile.write(gen_table_row([':-:' for _ in table[0]]) + '\n')
        self.ofile.write('\n'.join([gen_table_row(row) for row in table[1:]]) + '\n\n')

    def get_accent(self, text):
        return ' _' + text + '_ '

    def get_strong(self, text):
        return ' __' + text + '__ '

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


LINES_NORMAL_MAX = 8
LINES_SMALL_MAX = 12
LINES_SMALLER_MAX = 18
LINES_SPLIT_TRIGGER = 18

# Default slide dimensions for position hinting (can be made configurable)
# These will now serve as fallbacks if not provided by config
DEFAULT_SLIDE_WIDTH_PX = 1600
DEFAULT_SLIDE_HEIGHT_PX = 900

MARP_TARGET_WIDTH_PX = 1280
MARP_TARGET_HEIGHT_PX = 720


class MarpFormatter(Formatter):
    # write outputs to marp markdown
    def __init__(self, config: ConversionConfig):
        super().__init__(config)
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)#\+-\.])')
        # self.esc_re2 = re.compile(r'(<[^>]+>)') # Keep commented out for Marp to allow HTML
        self.last_title_info: Optional[Tuple[str, int]] = None # For managing (cont.) and fuzzy match

    def put_header(self):
        self.ofile.write('''---
marp: true
theme: default
paginate: true
---

<style>
section.small {
  font-size: 24px;
}
section.smaller {
  font-size: 20px;
}
section.smallest {
  font-size: 18px;
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
</style>

''')

    def _get_slide_content_metrics(self, elements_list: List[SlideElement]) -> Tuple[int, int]:
        """Calculates number of semantic lines and total characters."""
        line_count = 0
        char_count = 0
        for element in elements_list:
            if element.type == ElementType.Title:
                line_count += 1
                content = element.content.strip() if isinstance(element.content, str) else ""
                char_count += len(content)
            elif element.type == ElementType.ListItem:
                line_count += 1
                if isinstance(element.content, list): # List[TextRun]
                    for run in element.content:
                        char_count += len(run.text)
            elif element.type == ElementType.Paragraph:
                line_count += 1
                if isinstance(element.content, list): # List[TextRun]
                    # Estimate lines within a paragraph by counting newlines, plus one for the para itself
                    para_text = "".join(run.text for run in element.content)
                    char_count += len(para_text)
                    # line_count += para_text.count('\n') # More accurate line count if needed
                elif isinstance(element.content, str): # Should be List[TextRun]
                     char_count += len(element.content)

            # Assuming ElementType.CodeBlock content is a simple string
            elif hasattr(ElementType, 'CodeBlock') and element.type == ElementType.CodeBlock:
                line_count += (element.content.count('\n') + 1) if element.content else 1
                char_count += len(element.content)
            # Image and Table don't directly contribute to this line count heuristic for font size
        return line_count, char_count

    def _put_elements_on_slide(self, elements: List[SlideElement], is_continued_slide: bool = False):
        """Helper to output a list of elements. `last_title_info` is now an instance var."""
        last_element_type = None
        for element_idx, element in enumerate(elements):
            if last_element_type == ElementType.ListItem and element.type != ElementType.ListItem:
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
                if isinstance(element.content, str):
                     current_content_str = element.content
                elif isinstance(element.content, list): # TextRun
                     current_content_str = self.get_formatted_runs(element.content)


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
                                    effective_title = f'{title_text} (cont.)'
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
                case _ if hasattr(ElementType, 'CodeBlock') and element.type == ElementType.CodeBlock:
                    # Assumption: element.content (str), element.language (Optional[str])
                    code_content = getattr(element, 'content', '')
                    code_lang = getattr(element, 'language', None)
                    self.put_code_block(code_content, code_lang)
            
            last_element_type = element.type
        
        if last_element_type == ElementType.ListItem: # Ensure list footer if slide ends with list
            self.put_list_footer()

    def output(self, presentation_data: ParsedPresentation):
        self.put_header()
        self.last_title_info = None # Reset for each presentation

        num_total_slides = len(presentation_data.slides)
        marp_slide_counter = 0

        for slide_idx, slide in enumerate(presentation_data.slides):
            marp_slide_counter += 1

            all_elements = []
            original_slide_title_text: Optional[str] = None
            original_slide_title_level: Optional[int] = None

            if slide.type == SlideType.General:
                all_elements = slide.elements
            elif slide.type == SlideType.MultiColumn:
                all_elements = slide.preface + [el for col in slide.columns for el in col] # Flatten columns for now

            if not all_elements: # Skip empty slides
                 if marp_slide_counter < num_total_slides : # Check if it's not the last conceptual slide
                    self.ofile.write("\n---\n\n")
                 continue


            # Try to get the main title of the original slide for "(Continued)" logic
            first_title_el = next((el for el in all_elements if el.type == ElementType.Title), None)
            if first_title_el:
                original_slide_title_text = first_title_el.content.strip() if isinstance(first_title_el.content, str) else ""
                original_slide_title_level = first_title_el.level


            line_count, _ = self._get_slide_content_metrics(all_elements)

            def get_slide_class(lc: int) -> Optional[str]:
                if lc > LINES_SMALLER_MAX: return "smallest"
                if lc > LINES_SMALL_MAX: return "smaller"
                if lc > LINES_NORMAL_MAX: return "small"
                return None

            # Splitting logic
            if line_count > LINES_SPLIT_TRIGGER and len(all_elements) > 1: # Split if too many lines and multiple elements
                split_at_index = len(all_elements) // 2
                # Try to split after a paragraph or title, not in middle of list items if possible (simple split for now)
                # A more sophisticated split would find a natural break.
                
                elements_part1 = all_elements[:split_at_index]
                elements_part2 = all_elements[split_at_index:]

                if not elements_part1: # safety if all_elements had only 1 item but line_count was high
                    elements_part1 = elements_part2
                    elements_part2 = []


                # Part 1
                part1_line_count, _ = self._get_slide_content_metrics(elements_part1)
                slide_class_part1 = get_slide_class(part1_line_count)
                if slide_class_part1:
                    self.ofile.write(f"<!-- _class: {slide_class_part1} -->\n\n")
                self._put_elements_on_slide(elements_part1, is_continued_slide=False)
                
                # Notes are typically for the whole original slide. Output with the first part.
                if not self.config.disable_notes and slide.notes:
                    self.ofile.write("<!--\n")
                    for note_line in slide.notes:
                        self.ofile.write(f"{note_line}\n")
                    self.ofile.write("-->\n\n")

                self.ofile.write("\n---\n\n") # Marp slide separator for the new slide
                marp_slide_counter +=1

                # Part 2
                part2_line_count, _ = self._get_slide_content_metrics(elements_part2)
                slide_class_part2 = get_slide_class(part2_line_count)
                if slide_class_part2:
                    self.ofile.write(f"<!-- _class: {slide_class_part2} -->\n\n")
                
                if original_slide_title_text and original_slide_title_level:
                    continued_title = f"{original_slide_title_text} (Continued)"
                    self.put_title(continued_title, original_slide_title_level)
                    self.last_title_info = (continued_title, original_slide_title_level) 
                
                if elements_part2:
                    # If the first element of part2 is the original title, _put_elements_on_slide might skip it.
                    # This needs careful handling. Let's assume _put_elements_on_slide will render content
                    # other than the main continued title which we just put.
                    # A quick fix: if elements_part2[0] was the original_slide_title_el, pass elements_part2[1:]
                    if elements_part2 and first_title_el and elements_part2[0] == first_title_el and original_slide_title_text:
                         self._put_elements_on_slide(elements_part2[1:], is_continued_slide=True)
                    else:
                         self._put_elements_on_slide(elements_part2, is_continued_slide=True)

            else: # No splitting
                current_slide_class = get_slide_class(line_count)
                if current_slide_class:
                    self.ofile.write(f"<!-- _class: {current_slide_class} -->\n\n")
                self._put_elements_on_slide(all_elements, is_continued_slide=False)

                if not self.config.disable_notes and slide.notes:
                    self.ofile.write("<!--\n")
                    for note_line in slide.notes:
                        self.ofile.write(f"{note_line}\n")
                    self.ofile.write("-->\n\n")

            # Add slide separator if not the very last conceptual slide
            # This needs to be smarter if num_total_slides is for original slides, and we are splitting.
            # For now, assume marp_slide_counter reflects actual Marp slides being output.
            # The check `slide_idx < len(presentation_data.slides) - 1` is for original slides.
            # A better check is needed if the last original slide was split.
            # Let's always put one, and rely on Marp to ignore last one if empty.
            # Or, only if there are more original slides OR if current slide was split AND it was the last original one.
            is_last_original_slide = (slide_idx == num_total_slides - 1)
            was_split = (line_count > LINES_SPLIT_TRIGGER and len(all_elements) > 1)

            if not (is_last_original_slide and not was_split) : # Add --- if not the true end
                 self.ofile.write("\n---\n\n")


        self.close()

    def put_title(self, text, level):
        self.ofile.write('#' * level + ' ' + text + '\n\n')

    def put_list(self, text, level):
        self.ofile.write('  ' * level + '* ' + text.strip() + '\n')

    def put_para(self, text):
        self.ofile.write(text + '\n\n')

    def put_image(self, element: ImageElement):
        alt = element.alt_text if element.alt_text else ""
        quoted_path = urllib.parse.quote(element.path)
        
        # Keywords for Marp alt text, including w:, h:, bg, position etc.
        marp_alt_text_keywords = []
        
        original_slide_width_px = self.config.slide_width_px or DEFAULT_SLIDE_WIDTH_PX
        original_slide_height_px = self.config.slide_height_px or DEFAULT_SLIDE_HEIGHT_PX

        ppt_display_width = element.display_width_px
        ppt_display_height = element.display_height_px

        if ppt_display_width is None and self.config.image_width is not None:
            ppt_display_width = self.config.image_width
            if element.original_width_px and element.original_height_px and element.original_width_px > 0:
                 aspect_ratio = element.original_height_px / element.original_width_px
                 ppt_display_height = int(round(ppt_display_width * aspect_ratio))

        scaled_marp_display_width = None
        scaled_marp_display_height = None

        if ppt_display_width is not None and original_slide_width_px > 0:
            width_scale_factor = MARP_TARGET_WIDTH_PX / original_slide_width_px
            scaled_marp_display_width = int(round(ppt_display_width * width_scale_factor))

            if element.original_width_px and element.original_height_px and \
               element.original_width_px > 0 and scaled_marp_display_width > 0:
                image_aspect_ratio = element.original_height_px / element.original_width_px
                scaled_marp_display_height = int(round(scaled_marp_display_width * image_aspect_ratio))
            elif ppt_display_height is not None:
                scaled_marp_display_height = int(round(ppt_display_height * width_scale_factor))
        elif ppt_display_height is not None and original_slide_height_px > 0 and \
             element.original_width_px and element.original_height_px and element.original_height_px > 0 :
            height_scale_factor = MARP_TARGET_HEIGHT_PX / original_slide_height_px
            scaled_marp_display_height = int(round(ppt_display_height * height_scale_factor))
            if element.original_width_px > 0 and element.original_height_px > 0 : 
                image_aspect_ratio_inv = element.original_width_px / element.original_height_px
                scaled_marp_display_width = int(round(scaled_marp_display_height * image_aspect_ratio_inv))

        current_display_width = scaled_marp_display_width
        current_display_height = scaled_marp_display_height
        
        if current_display_width is not None and current_display_width > 0:
            marp_alt_text_keywords.append(f'w:{current_display_width}px') 
        if current_display_height is not None and current_display_height > 0:
            marp_alt_text_keywords.append(f'h:{current_display_height}px')

        # Rotation is not applied as we are not using HTML img tags.
        # If element.rotation is significant, it's visually lost.

        slide_width_for_hinting = MARP_TARGET_WIDTH_PX
        position_hint = None
        
        scaled_left_px = None
        if element.left_px is not None and original_slide_width_px > 0:
            scaled_left_px = int(round(element.left_px * (MARP_TARGET_WIDTH_PX / original_slide_width_px)))

        if scaled_left_px is not None and current_display_width is not None:
            image_center_x = scaled_left_px + (current_display_width / 2)
            slide_center_x = slide_width_for_hinting / 2
            center_threshold = slide_width_for_hinting * 0.10 
            
            left_third_boundary = slide_width_for_hinting / 3
            right_third_boundary = 2 * slide_width_for_hinting / 3

            if abs(image_center_x - slide_center_x) < center_threshold:
                position_hint = "center" 
            elif (scaled_left_px + current_display_width) < left_third_boundary + center_threshold : 
                position_hint = "left"
            elif scaled_left_px > right_third_boundary - center_threshold : 
                position_hint = "right"

        is_background_candidate = False
        if current_display_width is not None and current_display_height is not None:
            if current_display_width >= slide_width_for_hinting * 0.75 and \
               current_display_height >= MARP_TARGET_HEIGHT_PX * 0.75:
                is_background_candidate = True
                if position_hint == "left":
                    position_hint = "bg left"
                elif position_hint == "right":
                    position_hint = "bg right"
                else: 
                    position_hint = "bg" 

        effective_position_hint = position_hint or getattr(element, 'position_hint', None)
        
        has_bg_keyword = False
        if effective_position_hint:
            if effective_position_hint == "center":
                marp_alt_text_keywords.append("center") 
            elif effective_position_hint == "left":
                 marp_alt_text_keywords.append("left") # For CSS: img[alt~="left"]
            elif effective_position_hint == "right":
                 marp_alt_text_keywords.append("right") # For CSS: img[alt~="right"]
            elif effective_position_hint.startswith("bg"):
                bg_directive_parts = effective_position_hint.split(" ") 
                marp_alt_text_keywords.extend(bg_directive_parts)
                has_bg_keyword = True
                # Remove w: and h: if it's a background image, Marp handles sizing.
                marp_alt_text_keywords = [kw for kw in marp_alt_text_keywords if not (kw.startswith("w:") or kw.startswith("h:"))]
        
        # Construct final alt text string for Marp
        # Order: bg/positioning keywords, then original alt text, then sizing keywords.
        
        ordered_alt_keywords = []
        # Add specific Marp keywords first (bg, positioning)
        if "bg" in marp_alt_text_keywords: ordered_alt_keywords.append("bg")
        if "bg left" in marp_alt_text_keywords: ordered_alt_keywords = ["bg", "left"] # replace if more specific
        if "bg right" in marp_alt_text_keywords: ordered_alt_keywords = ["bg", "right"] # replace if more specific
        
        # Add "center", "left", "right" for non-bg images if present
        if not has_bg_keyword:
            if "center" in marp_alt_text_keywords: ordered_alt_keywords.append("center")
            if "left" in marp_alt_text_keywords: ordered_alt_keywords.append("left")
            if "right" in marp_alt_text_keywords: ordered_alt_keywords.append("right")

        # Add the original alt text
        if alt:
            ordered_alt_keywords.append(alt)
            
        # Add sizing keywords (w:, h:), unless it's a background image
        if not has_bg_keyword:
            for kw in marp_alt_text_keywords:
                if kw.startswith("w:") or kw.startswith("h:"):
                    ordered_alt_keywords.append(kw)
        
        final_marp_alt_string = " ".join(ordered_alt_keywords).strip()

        # Cropping is not supported with pure Markdown output.
        # is_cropped = (element.crop_left_pct or ... )
        # if is_cropped:
        #    logger.warning("Image cropping is present but not supported for pure Markdown Marp output. Image will be uncropped.")

        # Always output Marp Markdown image syntax
        self.ofile.write(f'![{final_marp_alt_string}]({quoted_path})\n\n')

    def put_code_block(self, code: str, language: Optional[str]):
        lang_tag = language if language else ""
        self.ofile.write(f'```{lang_tag}\n{code.strip()}\n```\n\n')

    def put_table(self, table):
        gen_table_row = lambda row: '| ' + ' | '.join([c.replace('\n', '<br />') for c in row]) + ' |'
        self.ofile.write(gen_table_row(table[0]) + '\n')
        self.ofile.write(gen_table_row([':-:' for _ in table[0]]) + '\n')
        self.ofile.write('\n'.join([gen_table_row(row) for row in table[1:]]) + '\n\n')

    def get_accent(self, text): # Italics
        return '*' + text.strip() + '*' 

    def get_strong(self, text): # Bold
        return '**' + text.strip() + '**' 

    def get_colored(self, text, rgb):
        # Standard HTML for color, Marp should support it
        return '<span style="color:%s">%s</span>' % (rgb_to_hex(rgb), text)

    def get_hyperlink(self, text, url):
        return '[' + text + '](' + url + ')'

    def esc_repl(self, match):
        return '\\' + match.group(0)

    def get_escaped(self, text):
        # Basic Markdown escaping
        text = re.sub(self.esc_re1, self.esc_repl, text)
        # text = re.sub(self.esc_re2, self.esc_repl, text) # Keep commented for Marp
        return text
