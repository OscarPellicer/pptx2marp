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

import argparse
import logging
from pathlib import Path
from dataclasses import replace

from pptx2md.entry import convert
from pptx2md.log import setup_logging
from pptx2md.types import ConversionConfig

setup_logging(compat_tqdm=True)
logger = logging.getLogger(__name__)


def parse_args() -> ConversionConfig:
    arg_parser = argparse.ArgumentParser(description='Convert pptx to markdown')
    arg_parser.add_argument('pptx_path', type=Path, help='Path to the pptx file to be converted')
    arg_parser.add_argument('-o', '--output-dir', type=Path, default='outputs', help='Output directory. Default: {default}')
    arg_parser.add_argument('-i', '--image-dir', type=Path, default=None, help='Directory to put images extracted')
    arg_parser.add_argument('-t', '--title', type=Path, help='Path to the custom title list file')
    arg_parser.add_argument('--image-width', type=int, help='Maximum image width in px')
    arg_parser.add_argument('--disable-image', action="store_true", help='Disable image extraction')
    arg_parser.add_argument('--disable-wmf',
                            action="store_true",
                            help='  Keep wmf formatted image untouched (avoid exceptions under linux)')
    arg_parser.add_argument('--disable-color', action="store_true", help='Do not add color HTML tags')
    arg_parser.add_argument('--disable-escaping',
                            action="store_true",
                            help='Do not attempt to escape special characters')
    arg_parser.add_argument('--disable-notes', action="store_true", help='Do not add presenter notes')
    arg_parser.add_argument('--enable-slides', action="store_true", help='Deliniate slides `\n---\n`')
    arg_parser.add_argument('--try-multi-column', action="store_true", help='Try to detect multi-column slides')
    arg_parser.add_argument('--md', action="store_true", help='Generate output as standard markdown')
    arg_parser.add_argument('--wiki', action="store_true", help='Generate output as wikitext (TiddlyWiki)')
    arg_parser.add_argument('--mdk', action="store_true", help='Generate output as madoko markdown')
    arg_parser.add_argument('--qmd', action="store_true", help='Generate output as quarto markdown presentation')
    arg_parser.add_argument('--marp', action="store_true", help='Generate output as marp markdown presentation')
    arg_parser.add_argument('--beamer', action="store_true", help='Generate output as LaTeX Beamer presentation')
    arg_parser.add_argument('--json', action="store_true", help='Generate output as the raw .pptx abstract syntax tree in JSON format')
    arg_parser.add_argument('--min-block-size',
                            type=int,
                            default=0,
                            help='Minimum character number of a text block to be converted')
    arg_parser.add_argument("--page", type=int, default=None, help="Only convert the specified page")
    arg_parser.add_argument(
        "--keep-similar-titles",
        action="store_true",
        help="Keep similar titles (allow for repeated slide titles - One or more - Add (cont.) to the title)")
    arg_parser.add_argument(
        '--disable-parser-cropping',
        action="store_false", 
        dest='apply_cropping_in_parser',
        help='Disable pre-cropping of images in the parser. Crop information (if available) will be passed to the formatter.'
    )
    arg_parser.set_defaults(apply_cropping_in_parser=True)

    args = arg_parser.parse_args()

    return ConversionConfig(
        pptx_path=args.pptx_path,
        output_path=None, #This will be automatically set in entry.py
        output_dir=args.output_dir,
        image_dir=None, #This will be automatically set in entry.py
        title_path=args.title,
        image_width=args.image_width,
        disable_image=args.disable_image,
        disable_wmf=args.disable_wmf,
        disable_color=args.disable_color,
        disable_escaping=args.disable_escaping,
        disable_notes=args.disable_notes,
        enable_slides=args.enable_slides,
        try_multi_column=args.try_multi_column,
        is_md=args.md,
        is_wiki=args.wiki,
        is_mdk=args.mdk,
        is_qmd=args.qmd,
        is_marp=args.marp,
        is_beamer=args.beamer,
        is_json=args.json,
        min_block_size=args.min_block_size,
        page=args.page,
        keep_similar_titles=args.keep_similar_titles,
        apply_cropping_in_parser=args.apply_cropping_in_parser,
    )


def main():
    config = parse_args()
    
    if config.pptx_path.is_dir():
        pptx_files = list(config.pptx_path.glob("*.pptx"))
        for pptx_file in pptx_files:
            file_config = replace(config, pptx_path=pptx_file)
            convert(file_config)
    else:
        convert(config)

if __name__ == '__main__':
    main()
