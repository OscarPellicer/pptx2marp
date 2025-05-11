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

import logging

import pptx2md.outputter as outputter
from pptx2md.parser import parse
from pptx2md.types import ConversionConfig
from pptx2md.utils import load_pptx, prepare_titles, emu_to_px

logger = logging.getLogger(__name__)


def convert(config: ConversionConfig):
    if config.title_path:
        config.custom_titles = prepare_titles(config.title_path)

    prs = load_pptx(config.pptx_path)

    # Extract and store actual slide dimensions in config
    if hasattr(prs, 'slide_width') and prs.slide_width is not None:
        config.slide_width_px = emu_to_px(prs.slide_width)
    if hasattr(prs, 'slide_height') and prs.slide_height is not None:
        config.slide_height_px = emu_to_px(prs.slide_height)

    logger.info("conversion started")
    logger.info(f"Detected slide dimensions: {config.slide_width_px}px width, {config.slide_height_px}px height.")

    ast = parse(config, prs)

    if str(config.output_path).endswith('.json'):
        with open(config.output_path, 'w') as f:
            f.write(ast.model_dump_json(indent=2))
        logger.info(f'presentation data saved to {config.output_path}')
        return

    if config.is_wiki:
        out = outputter.WikiFormatter(config)
    elif config.is_mdk:
        out = outputter.MadokoFormatter(config)
    elif config.is_qmd:
        out = outputter.QuartoFormatter(config)
    elif config.is_marp:
        out = outputter.MarpFormatter(config)
    else:
        out = outputter.MarkdownFormatter(config)

    out.output(ast)
    logger.info(f'converted document saved to {config.output_path}')
