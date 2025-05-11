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

from __future__ import print_function

import logging
import os
from functools import partial
from operator import attrgetter
from typing import List, Union
from pathlib import Path
import io # Add io import

from PIL import Image
from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from rapidfuzz import process as fuze_process
from tqdm import tqdm
from pptx.util import Emu # For EMU to Px conversion

from pptx2md.multi_column import get_multi_column_slide_if_present
from pptx2md.types import (
    ConversionConfig,
    GeneralSlide,
    ImageElement,
    ListItemElement,
    ParagraphElement,
    ParsedPresentation,
    SlideElement,
    TableElement,
    TextRun,
    TextStyle,
    TitleElement,
)
from pptx2md.utils import emu_to_px # Assuming emu_to_px is now in utils

logger = logging.getLogger(__name__)

picture_count = 0


def is_title(shape):
    if shape.is_placeholder and (shape.placeholder_format.type == PP_PLACEHOLDER.TITLE or
                                 shape.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE or
                                 shape.placeholder_format.type == PP_PLACEHOLDER.VERTICAL_TITLE or
                                 shape.placeholder_format.type == PP_PLACEHOLDER.CENTER_TITLE):
        return True
    return False


def is_text_block(config: ConversionConfig, shape):
    if shape.has_text_frame:
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.BODY:
            return True
        if len(shape.text) > config.min_block_size:
            return True
    return False


def is_list_block(shape) -> bool:
    levels = []
    for para in shape.text_frame.paragraphs:
        if para.level not in levels:
            levels.append(para.level)
        if para.level != 0 or len(levels) > 1:
            return True
    return False


def is_accent(font):
    if font.underline or font.italic or (
            font.color.type == MSO_COLOR_TYPE.SCHEME and
        (font.color.theme_color == MSO_THEME_COLOR.ACCENT_1 or font.color.theme_color == MSO_THEME_COLOR.ACCENT_2 or
         font.color.theme_color == MSO_THEME_COLOR.ACCENT_3 or font.color.theme_color == MSO_THEME_COLOR.ACCENT_4 or
         font.color.theme_color == MSO_THEME_COLOR.ACCENT_5 or font.color.theme_color == MSO_THEME_COLOR.ACCENT_6)):
        return True
    return False


def is_strong(font):
    if font.bold or (font.color.type == MSO_COLOR_TYPE.SCHEME and (font.color.theme_color == MSO_THEME_COLOR.DARK_1 or
                                                                   font.color.theme_color == MSO_THEME_COLOR.DARK_2)):
        return True
    return False


def get_text_runs(para) -> List[TextRun]:
    runs = []
    for run in para.runs:
        result = TextRun(text=run.text, style=TextStyle())
        if result.text == '':
            continue
        try:
            if run.hyperlink.address:
                result.style.hyperlink = run.hyperlink.address
        except:
            result.style.hyperlink = 'error:ppt-link-parsing-issue'
        if is_accent(run.font):
            result.style.is_accent = True
        if is_strong(run.font):
            result.style.is_strong = True
        if run.font.color.type == MSO_COLOR_TYPE.RGB:
            result.style.color_rgb = run.font.color.rgb
        runs.append(result)
    return runs


def process_title(config: ConversionConfig, shape, slide_idx) -> TitleElement:
    text = shape.text_frame.text.strip()
    if config.custom_titles:
        res = fuze_process.extractOne(text, config.custom_titles.keys(), score_cutoff=92)
        if not res:
            return TitleElement(content=text.strip(), level=max(config.custom_titles.values()) + 1)
        else:
            logger.info(f'Title in slide {slide_idx} "{text}" is converted to "{res[0]}" as specified in title file.')
            return TitleElement(content=res[0].strip(), level=config.custom_titles[res[0]])
    else:
        return TitleElement(content=text.strip(), level=1)


def process_text_blocks(config: ConversionConfig, shape, slide_idx) -> List[Union[ListItemElement, ParagraphElement]]:
    results = []
    if is_list_block(shape):
        for para in shape.text_frame.paragraphs:
            if para.text.strip() == '':
                continue
            text = get_text_runs(para)
            results.append(ListItemElement(content=text, level=para.level))
    else:
        # paragraph block
        for para in shape.text_frame.paragraphs:
            if para.text.strip() == '':
                continue
            text = get_text_runs(para)
            results.append(ParagraphElement(content=text))
    return results


def process_picture(config: ConversionConfig, shape, slide_idx) -> Union[ImageElement, None]:
    if config.disable_image:
        return None

    if not hasattr(shape, 'image') or not shape.image:
        logger.warning(f"Shape in slide {slide_idx} seems to be a picture but has no image data, skipped.")
        return None

    global picture_count

    # Initial properties from shape.image
    original_pic_ext = shape.image.ext.lower()
    current_image_blob = shape.image.blob
    # Pillow uses 'JPEG' for '.jpg', so map common extensions
    pil_format_map = {'jpg': 'JPEG', 'jpeg': 'JPEG', 'tif': 'TIFF', 'tiff': 'TIFF'}
    current_pil_format = pil_format_map.get(original_pic_ext, original_pic_ext.upper())
    
    # --- WMF Conversion (if applicable, happens before cropping) ---
    converted_from_wmf = False
    if original_pic_ext == 'wmf':
        if config.disable_wmf: # User wants to keep WMF as is, no conversion, no cropping by Pillow
            # Save WMF as is, ImageElement will reflect this
            # The rest of the logic will treat it like any other image but Pillow won't process it
            # This means no pre-cropping by this script.
            pass
        else:
            try:
                # Create a temporary path for Wand to read the WMF blob from
                # as Wand typically works with filenames.
                temp_wmf_path = Path(config.image_dir) / f"__temp_wmf_{picture_count}.{original_pic_ext}"
                with open(temp_wmf_path, 'wb') as tmp_f:
                    tmp_f.write(current_image_blob)

                from wand.image import Image as WandImage
                with WandImage(filename=str(temp_wmf_path)) as img:
                    img.format = 'png'
                    with io.BytesIO() as png_blob_io:
                        img.save(file=png_blob_io)
                        current_image_blob = png_blob_io.getvalue()
                current_pil_format = 'PNG'
                original_pic_ext = 'png' # Update extension for saving
                converted_from_wmf = True
                logger.info(f'WMF image in slide {slide_idx} converted to PNG for processing.')
                if temp_wmf_path.exists(): # Clean up temp file
                    try:
                        os.remove(temp_wmf_path)
                    except OSError as e:
                        logger.warning(f"Could not remove temp WMF file {temp_wmf_path}: {e}")

            except Exception as e:
                logger.warning(
                    f'Cannot convert wmf image in slide {slide_idx} to png for cropping, attempting to save as original. Error: {e}')
                # Fallback: treat as uncroppable by this script if WMF conversion failed
                # The original blob and ext will be used.
                # If disable_wmf was false, this means we tried and failed.


    # --- Image Cropping with Pillow (on original or WMF-converted blob) ---
    img_to_process = None
    needs_saving_after_pillow = False # Flag if Pillow processing happens

    if not (original_pic_ext == 'wmf' and config.disable_wmf): # Don't process WMF with Pillow if disabled
        try:
            img_to_process = Image.open(io.BytesIO(current_image_blob))
            # Ensure image is in a mode that supports saving to its target format (especially for transparency)
            if current_pil_format == 'PNG' and img_to_process.mode != 'RGBA':
                img_to_process = img_to_process.convert('RGBA')
            elif img_to_process.mode == 'P': # Palette mode, convert to RGB/RGBA
                 img_to_process = img_to_process.convert('RGBA' if 'A' in img_to_process.mode else 'RGB')


            needs_saving_after_pillow = True # Will need re-saving if opened by Pillow
        except Exception as e:
            logger.warning(f"Pillow could not open image blob for slide {slide_idx} (ext: {original_pic_ext}). Error: {e}. Skipping Pillow processing.")
            img_to_process = None # Cannot process further with Pillow
            needs_saving_after_pillow = False


    # These will be the dimensions of the image data *after* Pillow processing.
    # If Pillow fails or is skipped, they'll be from shape.image.size.
    
    # Initialize final dimensions to be those of the image blob *before* potential cropping
    if img_to_process:
        final_blob_w_px, final_blob_h_px = img_to_process.size
    elif hasattr(shape.image, 'size') and shape.image.size:
        final_blob_w_px, final_blob_h_px = shape.image.size
    else:
        final_blob_w_px, final_blob_h_px = None, None

    # Initialize crop percentages to be passed to ImageElement if not applied
    crop_l_for_element, crop_r_for_element, crop_t_for_element, crop_b_for_element = None, None, None, None

    if img_to_process:
        pil_original_w, pil_original_h = img_to_process.size # Dimensions before any Pillow crop

        crop_l_pct, crop_r_pct, crop_t_pct, crop_b_pct = 0.0, 0.0, 0.0, 0.0
        has_crop_info = False
        if hasattr(shape, 'pic') and shape.pic is not None:
            if hasattr(shape.pic, 'crop_left') and shape.pic.crop_left > 0.00001: # Use small epsilon
                 crop_l_pct = shape.pic.crop_left; has_crop_info = True
            if hasattr(shape.pic, 'crop_right') and shape.pic.crop_right > 0.00001:
                 crop_r_pct = shape.pic.crop_right; has_crop_info = True
            if hasattr(shape.pic, 'crop_top') and shape.pic.crop_top > 0.00001:
                 crop_t_pct = shape.pic.crop_top; has_crop_info = True
            if hasattr(shape.pic, 'crop_bottom') and shape.pic.crop_bottom > 0.00001:
                 crop_b_pct = shape.pic.crop_bottom; has_crop_info = True

        if has_crop_info and config.apply_cropping_in_parser:
            left = int(round(pil_original_w * crop_l_pct))
            top = int(round(pil_original_h * crop_t_pct))
            right = int(round(pil_original_w * (1.0 - crop_r_pct)))
            bottom = int(round(pil_original_h * (1.0 - crop_b_pct)))

            if left < right and top < bottom:
                try:
                    img_to_process = img_to_process.crop((left, top, right, bottom))
                    logger.info(f'Image in slide {slide_idx} pre-cropped in parser. New blob dims: {img_to_process.size}')
                    with io.BytesIO() as cropped_blob_io:
                        save_format = current_pil_format if current_pil_format else 'PNG'
                        try:
                            img_to_process.save(cropped_blob_io, format=save_format)
                        except KeyError: 
                            logger.warning(f"Format {save_format} not supported by Pillow for saving, falling back to PNG.")
                            save_format = 'PNG'
                            if img_to_process.mode != 'RGBA' and img_to_process.mode != 'RGB':
                                img_to_process = img_to_process.convert('RGBA')
                            img_to_process.save(cropped_blob_io, format=save_format)
                        current_image_blob = cropped_blob_io.getvalue()
                    
                    final_blob_w_px, final_blob_h_px = img_to_process.size
                    # Cropping applied, so no percentages needed for ImageElement
                    crop_l_for_element, crop_r_for_element, crop_t_for_element, crop_b_for_element = None, None, None, None
                except Exception as e:
                    logger.warning(f"Failed to apply crop in parser for image in slide {slide_idx}. Error: {e}")
                    # Fallback: use uncropped dimensions, and pass crop info
                    final_blob_w_px, final_blob_h_px = pil_original_w, pil_original_h
                    if has_crop_info: # Still pass the crop info if it existed
                        crop_l_for_element, crop_r_for_element, crop_t_for_element, crop_b_for_element = \
                            crop_l_pct, crop_r_pct, crop_t_pct, crop_b_pct
            else:
                logger.warning(f"Invalid crop dimensions for image in slide {slide_idx}, not applying crop in parser.")
                final_blob_w_px, final_blob_h_px = pil_original_w, pil_original_h
                if has_crop_info: # Pass the crop info
                    crop_l_for_element, crop_r_for_element, crop_t_for_element, crop_b_for_element = \
                        crop_l_pct, crop_r_pct, crop_t_pct, crop_b_pct
        elif has_crop_info: # Cropping info exists, but apply_cropping_in_parser is False
            logger.info(f"Parser cropping disabled for image in slide {slide_idx}. Crop info will be passed to formatter.")
            final_blob_w_px, final_blob_h_px = pil_original_w, pil_original_h # Dimensions of uncropped blob
            crop_l_for_element, crop_r_for_element, crop_t_for_element, crop_b_for_element = \
                crop_l_pct, crop_r_pct, crop_t_pct, crop_b_pct
        else: # No crop defined on shape.pic, or Pillow couldn't open
            # final_blob_w_px, final_blob_h_px already set from img_to_process.size or shape.image.size
            crop_l_for_element, crop_r_for_element, crop_t_for_element, crop_b_for_element = None, None, None, None
            
    # If Pillow processing was skipped entirely (e.g. disabled WMF)
    # final_blob_w_px, final_blob_h_px would have been set from shape.image.size initially.
    # Ensure they are not None if possible.
    if final_blob_w_px is None and hasattr(shape.image, 'size') and shape.image.size:
         final_blob_w_px, final_blob_h_px = shape.image.size


    # --- Saving the final blob ---
    file_prefix = ''.join(os.path.basename(config.pptx_path).split('.')[:-1])
    # Use original_pic_ext which is now 'png' if WMF was converted and processed
    pic_name_for_save = file_prefix + f'_{picture_count}'
    
    # Ensure image directory exists
    img_dir_path_obj = Path(config.image_dir)
    if not img_dir_path_obj.exists():
        img_dir_path_obj.mkdir(parents=True, exist_ok=True)
    
    # Final output path uses original_pic_ext (which might have been updated from wmf to png)
    output_path_obj = img_dir_path_obj / f'{pic_name_for_save}.{original_pic_ext}'

    with open(output_path_obj, 'wb') as f:
        f.write(current_image_blob)
    picture_count += 1


    # Determine relative path for ImageElement
    config_output_path_obj = Path(config.output_path)
    try:
        base_for_relpath = config_output_path_obj.parent 
        img_outputter_path = os.path.relpath(output_path_obj, base_for_relpath)
    except ValueError: 
        img_outputter_path = str(output_path_obj.resolve())
    
    saved_path_str = str(img_outputter_path).replace('\\', '/')

    # ImageElement properties
    # original_w/h_px are now the dimensions of the (potentially cropped) saved file
    # display_w/h_px are how this saved file is framed on the slide
    image_data = ImageElement(
        path=saved_path_str,
        original_width_px=final_blob_w_px,
        original_height_px=final_blob_h_px,
        original_filename=shape.image.filename if hasattr(shape.image, 'filename') else None, # Original filename from PPTX
        display_width_px=emu_to_px(shape.width),  # Size of the frame on slide
        display_height_px=emu_to_px(shape.height), # Size of the frame on slide
        left_px=emu_to_px(shape.left),
        top_px=emu_to_px(shape.top),
        rotation=shape.rotation if hasattr(shape, 'rotation') else 0.0,
        crop_left_pct=crop_l_for_element, 
        crop_right_pct=crop_r_for_element,
        crop_top_pct=crop_t_for_element,
        crop_bottom_pct=crop_b_for_element,
        alt_text=shape.name if hasattr(shape, 'name') else None
    )
    return image_data


def process_table(config: ConversionConfig, shape, slide_idx) -> Union[TableElement, None]:
    table = [[sum([get_text_runs(p)
                   for p in cell.text_frame.paragraphs], [])
              for cell in row.cells]
             for row in shape.table.rows]
    if len(table) > 0:
        return TableElement(content=table)
    return None


def ungroup_shapes(shapes) -> List[SlideElement]:
    res = []
    for shape in shapes:
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                res.extend(ungroup_shapes(shape.shapes))
            else:
                res.append(shape)
        except Exception as e:
            logger.warning(f'failed to load shape {shape}, skipped. error: {e}')
    return res


def process_shapes(config: ConversionConfig, current_shapes, slide_id: int) -> List[SlideElement]:
    results = []
    for shape in current_shapes:
        if is_title(shape):
            results.append(process_title(config, shape, slide_id))
        elif is_text_block(config, shape):
            results.extend(process_text_blocks(config, shape, slide_id))
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                pic = process_picture(config, shape, slide_id)
                if pic:
                    results.append(pic)
            except AttributeError as e:
                logger.warning(f'Failed to process picture, skipped: {e}')
        elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            table = process_table(config, shape, slide_id)
            if table:
                results.append(table)
        else:
            try:
                ph = shape.placeholder_format
                if ph.type == PP_PLACEHOLDER.OBJECT and hasattr(shape, "image") and getattr(shape, "image"):
                    pic = process_picture(config, shape, slide_id)
                    if pic:
                        results.append(pic)
            except:
                pass

    return results


def parse(config: ConversionConfig, prs: Presentation) -> ParsedPresentation:
    result = ParsedPresentation(slides=[])

    for idx, slide in enumerate(tqdm(prs.slides, desc='Converting slides')):
        if config.page is not None and idx + 1 != config.page:
            continue
        shapes = []
        try:
            shapes = sorted(ungroup_shapes(slide.shapes), key=attrgetter('top', 'left'))
        except:
            logger.warning('Bad shapes encountered in this slide. Please check or remove them and try again.')
            logger.warning('shapes:')
            try:
                for sp in slide.shapes:
                    logger.warning(sp.shape_type)
                    logger.warning(sp.top, sp.left, sp.width, sp.height)
            except:
                logger.warning('failed to print all bad shapes.')

        if not config.try_multi_column:
            result_slide = GeneralSlide(elements=process_shapes(config, shapes, idx + 1))
        else:
            multi_column_slide = get_multi_column_slide_if_present(
                prs, slide, partial(process_shapes, config=config, slide_id=idx + 1))
            if multi_column_slide:
                result_slide = multi_column_slide
            else:
                result_slide = GeneralSlide(elements=process_shapes(config, shapes, idx + 1))

        if not config.disable_notes and slide.has_notes_slide:
            text = slide.notes_slide.notes_text_frame.text
            if text:
                result_slide.notes.append(text)

        result.slides.append(result_slide)

    return result
