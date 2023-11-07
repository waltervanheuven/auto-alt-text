"""
Generate Alt Text for each picture in a powerpoint file using MLLM and V-L pre-trained models
"""

from typing import List
import os
import sys
import io
import argparse
import base64
import csv
import re
import requests
from PIL import Image
from pptx.oxml.ns import _nsmap
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.base import BaseShape
import open_clip
import torch
from transformers import AutoProcessor, AutoModelForVision2Seq

def check_server_is_running(url: str) -> bool:
    """ URL accessible? """    
    status: bool = False
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            status = True
    except requests.exceptions.Timeout:
        print("Timeout exception")
    except requests.exceptions.RequestException as e:
        print(f"Exception: {str(e)}")
    return status

def num2str(the_max: int, n:int) -> str:
    """ convert number to string with trailing zeros """
    s: str = f"{str(n)}"
    if the_max > 99:
        if n < 100:
            if n < 10:
                s = f"00{str(n)}"
            else:
                s = f"0{str(n)}"
    elif n < 10:
        s = f"0{str(n)}"
    return s

def bool_value(s: str) -> bool:
    """ convert str True or False to bool """
    assert(s is not None and len(s) > 0)
    return s.lower() == "true"

def bool_to_string(b: bool) -> str:
    """ convert bool to str """
    return "True" if b else "False"

# see https://github.com/scanny/python-pptx/pull/512
def shape_get_alt_text(shape: BaseShape) -> str:
    """ Alt-text defined in shape's `descr` attribute, or '' if not present. """
    return shape._element._nvXxPr.cNvPr.attrib.get("descr", "")

def shape_set_alt_text(shape: BaseShape, alt_text: str):
    """ Set alt-text in shape """
    shape._element._nvXxPr.cNvPr.attrib["descr"] = alt_text

# see https://stackoverflow.com/questions/63802783/check-if-image-is-decorative-in-powerpoint-using-python-pptx
def is_decorative(shape):
    """ check if image is decorative """
    # <adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/>
    _nsmap["adec"] = "http://schemas.microsoft.com/office/drawing/2017/decorative"
    cNvPr = shape._element._nvXxPr.cNvPr
    adec_decoratives = cNvPr.xpath(".//adec:decorative[@val='1']")
    return bool(adec_decoratives)

def process_images_from_pptx(file_path: str, generate: bool, settings: dict, savePP: bool, debug: bool = False) -> bool:
    """
    Loop through images in the slides of a Powerpint file and set image description based 
    on image description from Kosmos-2, OpenCLIP, or LLaVA
    """
    err: bool = False

    # get name, extension, folder from Powerpoint file
    file_name:str = os.path.basename(file_path)
    name:str = file_name.split(".")[0]
    extension:str = file_name.split(".")[1]
    dirname:str = os.path.dirname(file_path)

    # create folder to store images
    img_folder = os.path.join(dirname, name)
    if not os.path.isdir(img_folder):
        os.makedirs(img_folder)

    # Initialize presentation object
    print(f"Reading '{file_path}'")
    prs = Presentation(file_path)

    model_str:str = settings['model']

    # set output file name
    out_file_name:str = ""
    if model_str != "" and generate:
        out_file_name = os.path.join(dirname, f"{name}_{model_str}.txt")
    else:
        out_file_name = os.path.join(dirname, f"{name}.txt")

    nr_slides = len(prs.slides)

    # download and/or set up model
    if generate:
        err = init_model(settings)
        if err:
            print("Unable to init model.")
            return err

    # open file for writing
    with open(out_file_name, "w", encoding="utf-8") as fout:
        # write header
        if model_str != "" and generate:
            fout.write(f"Model\tFile\tSlide\tPicture\tAlt_Text\tDecorative\tPictFilePath{os.linesep}")
        else:
            fout.write(f"File\tSlide\tPicture\tAlt_Text\tDecorative\tPictFilePath{os.linesep}")

        # Loop through slides
        image_cnt:int = 0
        for slide_cnt, slide in enumerate(prs.slides):
            # loop through shapes
            slide_image_cnt = 0
            for shape in slide.shapes:
                 slide_image_cnt = process_shape(shape, None, None, 0, 0, model_str, name, extension, fout, generate, img_folder, slide_cnt, nr_slides, slide_image_cnt, settings, debug)

            image_cnt += slide_image_cnt

    print(f"Powerpoint file contains {slide_cnt + 1} slides and {image_cnt} images.")

    if generate and savePP:
        # Save file
        outfile:str = os.path.join(dirname, f"{name}_alt_text.{extension}")
        print(f"Saving Powerpoint file with new alt-text to {outfile}")
        prs.save(outfile)

    return err

def init_model(settings: dict) -> bool:
    """ download and init model for inference """
    err: bool = False
    model_str:str = settings["model"]
    prompt:str = settings["prompt"]

    if model_str == "kosmos-2":
        # Kosmos-2 model
        model_name:str = "microsoft/kosmos-2-patch14-224"
        print(f"Kosmos-2 model: '{model_name}'")
        print(f"prompt: '{prompt}'")
        settings["kosmos2-model"] = AutoModelForVision2Seq.from_pretrained(model_name)
        settings["kosmos2-processor"] = AutoProcessor.from_pretrained(model_name)
    elif model_str == "openclip":
        # OpenCLIP
        print(f"OpenCLIP model: '{settings['openclip_model_name']}'\npretrained: '{settings['openclip_pretrained']}'")
        model, _, transform = open_clip.create_model_and_transforms(
            model_name=settings["openclip_model_name"],
            pretrained=settings["openclip_pretrained"]
        )
        settings["openclip-model"] = model
        settings["openclip-transform"] = transform
    elif model_str == "llava":
        # LLaVA
        server_url = settings["llava_url"]
        if check_server_is_running(server_url):
            server_url = f"{server_url}/completion"
            print(f"LLaVA server: '{server_url}'")
            print(f"prompt: '{prompt}'")
        else:
            print(f"Unable to access server at '{server_url}'.")
            err = True
    else:
        print(f"Unknown model: '{model_str}'")
        err = True

    return err

def process_shape(shape: BaseShape, images_shape: BaseShape, images, base_left, base_top, model_str: str, name: str , extension:str , fout, generate: bool, img_folder: str,
                    slide_cnt: int, nr_slides: int, slide_image_cnt: int, settings: dict, debug: bool) -> int:
    """
    Recursive function to process shapes and shapes in groups on each slide 
        
    TODO: reduce the number of function arguments
    """

    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        # at top level of group shape create a new image based on images in the group
        #if images == None:
        #    images = []
        #    images_shape = shape

        for embedded_shape in shape.shapes:
            slide_image_cnt = process_shape(embedded_shape, images_shape, images, base_left, base_top, model_str, name, extension, fout, generate, img_folder, slide_cnt, nr_slides, slide_image_cnt, settings, debug)

        #if images:
        #    new_img = combine_images_in_group(images, images_shape)
        #    filename = os.path.join(img_folder, f'slide_{slide_cnt}_group.png')
        #    new_img.save(filename)
        #
        #    # Set Alt Text of group shape based on alt text of new group image
        #    # There is no function in python pptx to set alt text of a group shape
        #
        #    # reset images
        #    images = None

    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        err: bool = False
        image_file_path:str = ""
        decorative:bool = is_decorative(shape)

        # only generate alt text when generate options is True and decorative is False
        if generate and not decorative:
            err, image_file_path = set_alt_text(shape, images, base_left, base_top, img_folder, slide_cnt, nr_slides, slide_image_cnt, settings, debug)

        # report alt text
        if not err:
            stored_alt_text = shape_get_alt_text(shape)
            feedback = f"Slide: {slide_cnt + 1}, Picture: '{shape.name}', alt_text: '{stored_alt_text}', decorative: {bool_to_string(decorative)}"
            print(feedback)

            if model_str == "":
                fout.write(f"{name}.{extension}\t{slide_cnt + 1}\t{shape.name}\t{stored_alt_text}\t{bool_to_string(decorative)}\t{image_file_path}" + os.linesep)
            else:
                fout.write(f"{model_str}\t{name}.{extension}\t{slide_cnt + 1}\t{shape.name}\t{stored_alt_text}\t{bool_to_string(decorative)}\t{image_file_path}" + os.linesep)

            slide_image_cnt += 1
    
    return slide_image_cnt

def combine_images_in_group(images, group_shape):
    """ 
    Create new image based on shape

    TODO: Not yet working properly, image size is not correct
    """

    # EMU per Pixel estimate: not correct
    EMU_PER_PIXEL:int = int(914400 / 96)

    # Determine the size of the new image based on the group shape size
    new_img_width = int(group_shape.width / EMU_PER_PIXEL)
    new_img_height = int(group_shape.height / EMU_PER_PIXEL)
    new_img = Image.new('RGB', (new_img_width, new_img_height))

    # Paste each image into the new image at its relative position
    for image, left, top in images:
        new_img.paste(image, (int(left / EMU_PER_PIXEL), int(top / EMU_PER_PIXEL)))

    return new_img

def set_alt_text(shape: BaseShape, images, base_left, base_top, img_folder: str, slide_cnt: int, max_slides: int, image_cnt: int, settings: dict, debug: bool) -> bool:
    """ 
    Set alt text of image 
    
    TODO: reduce the number of function arguments
    """
    err: bool = False
    image_file_path: str = ""

    # get image
    if hasattr(shape, "image"):
        # get image, works with only with PNG, JPG?
        image_stream = shape.image.blob
        extension:str = shape.image.ext
    else:
        # get image for other formats, e.g. TIF
        # <Element {http://schemas.openxmlformats.org/presentationml/2006/main}pic at 0x15f2d6b20>
        try:
            slide_part = shape.part
            rId = shape._element.blip_rId
            image_part = slide_part.related_part(rId)
            image_stream = image_part.blob
            extension = image_part.partname.ext
        except AttributeError:
            print(f"Slide: {slide_cnt + 1 }, Picture '{shape.name}', unable to access image")
            err = True

    if not err:
        # Keep image in case the image is part of a group
        if images is not None:
            # Calculate the position of the image relative to the group
            image_group_part = Image.open(io.BytesIO(image_stream))
            left = base_left + shape.left
            top = base_top + shape.top
            images.append((image_group_part, left, top))

        # determine file name
        image_file_name:str = f"s{num2str(max_slides, slide_cnt + 1)}p{num2str(99, image_cnt + 1)}_{shape.name}"
        image_file_path = os.path.join(img_folder, f"{image_file_name}.{extension}")
        print(f"Saving and processing image: '{image_file_path}'...")

        # save image
        with open(image_file_path, "wb") as f:
            f.write(image_stream)

        alt_text: str = generate_description(image_file_path, settings)

        if debug:
            print(f"Len: {len(alt_text)}, Content: {alt_text}")

        if len(alt_text) > 0:
            shape_set_alt_text(shape, alt_text)
        else:
            print("Alt text is empty")

    return err, image_file_path

def generate_description(image_file_path: str, settings: dict, debug:bool=False) -> str:
    """ generate image text description using MLLM/VL model """
    alt_text: str = ""

    if settings["model"] == "kosmos-2":
        processor = settings["kosmos2-processor"]
        model = settings["kosmos2-model"]

        # read image
        im = Image.open(image_file_path)

        # prompt
        prompt:str = settings["prompt"]
        #prompt = "<grounding>An image of"
        #prompt = "<grounding> Describe this image in detail:"

        inputs = processor(text=prompt, images=im, return_tensors="pt")

        generated_ids = model.generate(
            pixel_values=inputs["pixel_values"],
            input_ids=inputs["input_ids"],
            attention_mask=inputs["attention_mask"],
            image_embeds=None,
            image_embeds_position_mask=inputs["image_embeds_position_mask"],
            use_cache=True,
            max_new_tokens=128,
        )
        generated_text = processor.batch_decode(generated_ids, skip_special_tokens=True)[0]

        # Specify `cleanup_and_extract=False` in order to see the raw model generation.
        #processed_text = processor.post_process_generation(generated_text, cleanup_and_extract=True)

        # processed_text, entities = processor.post_process_generation(generated_text)
        processed_text, _ = processor.post_process_generation(generated_text)

        # remove prompt
        p = re.sub('<[^<]+?>', '', prompt)
        processed_text = processed_text.replace(p.strip(), '')

        # capitalize
        alt_text = processed_text.strip().capitalize()

    elif settings["model"] == "openclip":
        model = settings["openclip-model"]
        transform = settings["openclip-transform"]

        # read image
        im = Image.open(image_file_path).convert("RGB")
        im = transform(im).unsqueeze(0)

        # use OpenCLIP model to create label
        with torch.no_grad(), torch.cuda.amp.autocast():
            generated = model.generate(im)

        # get picture description and remove trailing spaces
        alt_text = open_clip.decode(generated[0]).split("<end_of_text>")[0].replace("<start_of_text>", "").strip()

        # remove space before '.'
        alt_text = alt_text.replace(' .', '.')
    elif settings["model"] == "llava":
        server_url = f"{settings['llava_url']}/completion"
        prompt = settings["prompt"]

        # read image
        with open(image_file_path, 'rb') as img_file:
            img_byte_arr = img_file.read()

        # encode in base64
        img_base64 = base64.b64encode(img_byte_arr).decode('utf-8')

        # Use LLaVa to get image descriptions
        header = {"Content-Type": "application/json"}
        data = {
            "image_data": [{"data": img_base64, "id": 1}],
            "prompt": f"USER:[img-1] {prompt}\nASSISTANT:",
            "n_predict": 123,
            "temp": 0.1
        }
        try:
            response = requests.post(server_url, headers=header, json=data, timeout=10)
            response_data = response.json()

            if debug:
                print(response_data)
                print()
        except requests.exceptions.Timeout:
            print("Timeout")
        except requests.exceptions.RequestException as e:
            print(f"Exception: {str(e)}")
        else:
            # get picture description and remove trailing spaces
            alt_text = response_data.get('content', '').strip()

            # remove returns
            alt_text = alt_text.replace('\r', '')

    return alt_text

def add_alt_text_from_file(file_path: str, file_path_txt_file: str) -> bool:
    """
    Add alt text specified in a text file (e.g. generated by this script and edited to correct or improve)
    Text file should have a header and the same columns as the output files generated by this script
    """
    err: bool = False

    # Check if text file is exists
    if not os.path.isfile(file_path_txt_file):
        print(f"Unable to access file: {file_path_txt_file}")
        return False

    # get name, extension, folder from Powerpoint file
    file_name:str = os.path.basename(file_path)
    name:str = file_name.split(".")[0]
    extension:str = file_name.split(".")[1]
    dirname:str = os.path.dirname(file_path)

    # process txt file
    print(f"Reading: {file_path_txt_file}...")
    csv_rows = []
    with open(file_path_txt_file, "r", encoding="utf-8") as file:
        # assume tab delimited file
        csv_reader = csv.reader(file, delimiter="\t")

        # skip header
        next(csv_reader)

        for row in csv_reader:
            csv_rows.append(row)

    # process powerpoint file
    print(f"Processing Powerpoint file: {file_path}")
    prs = Presentation(file_path)

    # Loop through slides
    image_cnt:int = 0
    for slide_cnt, slide in enumerate(prs.slides):
        # loop through shapes
        slide_image_cnt = 0
        for shape in slide.shapes:
            image_cnt, slide_image_cnt = process_shapes_from_file(shape, csv_rows, image_cnt, slide_cnt, slide_image_cnt)

    if not err:
        # Save file
        outfile:str = os.path.join(dirname, f"{name}_alt_text.{extension}")
        print(f"Saving Powerpoint file with new alt-text to {outfile}")
        prs.save(outfile)

    return err

def process_shapes_from_file(shape: BaseShape, csv_rows, image_cnt: int, slide_cnt:int, slide_image_cnt:int) -> int:
    """ recursive function to process shapes and shapes within groups """
    # Check if the shape has a picture
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for embedded_shape in shape.shapes:
            image_cnt, slide_image_cnt = process_shapes_from_file(embedded_shape, csv_rows, image_cnt, slide_cnt, slide_image_cnt)

    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        decorative_pptx:bool = is_decorative(shape)

        # get decorative
        decorative:bool = bool_value(csv_rows[image_cnt][5])

        # change decorative status
        if decorative_pptx != decorative:
            # set decorative status of image
            print(f"Side: {slide_cnt}, {shape.name}, can't set the docorative status to: {bool_to_string(decorative)}")

        alt_text: str = ""
        if not decorative:
            # get alt text from text file
            alt_text = csv_rows[image_cnt][4]

        # set alt text
        shape_set_alt_text(shape, alt_text)
        
        slide_image_cnt += 1
        image_cnt += 1

    return image_cnt, slide_image_cnt

def main(argv: List[str]) -> int:
    """ main """
    err: bool = False

    parser = argparse.ArgumentParser(description='Add alt-text automatically to images in Powerpoint')
    parser.add_argument("file", type=str, help="Powerpoint file")
    parser.add_argument("--generate", action='store_true', default=False, help="flag to generate alt-text to images")
    parser.add_argument("--model", type=str, default="", help="Model type: kosmos-2, openclip, llava")
    # LLaVA
    parser.add_argument("--server", type=str, default="http://localhost", help="LLaVA server URL, default=http://localhost")
    parser.add_argument("--port", type=str, default="8007", help="LLaVA server port, default=8007")
    # OpenCLIP
    parser.add_argument("--openclip", type=str, default="coca_ViT-L-14", help="OpenCLIP model name")
    parser.add_argument("--pretrained", type=str, default="mscoco_finetuned_laion2B-s13B-b90k", help="OpenCLIP pretrained model")
    #
    parser.add_argument("--prompt", type=str, default="", help="Custom prompt for Kosmos-2 or LLaVA")
    parser.add_argument("--save", action='store_true', default=False, help="flag to save powerpoint file with updated alt texts")
    parser.add_argument("--add_from_file", type=str, default="", help="Add alt texts from specified file to powerpoint file")
    #
    parser.add_argument("--debug", action='store_true', default=False, help="flag for debugging")

    args = parser.parse_args()

    prompt:str = args.prompt
    model_str:str = args.model.lower()
    if model_str == "llava":
        if args.prompt == "":
            prompt = "You are an expert caption writer. Write a concise and accurate image caption. The accuracy is critically important to me."
            prompt = "Describe the image."
    elif model_str == "kosmos-2":
        if args.prompt == "":
            #prompt = "<grounding>An image of"
            prompt = "<grounding>Describe this image in detail:"

    # Read PowerPoint file and list images
    powerpoint_file_name = args.file
    if not os.path.isfile(powerpoint_file_name):
        print(f"Error: File {powerpoint_file_name} not found.")
        err = True
    else:

        settings = {
            "model": model_str,
            "kosmos2_model": None,
            "kosmos2_pretrained": None,
            "openclip_model_name": args.openclip,
            "openclip_pretrained": args.pretrained,
            "openclip-model": None,
            "openclip-transform": None,
            "llava_url": f"{args.server}:{args.port}",
            "prompt": prompt
        }
        if args.add_from_file != "":
            err = add_alt_text_from_file(powerpoint_file_name, args.add_from_file)
        else:
            err = process_images_from_pptx(powerpoint_file_name, args.generate, settings, args.save, args.debug)

    return int(err)

if __name__ == "__main__":
    EXIT_CODE = main(sys.argv[1:])
    sys.exit(EXIT_CODE)
