""" utils.py """

from typing import Tuple
import os
import sys
import io
import platform
import subprocess
import shutil
import base64
import requests
from PIL import Image

def check_server_is_running(url: str) -> bool:
    """ URL accessible? """    
    status:bool = False
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            status = True
    except requests.exceptions.Timeout:
        print("Timeout exception", file=sys.stderr)
    except requests.exceptions.RequestException as e:
        print(f"Exception: {str(e)}", file=sys.stderr)

    return status

def num2str(the_max: int, n: int) -> str:
    """ convert number to string with trailing zeros """
    s:str = f"{str(n)}"
    if the_max > 99:
        if n < 100:
            if n < 10:
                s = f"00{str(n)}"
            else:
                s = f"0{str(n)}"
    elif n < 10:
        s = f"0{str(n)}"

    return s

def str2bool(s: str) -> bool:
    """ convert str True or False to bool """
    assert(s is not None and len(s) > 0)
    return s.lower() == "true"

def bool2str(b: bool) -> str:
    """ convert bool to str """
    return "True" if b else "False"

def check_readonly_formats(image_file_path: str) -> Tuple[str, str, bool]:
    """
    Check if image format is WMF, WME, or PSD which can not be converted using the pillow library.

    Function converts WMF (vector format) to JPEG using LibreOffice.
    
    Other read only formats not yet tested. Conversion only tested on macOS.
    """
    readonly: bool = False
    msg: str = ""
    new_image_file_path = image_file_path
    err: bool = False

    with Image.open(image_file_path) as img:

        if img.format in ['WMF', 'WME']:
            msg = "A windows media format file."
            readonly = True
        elif img.format in ['PSD']:
            msg = "An Adobe Photoshop file."
            readonly = True

        if readonly and img.format in ['WMF']:
            # convert images to PNG
            dirname:str = os.path.dirname(image_file_path)
            basename:str = os.path.basename(image_file_path).split(".")[0]
            new_image_file_path = os.path.join(os.path.dirname(image_file_path), f"{basename}.png")

            print(f"Converting {img.format} to PNG...")
            try:
                # convert WMF to PNG using headless libreoffice
                if platform.system() != "Windows":
                    # convert using LibreOffice (headless)
                    cmd:list[str] = ["soffice", "--headless", "--convert-to", "png", image_file_path, "--outdir", dirname]
                    path_to_cmd = shutil.which(cmd[0])
                    if path_to_cmd is not None:
                        r = subprocess.run(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, shell=False, check=True)
                    else:
                        print("Warning, LibreOffice not installed.")
                elif platform.system() == "Windows":
                    # convert using magick
                    cmd:list[str] = ["magick", "convert", image_file_path, new_image_file_path]
                    path_to_cmd = shutil.which(cmd[0])
                    if path_to_cmd is not None:
                        r = subprocess.run(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, shell=False, check=True)
                    else:
                        print("Warning, ImageMagick not installed.")
            except subprocess.CalledProcessError as e:
                msg = f"soffice CalledProcessError: {str(e)}"
                err = True
            except subprocess.TimeoutExpired as e:
                msg = f"soffice TimeoutExpired: {str(e)}"
                err = True
            except OSError as e:
                msg = f"soffice OSError, file does not exist?: {str(e)}"
                err = True
            #except Exception as e:
            #    msg = f"soffice exception: {str(e)}"
            #    err = True
            else:
                readonly = False

            if err:
                readonly = True
                new_image_file_path = image_file_path
                print(r.stderr, file=sys.stderr)

    if readonly:
        print(f"Warning, unable to open '{img.format}' file. Replace image in powerpoint with PNG, TIFF, or JPEG version.")

    return new_image_file_path, readonly, msg

def convert_img_to_jpg(
        image_file_path: str,
        verbose: bool = False,
    ) -> str:
    """ convert image file to jpg """
    
    with Image.open(image_file_path) as img:
        if img.format != 'JPEG':
            # convert if not already in jpg
            basename: str = os.path.basename(image_file_path).split(".")[0]
            jpeg_image_file_path = os.path.join(os.path.dirname(image_file_path), f"{basename}.jpg")

            img = img.convert('RGB')
            img.save(jpeg_image_file_path, 'JPEG')
            image_file_path = jpeg_image_file_path

    if verbose:
        print(f"Image file size: {os.path.getsize(image_file_path):,} bytes")

    return image_file_path

def img_file_to_base64(image_file_path: str , settings: dict) -> str:
    """ load image, resize, and convert to base64_str """

    with Image.open(image_file_path) as original_img:
        im = original_img.convert("RGB")

        # resize image
        im = resize(im, settings)

        # check
        buffer = io.BytesIO()
        im.save(buffer, format=original_img.format.upper())
        buffer.seek(0)

    # Encode the image bytes to Base64
    base64_bytes = base64.b64encode(buffer.getvalue())

    # str
    base64_str = base64_bytes.decode('utf-8')

    return base64_str

def resize(
        image: Image.Image,
        settings: dict,
        verbose: bool = False
    ) -> Image.Image:
    """ resize image """
    px: int = settings["img_size"]
    if px != 0:
        # only resize if img_size != 0
        if image.width > px or image.height > px:
            new_size = (min(px, image.width), min(px, image.height))
            if verbose:
                print(f"Resize image from ({image.width}, {image.height}) to {new_size}")
            image = image.resize(new_size)

    return image
