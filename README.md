# Auto-Alt-Text

Automatically create `Alt Text` for images in Powerpoint presentations using Multimodal Large Language Models (MLLM), such as [Kosmos-2](https://github.com/microsoft/unilm/tree/master/kosmos-2) and [LLaVA](https://llava-vl.github.io), or Visual-Language (VL) pre-trained models, such as [OpenCLIP](https://github.com/mlfoundations/open_clip). The Python script will create a text file with the generated `Alt Text` as well as apply these to the images in the PowerPoint file and save the updated Powerpoint to a new file.

Please note that the Python script has been tested so far only with Python 3.11 on macOS Sonoma (M1 MacBook Pro). However, it should also work on Windows and Linux. The models used run all locally.

## Setup

```sh
python3 -m venv venv
source venv/bin/activate

pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
```

Show current alt text of images in a Powerpoint file. Python script also creates a `txt` file with the alt text of each image in the Powerpoint file.

```sh
python source/auto_alt_text.py pptx/test1.pptx
# output is also written to `tmp/test.txt`
```

## Kosmos-2

Example command for using [Kosmos-2](https://github.com/microsoft/unilm/tree/master/kosmos-2) to generate descriptions of images in Powerpoint files.

```sh
# generate alt text for images in Powerpoint file based on text generated by Kosmos-2
# note that all images in the powerpoint files are saved separately in a folder
python source/auto_alt_text.py pptx/test1.pptx --model kosmos-2 --generate

# to save a copy of the Powerpoint file with the generated alt texts
# for the images add --save. Powerpoint file will be saved to '<filename>_alt_text.pptx'
python source/auto_alt_text.py pptx/test1.pptx --model kosmos-2 --generate --save

# custom prompt to get brief image descriptions
python source/auto_alt_text.py pptx/test1.pptx --model kosmos-2 --generate --save --prompt "<grounding>An image of"
```

## OpenCLIP

The python script can also use [OpenCLIP](https://github.com/mlfoundations/open_clip) to generate descriptions of images in Powerpoint files.

```sh
# only show alt text already available
python source/auto_alt_text.py pptx/test1.pptx

# generate alt text for images in Powerpoint file based on text generated by OpenCLIP
# note that all images in the powerpoint files are saved separately in a folder
python source/auto_alt_text.py pptx/test1.pptx --model openclip --generate

# to save a copy of the Powerpoint file with the generated alt texts
# for the images add --save. Powerpoint file will be saved to '<filename>_alt_text.pptx'
python source/auto_alt_text.py pptx/test1.pptx --model openclip --generate --save

# specify specific OpenCLIP model and pretained model
python source/auto_alt_text.py pptx/test1.pptx --model openclip --add --openclip coca_ViT-L-14 --pretrained mscoco_finetuned_laion2B-s13B-b90k
```

## LLaVA

If you want to use LLaVA to generate image descriptions, you need to set up a LLaVA server. A fast implementation of LLaVA is available through [llama.cpp](https://github.com/ggerganov/llama.cpp).

Steps to set up a local LLaVA server:

```sh
# clone llama.cpp repository in a tmp folder
mkdir tmp
cd tmp
git clone https://github.com/ggerganov/llama.cpp.git

# build
cd llama.cpp
make

# create folder `llava` in main folder `auto-alt-text`
cd ..
mkdir llava
mkdir llava/models
```

Copy `server` from the `tmp/llama.cpp/` folder to the `llava` folder. For Metal on macOS also copy `ggml-metal.metal` that can be found in `tmp/llama.cpp/` to the folder `llava`.

Models for [LLaVA](https://llava-vl.github.io) can be found on huggingface: [https://huggingface.co/mys/ggml_llava-v1.5-7b/tree/main](https://huggingface.co/mys/ggml_llava-v1.5-7b/tree/main)

Download `ggml-model-q5_k.gguf` and `mmproj-model-f16.gguf` and move the files to the folder `models` in the folder `llava`.

### Start LLaVA server

```sh
./llava/server -t 4 -c 4096 -ngl 50 -m llava/models/ggml-model-q5_k.gguf --host 0.0.0.0 --port 8007 --mmproj llava/models/mmproj-model-f16.gguf
```

### Example of using LLaVA

```sh
# add alt text based on text generated by LLaVA
# note that all images in the powerpoint files are saved separately in a folder
python source/auto_alt_text.py pptx/test1.pptx --model llava --generate 

# to save a copy of the Powerpoint file with the generated alt texts
# for the images add --save. Powerpoint file will be saved to '<filename>_alt_text.pptx'
python source/auto_alt_text.py pptx/test1.pptx --model llava --generate --save

# specify a different prompt
python source/auto_alt_text.py pptx/test1.pptx --model llava --generate --prompt "Describe in simple words using maximal 125 characters"
```

## Edit generated alt text and apply to Powerpoint file

The generated alt text is saved to a text file so that it can be edited. You can apply the edited alt text in the file to the powerpoint file using the command below. The Powerpoint file is saved as `<filename>_alt_text.pptx`.

```sh
python source/auto_alt_text.py pptx/test1.pptx --add_from_file pptx/test1_kosmos-2_edited.txt
```

## Help

Add `--help` to show command line options.

```sh
python source/auto_alt_text.py --help
```

## Limitations

Script will save each image in a shape group. Ideally, it should combine them to create a combined image so that it can set the alt text of the shape group.
