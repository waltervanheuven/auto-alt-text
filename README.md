# Auto-Alt-Text

Automatically create alt-text for images in Powerpoint files using [LLaVA](https://llava-vl.github.io), [OpenCLIP](https://github.com/mlfoundations/open_clip), or [Kosmos-2](https://github.com/microsoft/unilm/tree/master/kosmos-2).

## Setup

```sh
python3 -m venv venv
source venv/bin/activate

pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
```

Show current alt text of images in a Powerpoint file. Python script also creates a `txt` file with the alt text of each image in the Powerpoint file.

```sh
python source/auto-alt-text-pptx.py tmp/test.pptx

# output is also written to `tmp/test.txt`
```

## Kosmos-2

Example command for using [Kosmos-2](https://github.com/microsoft/unilm/tree/master/kosmos-2) to generate descriptions of images in Powerpoint files.

```sh
# generate alt text for images in Powerpoint file based on text generated by Kosmos-2
# note that all images in the powerpoint files are saved separately 
python source/auto-alt-text-pptx.py tmp/test.pptx --type kosmos-2 --generate

# to save a copy of the Powerpoint file as '<name>_alt_text.pptx' add --save
python source/auto-alt-text-pptx.py tmp/test.pptx --type kosmos-2 --generate --save
```

## OpenCLIP

The python script can also use [OpenCLIP](https://github.com/mlfoundations/open_clip) to generate descriptions of images in Powerpoint files.

```sh
# only show alt text already available
python source/auto-alt-text-pptx.py tmp/test.pptx

# generate alt text for images in Powerpoint file based on text generated by OpenCLIP
# note that all images in the powerpoint files are saved separately 
python source/auto-alt-text-pptx.py tmp/test.pptx --type openclip --generate

# to save a copy of the Powerpoint file as '<name>_alt_text.pptx' add --save
python source/auto-alt-text-pptx.py tmp/test.pptx --type openclip --generate --save

# specify specific OpenCLIP model and pretained model
python source/auto-alt-text-pptx.py tmp/test.pptx --type openclip --add --model coca_ViT-L-14 --pretrained mscoco_finetuned_laion2B-s13B-b90k
```

## LLaVA

If you want to use LLaVA to generate image descriptions, you need to set up a LLaVA server. An implementation of LLaVA is available in [llama.cpp](https://github.com/ggerganov/llama.cpp).

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
# note that all images in the powerpoint files are saved separately 
python source/auto-alt-text-pptx.py tmp/test.pptx --type llava --generate 

# to save a copy of the Powerpoint file as '<name>_alt_text.pptx' add --save
python source/auto-alt-text-pptx.py tmp/test.pptx --type llava --generate --save

# specify a different prompt
python source/auto-alt-text-pptx.py tmp/test.pptx --type llava --generate --prompt "Describe in simple words using maximal 125 characters"
```
