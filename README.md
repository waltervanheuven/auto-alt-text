# Auto-Alt-Text

Automatically create alt-text for images in Powerpoint files using [LLaVA](https://llava-vl.github.io) or [OpenClip](https://github.com/mlfoundations/open_clip).

## Setup

```sh
python3 -m venv venv
source venv/bin/activate

pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
```

Show images and current alt text in Powerpoint file. Python script also creates a `txt` file with the alt text of each image in the powerpoint file.

```sh
python source/auto-alt-text-pptx.py tmp/test.pptx

# output is also written to `tmp/test.txt`
```

## Set up LLaVA server (only required for LLaVA version)

Set up LLaVA server so that the Python script can use this server to obtain descriptions of each image. Script will set the alt text based on the descriptions provided by LLaVA.

The implementation of LLaVA is available in [llama.cpp](https://github.com/ggerganov/llama.cpp).

```sh
# clone llama.cpp repository
git clone https://github.com/ggerganov/llama.cpp.git

# build
make

# create folder `llama.cpp` in main folder `auto-alt-text`
mkdir llama.cpp
mkdir llama.cpp/models
```

Copy `server` from `llama.cpp` directory to `llama.cpp` within the folder `auto-alt-text`. For Metal on macOS also copy `ggml-metal.metal` to the folder `llama.cpp`.

Required models for [LLaVA](https://llava-vl.github.io) can be found on huggingface: [https://huggingface.co/mys/ggml_llava-v1.5-7b/tree/main](https://huggingface.co/mys/ggml_llava-v1.5-7b/tree/main)

Download `ggml-model-q5_k.gguf` and `mmproj-model-f16.gguf` and move the files to the folder `models` in the folder `llama.cpp`.

Start server:

```sh
./llama.cpp/server -t 4 -c 4096 -ngl 50 -m llama.cpp/models/ggml-model-q5_k.gguf --host 0.0.0.0 --port 8007 --mmproj llama.cpp/models/mmproj-model-f16.gguf
```

Add alt text using LLaVA. Note that the script also saves each image in the Powerpoint so that it is possible to use other LLMs for the description of the images.

```sh
python source/auto-alt-text-pptx-llava.py tmp/test.pptx --add

python source/auto-alt-text-pptx-llava.py tmp/test.pptx --add --prompt "Describe in simple words"
```

For OpenClip no server is needed. Model will be downloaded the first time it is used.

```sh
python source/auto-alt-text-pptx-openclip.py tmp/test.pptx
```
