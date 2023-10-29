# Auto-Alt-Text

Automatically create alt-text for images in Powerpoint files using [LLaVA](https://llava-vl.github.io).

## Setup

```sh
python3 -m venv venv
source venv/bin/activate

pip install --upgrade pip setuptools wheel
pip install python-pptx Pillow requests
```

Show images and current alt text in Powerpoint file

```sh
python source/auto-alt-text.py tmp/test.pptx
```

## Install llama.cp (LLaVA)

```sh
# clone llama.cpp repository
git clone https://github.com/ggerganov/llama.cpp.git

# build
make

# create folder `llama.cpp` in main folder `auto-alt-text`
mkdir llama.cpp
mkdir llama.cpp/models
```

Copy `server` from `llama.cpp` directory to `llama.cpp` with `auto-alt-text`. For Metal on macOS also copy `ggml-metal.metal` to the folder `llama.cpp`.

Required models for [LLaVA](https://llava-vl.github.io) can be found on huggingface: [https://huggingface.co/mys/ggml_llava-v1.5-7b/tree/main](https://huggingface.co/mys/ggml_llava-v1.5-7b/tree/main)

Download `ggml-model-q5_k.gguf` and `mmproj-model-f16.gguf` and move to folder `ggml_llava-v1.5-7b` within the folder `llama.cpp/models`.

Start server:

```sh
./llama.cpp/server -t 4 -c 4096 -ngl 50 -m llama.cpp/models/ggml-model-q5_k.gguf --host 0.0.0.0 --port 8007 --mmproj llama.cpp/models/mmproj-model-f16.gguf
```

Test:

```sh
python source/main.py tmp/test.pptx --add
```
