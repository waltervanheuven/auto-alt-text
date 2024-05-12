# Auto-Alt-Text

Automatically create `Alt Text` for images and other objects in Powerpoint presentations using Multimodal Large Language Models (MLLM) or Visual-Language (VL) pre-trained models. The Python script will create a text file with the generated `Alt Text` as well as apply these to the images and objects in the PowerPoint file and save the updated Powerpoint to a new file.

The script currently supports the following models:

- [Qwen-VL](https://github.com/QwenLM/Qwen-VL)
- [CogVLM](https://github.com/THUDM/CogVLM)
- [Kosmos-2](https://github.com/microsoft/unilm/tree/master/kosmos-2)
- [OpenCLIP](https://github.com/mlfoundations/open_clip)
- [GPT-4V](https://openai.com/research/gpt-4v-system-card)
- Multimodal models, such as [LLaVA](https://llava-vl.github.io) are supported through [Ollama](https://ollama.com)

All models, except GPT-4V, run locally. GPT-4V requires API access. By default, images are resized so that width and height are maximum 500 pixels before inference. The [Qwen-VL](https://github.com/QwenLM/Qwen-VL) model requires an NVIDIA RTX A4000 (or better), or an M1-Max or better. For inference hardware requirements of Cog-VLM, check the [github page](https://github.com/THUDM/CogVLM).

## Setup

### macOS/Linux

Install latest Python 3.11 on macOS using [brew](https://brew.sh).

```sh
git clone https://github.com/waltervanheuven/auto-alt-text.git
cd auto-alt-text

python3 -m venv venv
source venv/bin/activate

pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
```

To generate `Alt Text` for Windows Metafile (WMF) images in Powerpoint on macOS and Linux, the script needs [LibreOffice](https://www.libreoffice.org) to convert WMF to a bitmap format. On macOS use [brew](https://brew.sh) to install LibreOffice. Furthermore, for additional functionality install [qpdf](https://github.com/qpdf/qpdf) and [ImageMagick](https://imagemagick.org).

```sh
brew install libreoffice
brew install qpdf
brew install imagemagick
```

For cuda support on Linux, follow instructions on the [PyTorch website](https://pytorch.org/get-started/locally/) to install torch with cuda support.

### Windows

Install latest Python 3.11 on Windows using, for example, [scoop](https://scoop.sh).

```sh
git clone https://github.com/waltervanheuven/auto-alt-text.git
cd auto-alt-text

python311 -m venv venv
.\venv\Scripts\activate

python -m pip install --upgrade pip setuptools wheel
python -m pip install -r .\requirements.txt

# for cuda support install torch (cuda 12.1)
python -m pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu121
```

To generate `Alt Text` for Windows Metafile (WMF) images in Powerpoint on Windows, the script needs [ImageMagick](https://imagemagick.org) to convert WMF to a bitmap format. Use [scoop](https://scoop.sh) to install imagemagick.

```sh
scoop install main/imagemagick
scoop install main/qpdf
```

### Additional libs for Qwen-VL and Cog-VL models

```sh
# Qwen-VL
pip install tiktoken transformers_stream_generator einops optimum

# for auto_gptq on MacOS, run:
BUILD_CUDA_EXT=0 pip install auto_gptq
# else
pip install auto_gptq

# Cog-VL
pip install matplotlib optimum scipy
pip install xformers accelerate bitsandbytes
```

## Generate accessibility report

Show current alt text of objects (e.g. images, shapes, group shapes) in a Powerpoint file and generate an alt text accessibility report. A tab-delimited text file is created with the alt text of each object in the Powerpoint file.

```sh
python source/auto_alt_text.py pptx/test1.pptx --report
# output is written to `pptx/test1.txt`
```

## Kosmos-2

Example command for using [Kosmos-2](https://github.com/microsoft/unilm/tree/master/kosmos-2). Script will download the Komos-2 model (~6.66GB).

```sh
# Generate alt text for images in the Powerpoint file based using the specified model (e.g. kosmos-2)
#
# Note that all images in the powerpoint files are saved separately in a folder
# Powerpoint file with the alt texts will be saved to '<filename>_<model_name>.pptx'
python source/auto_alt_text.py pptx/test1.pptx --model kosmos-2

# custom prompt to get brief image descriptions
# for Kosmos-2 start prompt with <grounding>
python source/auto_alt_text.py pptx/test1.pptx --model kosmos-2 --prompt "<grounding>An image of"
```

## Qwen-VL

Example command for using [Qwen-VL](https://github.com/QwenLM/Qwen-VL). Script will download the Qwen-VL-Chat model (~9.75GB) if CUDA support is available. On Apple Silicon Macs the Qwen-VL-Chat-Int4 is used.

Qwen-VL only tested with an RTX A4000 GPU on Windows and with an M1-Max on macOS (32GB RAM).

```sh
python source/auto_alt_text.py pptx/test1.pptx --model qwen-vl

# custom prompt to get brief image descriptions
python source/auto_alt_text.py pptx/test1.pptx --model qwen-vl --prompt "What is the key information illustrated in this image"
```

## OpenCLIP

The Python script can also use [OpenCLIP](https://github.com/mlfoundations/open_clip) to generate descriptions of images in Powerpoint files. There are many OpenCLIP models and pretrained models that you can use. To find out the available models, use `--show_openclip_models`. The default model is `coca_ViT-L-14` and the pretrained model is `mscoco_finetuned_laion2B-s13B-b90k` (~2.55Gb model file will be downloaded).

```sh
python source/auto_alt_text.py pptx/test1.pptx --model openclip

# list available OpenCLIP models
python source/auto_alt_text.py pptx/test1.pptx --show_openclip_models

# specify specific OpenCLIP model and pretained model
python source/auto_alt_text.py pptx/test1.pptx --model openclip --openclip_model coca_ViT-L-14 --openclip_pretrained mscoco_finetuned_laion2B-s13B-b90k
```

## GPT-4V

To use [GPT-4V](https://openai.com/research/gpt-4v-system-card) you need to have [API access](https://help.openai.com/en/articles/7102672-how-can-i-access-gpt-4). Images will be send to OpenAI servers for inference. Costs for using the API depend on the size and number the images. API access [pricing information](https://openai.com/pricing#language-models). The script uses the OPENAI_API_KEY environment variable. Information how to set/add this variable can be found in the [OpenAI quickstart docs](https://platform.openai.com/docs/quickstart?context=python).

```sh
python source/auto_alt_text.py pptx/test1.pptx --model gpt-4v

# custom prompt
python source/auto_alt_text.py pptx/test1.pptx --model gpt-4v --prompt "Describe clearly in two sentences"
```

## LLaVA and other multimodal LLMs

LLaVA or other multimodal large language models can be used through [Ollama](https://ollama.com/). These models will run locally. Which model you can use depends on the capabilities of your computer (e.g. memory, GPU).

To install Ollama [download the Ollama app](https://ollama.com/download).

Next, download LLaVA model.

```sh
# download latest LLaVA model (v1.6)
ollama pull llava

# Check which models are available on your computer
ollama list
```

### Example of using LLaVA through Ollama

```sh
python source/auto_alt_text.py pptx/test1.pptx --model  llava --use_ollama

# to disable default image resizing to 500px x 500px, set resize size to 0
python source/auto_alt_text.py pptx/test1.pptx --model  llava --use_ollama --resize 0

# specify a different prompt
python source/auto_alt_text.py pptx/test1.pptx --model llava --use_ollama --prompt "Describe in simple words using one sentence."

# specify differ server or port of the ollama server, default server is localhost, and port is 11434
python source/auto_alt_text.py pptx/test1.pptx --model  llava --use_ollama --server my_server.com --port 3456
```

## Edit generated alt texts and apply to Powerpoint file

The generated alt texts are saved to a text file so that these it can be edited. You can apply the edited alt texts in the file to the powerpoint file using the option `--replace`. The Powerpoint file is saved as `<filename>_alt_text.pptx`.

```sh
python source/auto_alt_text.py pptx/test1.pptx --replace pptx/test1_kosmos-2_edited.txt
```

## Presenter notes

The models are prompted to generate alt texts using one or two senteneces for each image. For complex images and figures this description might not be sufficient, therefore a longer desciption of the slide as a whole can be generated to improve accessibility. This slide description will be placed in the slide presenter notes. The most accurate slide descriptions will be generated by multimodal LLMs (e.g. LLaVA, GPT-4V). To create slide descriptions when the slide has at least one image or non-text object, add `--add_to_notes`.

```sh
python source/auto_alt_text.py pptx/test1.pptx --model llava:latest --use_ollama --add_to_notes
```

## Help

Add `--help` to show all command line options.

```sh
python source/auto_alt_text.py --help
```

## Known issues

- If the script reports `Unable to access image file:`, delete the generated folder for the pptx file.
- OpenCLIP at the moment only works with cuda devices.
- Qwen-VL does work with 'mps' but slow because fail back to CPU.
