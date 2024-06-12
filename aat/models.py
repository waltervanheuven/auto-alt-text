""" models.py """

from typing import Tuple, List
import os
import sys
import platform
import json
import re
import open_clip
import torch
import requests
import psutil
from PIL import Image
from transformers import AutoProcessor, AutoModelForVision2Seq
from transformers import AutoModelForCausalLM, AutoTokenizer, LlamaTokenizer
#from transformers import LlavaNextProcessor, LlavaNextForConditionalGeneration
from transformers.generation import GenerationConfig
from .utils import resize, check_readonly_formats, convert_img_to_jpg, img_file_to_base64
if platform.system() == "Darwin":
    from mlx_vlm import load, generate

def init_model(
        settings: dict,
    ) -> bool:
    """ download and init model for inference """
    err: bool = False
    model_str: str = settings["model"]
    prompt: str = settings["prompt"]

    if settings["use_ollama"]:
        err = setup_ollama(settings, prompt)
    elif settings["use_mlx_vlm"]:
        err = setup_mlx_vlm(settings)
    elif model_str == "kosmos-2":
        err = setup_kosmos2(settings, prompt)
    elif model_str == "openclip":
        err = setup_openclip(settings)
    elif model_str == "qwen-vl":
        err = setup_qwen_vl(settings, prompt)
    elif model_str == "cogvlm" or model_str == "cogvlm2":
        err = setup_cog_vlm(settings, prompt)
    elif model_str == "phi3-vision":
        err = setup_phi3_vision(settings)
    elif model_str == "gpt-4o" or model_str == "gpt-4-turbo":
        print("OpenAI")
        print(f"model: {model_str}")
        print(f"prompt: '{prompt}'")
    else:
        print(f"Unknown model: '{model_str}'")
        err = True

    return err

def setup_kosmos2(settings: dict, prompt: str) -> bool:
    """ setup Kosmos-2 """
    model_name = "microsoft/kosmos-2-patch14-224"
    print(f"Kosmos-2 model: '{model_name}'")
    print(f"prompt: '{prompt}'")
    m = AutoModelForVision2Seq.from_pretrained(model_name)
    p = AutoProcessor.from_pretrained(model_name)
    if settings["cuda_available"]:
        print("Using CUDA.")
        m.cuda()
    settings["kosmos2-model"] = m
    settings["kosmos2-processor"] = p
    return False

def kosmos2(
        image_file_path: str,
        extension: str,
        settings: dict,
        for_notes: bool = False,
        verbose: bool = False
    ) -> Tuple[str, bool]:
    """ get image description from Kosmos-2 """
    err:bool = False

    # check if readonly
    image_file_path, readonly, msg = check_readonly_formats(image_file_path, extension, verbose)
    if readonly:
        return msg, False

    with Image.open(image_file_path) as img:

        # resize image
        img = resize(img, settings, verbose)

        # prompt
        if for_notes:
            prompt = settings["prompt_notes"]
        else:
            prompt = settings["prompt"]

        processor:str = settings["kosmos2-processor"]
        model:str = settings["kosmos2-model"]

        inputs = processor(text=prompt, images=img, return_tensors="pt")
        if settings["cuda_available"]:
            generated_ids = model.generate(
                pixel_values=inputs["pixel_values"].cuda(),
                input_ids=inputs["input_ids"].cuda(),
                attention_mask=inputs["attention_mask"].cuda(),
                image_embeds=None,
                image_embeds_position_mask=inputs["image_embeds_position_mask"].cuda(),
                use_cache=True,
                max_new_tokens=256,
            )
        else:
            generated_ids = model.generate(
                pixel_values=inputs["pixel_values"],
                input_ids=inputs["input_ids"],
                attention_mask=inputs["attention_mask"],
                image_embeds=None,
                image_embeds_position_mask=inputs["image_embeds_position_mask"],
                use_cache=True,
                max_new_tokens=256,
            )

    generated_text = processor.batch_decode(generated_ids, skip_special_tokens=True)[0]

    # Specify `cleanup_and_extract=False` in order to see the raw model generation.
    #processed_text = processor.post_process_generation(generated_text, cleanup_and_extract=True)

    # processed_text, entities = processor.post_process_generation(generated_text)
    processed_text, _ = processor.post_process_generation(generated_text)

    # remove prompt
    p:str = re.sub('<[^<]+?>', '', prompt)
    processed_text = processed_text.replace(p.strip(), '')

    # capitalize
    alt_text:str = processed_text.strip().capitalize()

    return alt_text, err

def show_openclip_models() -> None:
    """ show available openclip model """
    openclip_models = open_clip.list_pretrained()
    print("OpenCLIP models:")
    for m, p in openclip_models:
        print(f"Model: {m}, pretrained model: {p}")

def setup_openclip(settings: dict) -> bool:
    """ setup OpenCLIP model """
    err: bool = False

    print(f"OpenCLIP model: '{settings['openclip_model_name']}'\npretrained model: '{settings['openclip_pretrained']}'")

    #model, preprocess = open_clip.create_model_from_pretrained(settings["openclip_model_name"])
    #tokenizer = open_clip.get_tokenizer(settings["openclip_pretrained"])

    if settings["cuda_available"]:
        my_device = "cuda"
    else:
        # CPU is slow, mps not supported yet
        my_device = "cpu"
        #if platform.system() == "Darwin":
        #    my_device = "mps"

    if not err:
        model, _, preprocess = open_clip.create_model_and_transforms(
            model_name=settings["openclip_model_name"],
            pretrained=settings["openclip_pretrained"],
            device=my_device,
            #precision="fp16"
        )
        settings["openclip-model"] = model
        settings["openclip-preprocess"] = preprocess

    return err

def openclip(
        image_file_path: str,
        extension: str,
        settings: dict,
        #for_notes: bool = False,
        verbose: bool = False
    ) -> Tuple[str, bool]:
    """ get image description from OpenCLIP """
    err: bool = False

    # check if readonly
    image_file_path, readonly, msg = check_readonly_formats(image_file_path, extension, verbose)
    if readonly:
        return msg, False

    with Image.open(image_file_path).convert('RGB') as img:
        # resize image
        img = resize(img, settings, verbose)

        preprocess = settings["openclip-preprocess"]
        img = preprocess(img).unsqueeze(0)

    # use OpenCLIP model to create label
    model = settings["openclip-model"]

    if settings["cuda_available"]:
        with torch.no_grad(), torch.cuda.amp.autocast():
            generated = model.generate(img)
    else:
        #if platform.system() == "Darwin":
        #    with torch.no_grad(), torch.autocast('mps'):
        #        generated = model.generate(img)
        #else:
        with torch.no_grad(), torch.autocast('cpu'):
            generated = model.generate(img)

    # get picture description and remove trailing spaces
    alt_text = open_clip.decode(generated[0]).split("<end_of_text>")[0].replace("<start_of_text>", "").strip()

    # remove space before '.' and capitalize
    alt_text = alt_text.replace(' .', '.').capitalize()

    return alt_text, err

def setup_qwen_vl(settings: dict, prompt: str) -> bool:
    """ setip Qwen VL model """
    err: bool = False
    total_memory_bytes = psutil.virtual_memory().total
    total_memory_gb = total_memory_bytes / (1024**3)

    if settings["cuda_available"]:
        model_name = "Qwen/Qwen-VL-Chat"
        #model_name = "Qwen/Qwen-VL-Max"
        print(f"Qwen-VL model: '{model_name}'")
        print(f"prompt: '{prompt}'")
        print("Using CUDA.")

        tokenizer = AutoTokenizer.from_pretrained(model_name, trust_remote_code=True)
        model = AutoModelForCausalLM.from_pretrained(model_name, device_map="cuda", trust_remote_code=True).eval()
        model.generation_config = GenerationConfig.from_pretrained(model_name, trust_remote_code=True)
    elif settings['mps_available']:
        os.environ["ACCELERATE_USE_MPS_DEVICE"] = "True"
        model_name = "Qwen/Qwen-VL-Chat-Int4"
        #model_name = "Qwen/Qwen-VL"

        if total_memory_gb >= 32:
            print(f"Qwen-VL model: '{model_name}'")
            print(f"prompt: '{prompt}'")

            tokenizer = AutoTokenizer.from_pretrained(model_name, trust_remote_code=True)
            #model = AutoModelForCausalLM.from_pretrained(model_name, load_in_4bit=True, device_map="mps", trust_remote_code=True).eval()
            model = AutoModelForCausalLM.from_pretrained(model_name, device_map="mps", trust_remote_code=True).eval()
            model.generation_config = GenerationConfig.from_pretrained(model_name, trust_remote_code=True)
        else:
            print(f"Model '{model_name}' requires >= 32GB RAM.", file=sys.stderr)
            err = True
    else:
        print(f"Model '{model_name}' requires a GPU with CUDA support", file=sys.stderr)
        err = True

    settings["qwen-vl-model"] = model
    settings["qwen-vl-tokenizer"] = tokenizer

    return err

def qwen_vl(
        image_file_path: str,
        extension: str,
        settings: dict,
        for_notes: bool = False,
        verbose: bool = False
    ) -> Tuple[str, bool]:
    """ get image description from Qwen-VL """
    err: bool = False

    # check if readonly
    image_file_path, readonly, msg = check_readonly_formats(image_file_path, extension, verbose)
    if readonly:
        return msg, False

    # prompt
    prompt: str
    if for_notes:
        prompt = settings["prompt_notes"]
    else:
        prompt = settings["prompt"]

    model: str = settings["qwen-vl-model"]
    tokenizer: str = settings["qwen-vl-tokenizer"]

    # with Image.open(image_file_path).convert('RGB') as img:
    #     # resize image
    #     img = resize(img, settings)
    #     # prompt
    #     prompt:str = settings["prompt"]
    #     model:str = settings["cogvlm-model"]
    #     tokenizer:str = settings["cogvlm-tokenizer"]
    #     query = tokenizer.from_list_format([
    #         {'image': img},
    #         {'text': prompt},
    #     ])
    #     alt_text, history = model.chat(tokenizer, query=query, history=None)

    alt_text, _ = model.chat(tokenizer, query=f'<img>{image_file_path}</img>{prompt}', history=None)

    return alt_text, err

def setup_cog_vlm(settings: dict, prompt: str) -> bool:
    """ setup cog-vlm model """
    err: bool = False
    model_str: str = settings["model"]

    if model_str == "cogvlm":
        model_name = "THUDM/cogvlm-chat-hf"
    elif model_str == "cogvlm2":
        model_name = "THUDM/cogvlm2-llama3-chat-19B"
    print(f"CogVLM model: '{model_name}'")
    print(f"prompt: '{prompt}'")

    if settings["cuda_available"]:
        print("Using CUDA.")

    if model_str == "cogvlm":
        tokenizer = LlamaTokenizer.from_pretrained('lmsys/vicuna-7b-v1.5')
    elif model_str == "cogvlm2":
        tokenizer = AutoTokenizer.from_pretrained(
            model_name,
            trust_remote_code=True
        )

    torch_type = torch.bfloat16 if torch.cuda.is_available() and torch.cuda.get_device_capability()[0] >= 8 else torch.float16
    if settings["cuda_available"]:
        if model_str == "cogvlm":
            model = AutoModelForCausalLM.from_pretrained(
                model_name,
                load_in_4bit=True,
                #torch_dtype=torch.bfloat16,
                low_cpu_mem_usage=True,
                trust_remote_code=True
            ).to('cuda').eval()
        elif model_str == "cogvlm2":
            model = AutoModelForCausalLM.from_pretrained(
                model_name,
                torch_dtype=torch_type,
                trust_remote_code=True,
            ).to('cuda').eval()
    else:
        if model_str == "cogvlm":
            print("CogVLM requires a GPU with CUDA support.", file=sys.stderr)
            err = True
        elif model_str == "cogvlm2":
            model = AutoModelForCausalLM.from_pretrained(
                model_name,
                torch_dtype=torch_type,
                trust_remote_code=True,
            ).to('cpu').eval()

        # if settings['mps_available']:
        #     os.environ["ACCELERATE_USE_MPS_DEVICE"] = "True"
        #     print("Activating mps device for accelerate.")

        # #     print("Not yet working on mps devices")
        # model = AutoModelForCausalLM.from_pretrained(
        #     model_name,
        #     load_in_4bit=True,
        #     #torch_dtype=torch.bfloat16,
        #     low_cpu_mem_usage=True,
        #     trust_remote_code=True,
        # ).to('mps').eval()
        # # else:
        #     print("Not yet working on cpu")
        #     model = AutoModelForCausalLM.from_pretrained(
        #         model_name,
        #         load_in_4bit=True,
        #         trust_remote_code=True,
        #     ).to('cpu').eval()

    settings["cogvlm-model"] = model
    settings["cogvlm-tokenizer"] = tokenizer

    return err

def cog_vlm(
        image_file_path: str,
        extension: str,
        settings: dict,
        for_notes: bool = False,
        verbose: bool = False
    ) -> Tuple[str, bool]:
    """ get image description from CogVLM1 / CogVLM2 """
    err: bool = False

    # check if readonly
    image_file_path, readonly, msg = check_readonly_formats(image_file_path, extension, verbose)
    if readonly:
        return msg, False

    with Image.open(image_file_path).convert('RGB') as img:

        # resize image
        img = resize(img, settings, verbose)

        # prompt
        if for_notes:
            prompt:str = settings["prompt_notes"]
        else:
            prompt:str = settings["prompt"]

        model:str = settings["cogvlm-model"]
        tokenizer:str = settings["cogvlm-tokenizer"]

        inputs = model.build_conversation_input_ids(tokenizer, query=prompt, history=[], images=[img])

    if settings["cuda_available"]:
        inputs = {
            'input_ids': inputs['input_ids'].unsqueeze(0).to('cuda'),
            'token_type_ids': inputs['token_type_ids'].unsqueeze(0).to('cuda'),
            'attention_mask': inputs['attention_mask'].unsqueeze(0).to('cuda'),
            'images': [[inputs['images'][0].to('cuda').to(torch.bfloat16)]],
        }
    # elif settings["mps_available"]:
    #     inputs = {
    #         'input_ids': inputs['input_ids'].unsqueeze(0).to('mps'),
    #         'token_type_ids': inputs['token_type_ids'].unsqueeze(0).to('mps'),
    #         'attention_mask': inputs['attention_mask'].unsqueeze(0).to('mps'),
    #         'images': [[inputs['images'][0].to('mps').to(torch.bfloat16)]],
    #     }
    else:
        inputs = {
            'input_ids': inputs['input_ids'].unsqueeze(0).to('cpu'),
            'token_type_ids': inputs['token_type_ids'].unsqueeze(0).to('cpu'),
            'attention_mask': inputs['attention_mask'].unsqueeze(0).to('cpu'),
            'images': [[inputs['images'][0].to('cpu').to(torch.bfloat16)]],
        }

    gen_kwargs = {"max_length": 2048, "do_sample": False}

    alt_text: str = ""
    with torch.no_grad():
        outputs = model.generate(**inputs, **gen_kwargs)
        outputs = outputs[:, inputs['input_ids'].shape[1]:]

        alt_text = tokenizer.decode(outputs[0])

    return alt_text, err

def use_openai(
        image_file_path: str,
        extension: str,
        settings: dict,
        for_notes: bool = False,
        verbose: bool = False,
        debug: bool = False
    ) -> Tuple[str, bool]:
    """ get image description from GPT-4V """
    err: bool = False
    alt_text:str = "Error"

    # check if readonly
    image_file_path, readonly, msg = check_readonly_formats(image_file_path, extension, verbose)
    if readonly:
        return msg, False

    api_key = os.environ.get("OPENAI_API_KEY")
    if api_key is None or api_key == "":
        # otherwise get api key from args
        api_key = settings['openai_api_key']
        if api_key == "":
            print("OPENAI_API_KEY not found in environment", file=sys.stderr)
            err = True
    else:
        image_file_path = convert_img_to_jpg(image_file_path, verbose)
        img_base64_str = img_file_to_base64(image_file_path, settings)

        # prompt
        prompt:str
        if for_notes:
            prompt = settings["prompt_notes"]
        else:
            prompt = settings["prompt"]

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
        payload = {
            "model": settings["model"],
            "messages": [
            {
                "role": "user",
                "content": [
                {
                    "type": "text",
                    "text": prompt
                },
                {
                    "type": "image_url",
                    "image_url": {
                    "url": f"data:image/{extension};base64,{img_base64_str}"
                    }
                }
                ]
            }
            ],
            "max_tokens": 300
        }

        openai_server = "https://api.openai.com/v1/chat/completions"
        try:
            response = requests.post(openai_server, headers=headers, json=payload, timeout=20)

            json_out = response.json()

            if debug:
                print(json.dumps(json_out, indent=4))

            if 'error' in json_out:
                print()
                print(json_out['error']['message'], file=sys.stderr)
                err = True
            else:
                alt_text = json_out["choices"][0]["message"]["content"]
        except requests.exceptions.ConnectionError:
            print(f"ConnectionError: Unable to access the server at: '{openai_server}'", file=sys.stderr)
            err = True
        except TimeoutError:
            print("TimeoutError", file=sys.stderr)
            err = True
        except Exception as e:
            print(f"Exception: '{str(e)}'", file=sys.stderr)
            err = True

    return alt_text, err

def setup_ollama(settings: dict, prompt: str) -> bool:
    """ setup Ollama """
    err: bool = False
    print(f"Ollama server: {settings['ollama_url']}")

    # check if Ollama model is available on the server
    err, full_model_name = check_ollama_model_available(settings)

    if not err:
        settings['model'] = full_model_name
        print(f"Model: '{settings['model']}'")
        print(f"Prompt: '{prompt}'")
    
    return err

def check_ollama_model_available(settings: dict) -> Tuple[bool, str]:
    """ Check if model is available on Ollama server """
    err: bool = False
    model_specified = settings["model"]
    if ":" not in model_specified:
        model_specified = f"{model_specified}:latest"

    # check if model available
    ollama_url = f"{settings['ollama_url']}/api/tags"
    err, all_models = get_ollama_models(ollama_url)

    if not err:
        err = False
        if model_specified not in all_models:
            err = True

        if err:
            print(f"Model: '{model_specified}' not available on the Ollama server", file=sys.stderr)
            print("Models available on the Ollama server:", file=sys.stderr)
            for m in all_models:
                print(f"  {m}", file=sys.stderr)
            print()
            print("Please pull the model using Ollama or use one of the other models available", file=sys.stderr)

    return err, model_specified

def get_ollama_models(ollama_url: str) -> Tuple[bool, List[str]]:
    """ get list of models available on the server """
    all_models: List[str] = []
    err: bool = False

    try:
        response = requests.get(ollama_url, timeout=10)
        response.raise_for_status()

    except requests.exceptions.ConnectionError:
        print(f"ConnectionError: Unable to access the server at: '{ollama_url}'", file=sys.stderr)
        err = True
    except requests.exceptions.InvalidSchema:
        print(f"InvalidSchema: Unable to access the server at: '{ollama_url}'", file=sys.stderr)
        err = True
    except TimeoutError:
        print("TimeoutError", file=sys.stderr)
        err = True
    else:
        try:
            json_out = response.json()
        except requests.exceptions.JSONDecodeError:
            print("JSONDecodeError", file=sys.stderr)
            err = True
        else:
            ollama_model_response = json_out["models"]
            for m in ollama_model_response:
                all_models.append(m['name'])

    return err, all_models

def use_ollama(
        image_file_path: str,
        extension: str,
        settings: dict,
        for_notes: bool = False,
        verbose: bool = False,
        debug: bool = False
    ) -> Tuple[str, bool]:
    """ get image description from model accessed Ollama server """
    err: bool = False
    alt_text: str = "Error"

    # check if readonly
    image_file_path, readonly, msg = check_readonly_formats(image_file_path, extension, verbose)
    if readonly:
        return msg, False

    #image_file_path = convert_img_to_jpg(image_file_path, verbose)
    img_base64_str = img_file_to_base64(image_file_path, settings)

    if len(img_base64_str) > 0:
        # prompt
        prompt: str
        if for_notes:
            prompt = settings["prompt_notes"]
        else:
            prompt = settings["prompt"]

        headers = {
            "Content-Type": "application/json",
        }

        payload = {
            "model": settings["model"],
            "prompt": f"{prompt}",
            "stream": False,
            "images": [ img_base64_str ]
        }

        # http://localhost:11434/api/generate
        ollama_url = f"{settings['ollama_url']}/api/generate"
        try:
            response = requests.post(ollama_url, headers=headers, json=payload, timeout=60)
            response.raise_for_status()
        except requests.exceptions.ConnectionError:
            print(f"ConnectionError: Unable to access the server at: '{ollama_url}'", file=sys.stderr)
            err = True
        except requests.exceptions.ReadTimeout:
            print("ReadTimeout", file=sys.stderr)
            err = True
        except requests.exceptions.HTTPError:
            print("HTTPError", file=sys.stderr)
            err = True
        except TimeoutError:
            print("TimeoutError", file=sys.stderr)
            err = True
        else:
            json_out = response.json()

            if debug:
                print(json.dumps(json_out, indent=4))

            if 'error' in json_out:
                print("ERROR in ouput", file=sys.stderr)
                print(json.dumps(json_out, indent=4))
                err = True
            else:
                alt_text = json_out["response"]
                # remove newlines
                # alt_text = alt_text.replace("\n", " ")
                # remove double spaces
                alt_text = alt_text.replace("  ", " ")
    else:
        print("Error, image size is zero", file = sys.stderr)
        err = True

    return alt_text, err

def setup_phi3_vision(settings: dict) -> bool:
    """ Setup Phi3 Vision, only works with certain GPUs, see https://huggingface.co/microsoft/Phi-3-vision-128k-instruct for more info """
    model_id = "microsoft/Phi-3-vision-128k-instruct"
    
    model = AutoModelForCausalLM.from_pretrained(model_id, device_map="cuda", trust_remote_code=True, torch_dtype="auto")
    #model = AutoModelForCausalLM.from_pretrained('microsoft/Phi-3-vision-128k-instruct', device_map="cuda", trust_remote_code=True, torch_dtype="auto", _attn_implementation="eager")

    processor = AutoProcessor.from_pretrained(model_id, trust_remote_code=True)

    settings['model'] = model_id
    settings['phi3-vision-model'] = model
    settings['phi3-vision-tokenizer'] = processor
    
    return False

def phi3_vision(
        image_file_path: str,
        extension: str,
        settings: dict,
        for_notes: bool = False,
        verbose: bool = False
    ) -> Tuple[str, bool]:
    """ Phi3 Vision only works with certain GPUs, see https://huggingface.co/microsoft/Phi-3-vision-128k-instruct for more info """
    err: bool = False

    # check if readonly
    image_file_path, readonly, msg = check_readonly_formats(image_file_path, extension, verbose)

    if readonly:
        return msg, False

    #image_file_path = convert_img_to_jpg(image_file_path, verbose)
    img_base64_str = img_file_to_base64(image_file_path, settings)

    # prompt
    if for_notes:
        p = settings["prompt_notes"]
    else:
        p = settings["prompt"]

    prompt = f"<|user|>\n<|{img_base64_str}|>\n{p}<|end|>\n<|assistant|>\n"

    processor = settings['phi3-vision-tokenizer']
    model = settings['phi3-vision-model']

    messages = [ 
        {"role": "user", "content": prompt},
    ]

    prompt = processor.tokenizer.apply_chat_template(messages, tokenize=False, add_generation_prompt=True)
    inputs = processor(prompt, [img_base64_str], return_tensors="pt").to("cuda:0")

    generation_args = {
        "max_new_tokens": 500,
        "temperature": 0.0,
        "do_sample": False, 
    }

    generate_ids = model.generate(**inputs, eos_token_id=processor.tokenizer.eos_token_id, **generation_args)

    # remove input tokens 
    generate_ids = generate_ids[:, inputs['input_ids'].shape[1]:]
    response = processor.batch_decode(generate_ids, skip_special_tokens=True, clean_up_tokenization_spaces=False)[0]

    # capitalize
    alt_text:str = response # processed_text.strip().capitalize()

    return alt_text, err

def setup_mlx_vlm(settings: dict) -> bool:
    """ Setup MLX VLM """
    #model_path = f"mlx-community/{settings['model']}" # mlx-community/llava-1.5-7b-4bit
    model_path = settings['model']
    tokenizer_config = {
        "trust_remote_code": True
    }
    model, processor = load(model_path) #, tokenizer_config)

    settings['mlx-vlm-model'] = model
    settings['mlx-vlm-tokenizer'] = processor

    print(f"Model: '{settings['model']}'")
    print(f"Prompt: {settings['prompt']}")
    
    return False

def use_mlx_vlm(
        image_file_path: str,
        extension: str,
        settings: dict,
        for_notes: bool = False,
        verbose: bool = False,
        debug: bool = False
    ) -> Tuple[str, bool]:
    """ MLX VLM """
    err: bool = False
    alt_text: str = ""

    model = settings['mlx-vlm-model']
    processor = settings['mlx-vlm-tokenizer']

    if model is not None and processor is not None:

        # check if readonly
        image_file_path, readonly, msg = check_readonly_formats(image_file_path, extension, verbose)

        if readonly:
            return msg, False

        #image_file_path = convert_img_to_jpg(image_file_path, verbose)
        img_base64_str = img_file_to_base64(image_file_path, settings)

        if for_notes:
            the_prompt = f"<image>\n{settings['prompt_notes']}"
        else:
            the_prompt = f"<image>\n{settings['prompt']}"

        prompt = processor.tokenizer.apply_chat_template(
           [ {"role": "user", "content": the_prompt} ],
           tokenize=False,
           add_generation_prompt=True,
        )

        output = generate(model, processor, img_base64_str, prompt, verbose=True)

        #output = output.replace("\n", "")
        output = output.replace("</s>", "")
        alt_text = output

    else:
        err = True

    return alt_text, err
