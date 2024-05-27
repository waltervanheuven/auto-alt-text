"""
Generate Alt Text for each picture in a powerpoint file using MLLM and V-L pre-trained models
"""
import sys
from .process import process_pptx

if __name__ == "__main__":
    sys.exit(process_pptx())
