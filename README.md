
# 📘 Multi-PDF Summarizer — End-to-End PDF Summarization System
### 🔥 AI-Powered PDF Summaries | Streamlit UI | FastAPI | Docker | Local LLM Support

[![Python](https://img.shields.io/badge/Python-3.9+-blue)]()
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)]()
[![Build Status](https://img.shields.io/badge/Status-Active-success)]()
[![HuggingFace Models](https://img.shields.io/badge/LLM-BART%20%7C%20LLaMA%20%7C%20Mistral-orange)]()

---

## 🚀 Overview
**Multi-PDF Summarizer** is a complete end-to-end system that allows users to:
- Upload multiple PDF files
- Automatically extract text
- Process long documents using chunking
- Generate accurate, concise summaries using local or cloud LLMs
- View summaries using a clean Streamlit UI
- Use an optional FastAPI backend for automation
- Deploy anywhere using Docker

This project is ideal for students, researchers, corporates, legal firms, medical documentation, report automation, and data analysts.

---


<img width="829" height="581" alt="pdf summarize" src="https://github.com/user-attachments/assets/8c8d97c1-97f7-4330-a8f9-19837d54679d" />



A basic Python script that summarizes text files using a pre-trained BART model from Hugging Face.

## Setup

1. Install Python 3.7 or higher
2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

## How to Use

### Summarize a single file:
```
python simple_summarizer.py path/to/your/file.txt
```

### Summarize all .txt files in a directory:
```
python simple_summarizer.py path/to/your/directory
```

## Example

1. Create a file called `sample.txt` with some text
2. Run: `python simple_summarizer.py sample.txt`
3. The summary will be displayed in the console

## Notes

- The first run will download the BART model (about 1.5GB)
- Works best with English text
- Only processes .txt files

## Requirements

- Python 3.7+
- transformers
- torch
- nltk
