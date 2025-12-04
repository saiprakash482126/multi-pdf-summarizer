# Simple Text Summarizer

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
