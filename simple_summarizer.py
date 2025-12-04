import os
import nltk
from transformers import pipeline
from typing import List, Dict
import argparse

# Download required NLTK data
nltk.download('punkt', quiet=True)
nltk.download('punkt_tab')
class SimpleSummarizer:
    def __init__(self):
        """Initialize the summarizer with a pre-trained model."""
        print("Loading the summarization model (this may take a minute)...")
        self.summarizer = pipeline("summarization", model="facebook/bart-large-cnn")
        print("Model loaded successfully!")
    
    def read_text_file(self, file_path: str) -> str:
        """Read text from a .txt file."""
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                return file.read()
        except Exception as e:
            print(f"Error reading {file_path}: {str(e)}")
            return ""
    
    def chunk_text(self, text: str, max_length: int = 1024) -> List[str]:
        """Split text into chunks that are suitable for the summarization model."""
        sentences = nltk.sent_tokenize(text)
        chunks = []
        current_chunk = ""
        
        for sentence in sentences:
            if len(current_chunk) + len(sentence) + 1 <= max_length:
                current_chunk += " " + sentence
            else:
                chunks.append(current_chunk.strip())
                current_chunk = sentence
        
        if current_chunk:
            chunks.append(current_chunk.strip())
            
        return chunks
    
    def summarize_text(self, text: str) -> str:
        """Generate a summary of the input text."""
        if not text.strip():
            return ""
            
        chunks = self.chunk_text(text)
        summaries = []
        
        for chunk in chunks:
            if len(chunk.split()) < 10:  # Skip very short chunks
                continue
                
            summary = self.summarizer(
                chunk,
                max_length=130,
                min_length=30,
                do_sample=False,
                truncation=True
            )
            summaries.append(summary[0]['summary_text'])
        
        return ' '.join(summaries)
    
    def process_files(self, input_path: str) -> Dict[str, str]:
        """Process .txt files in the input directory."""
        if not os.path.exists(input_path):
            print(f"Error: The path '{input_path}' does not exist.")
            return {}
            
        summaries = {}
        
        if os.path.isfile(input_path) and input_path.endswith('.txt'):
            print(f"\nProcessing: {input_path}")
            text = self.read_text_file(input_path)
            if text:
                summary = self.summarize_text(text)
                summaries[os.path.basename(input_path)] = summary
        elif os.path.isdir(input_path):
            print(f"Processing files in: {input_path}")
            for filename in os.listdir(input_path):
                if filename.endswith('.txt'):
                    file_path = os.path.join(input_path, filename)
                    print(f"\nProcessing: {filename}")
                    text = self.read_text_file(file_path)
                    if text:
                        summary = self.summarize_text(text)
                        summaries[filename] = summary
        else:
            print("Error: Please provide a valid .txt file or directory containing .txt files")
            
        return summaries

def main():
    parser = argparse.ArgumentParser(description='Simple Text Summarizer')
    parser.add_argument('input_path', type=str, 
                       help='Path to a .txt file or directory containing .txt files')
    
    args = parser.parse_args()
    
    print("Simple Text Summarizer")
    print("=====================")
    
    summarizer = SimpleSummarizer()
    summaries = summarizer.process_files(args.input_path)
    
    if summaries:
        print("\n=== Summaries ===")
        for filename, summary in summaries.items():
            print(f"\nFile: {filename}")
            print("-" * 50)
            print(summary)
            print("\n" + "="*50)
    else:
        print("\nNo .txt files were processed. Please check your input path.")

if __name__ == "__main__":
    main()
