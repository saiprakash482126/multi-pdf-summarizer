import os
import nltk
from transformers import pipeline
from pathlib import Path
import time
from datetime import datetime
import requests
from msal import ConfidentialClientApplication
import json
from urllib.parse import urlparse, parse_qs
from dotenv import load_dotenv
import tempfile

# Load environment variables
load_dotenv()

# Configuration
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID = os.getenv('TENANT_ID')
SHAREPOINT_SITE_URL = os.getenv('SHAREPOINT_SITE_URL')
DRIVE_ID = os.getenv('DRIVE_ID')

# Microsoft Graph API endpoints
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
SCOPES = ['https://graph.microsoft.com/.default']
GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0'

class SharePointSummarizer:
    def __init__(self):
        """Initialize the summarizer and authentication."""
        self.app = ConfidentialClientApplication(
            CLIENT_ID,
            authority=AUTHORITY,
            client_credential=CLIENT_SECRET
        )
        self.access_token = self._get_access_token()
        self.headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        # Initialize NLTK and summarization model
        nltk.download('punkt', quiet=True)
        print("Loading the summarization model (this may take a minute)...")
        self.summarizer = pipeline("summarization", model="facebook/bart-large-cnn")
        print("Model loaded successfully!")
    
    def _get_access_token(self):
        """Get access token using client credentials flow."""
        result = None
        result = self.app.acquire_token_silent(SCOPES, account=None)
        
        if not result:
            print("Getting new token...")
            result = self.app.acquire_token_for_client(scopes=SCOPES)
        
        if "access_token" in result:
            return result['access_token']
        else:
            print(result.get("error"))
            print(result.get("error_description"))
            raise Exception("Could not get access token")
    
    def _make_graph_request(self, endpoint, method='GET', json_data=None):
        """Make a request to Microsoft Graph API."""
        url = f"{GRAPH_ENDPOINT}{endpoint}"
        
        try:
            if method.upper() == 'GET':
                response = requests.get(url, headers=self.headers)
            elif method.upper() == 'POST':
                response = requests.post(url, headers=self.headers, json=json_data)
            
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.RequestException as e:
            print(f"Error making request to {url}: {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"Response: {e.response.text}")
            return None
    
    def extract_drive_and_item_id(self, sharepoint_url):
        """Extract drive and item ID from SharePoint URL."""
        try:
            # If it's a sharing link, we need to get the drive and item ID
            if 'sharepoint.com/:f:/' in sharepoint_url or 'sharepoint.com/:f:/' in sharepoint:
                # Extract the sharing token
                parsed = urlparse(sharepoint_url)
                params = parse_qs(parsed.fragment)
                sharing_token = params.get('id', [''])[0]
                
                # Get the drive item from the sharing URL
                result = self._make_graph_request(f'/shares/u!{sharing_token}/driveItem')
                if result:
                    return result.get('parentReference', {}).get('driveId'), result.get('id')
            
            return None, None
            
        except Exception as e:
            print(f"Error extracting drive and item ID: {str(e)}")
            return None, None
    
    def list_files_in_folder(self, folder_id=None):
        """List all files in the specified folder or root."""
        endpoint = f"/drives/{DRIVE_ID}/items/{folder_id}/children" if folder_id else f"/drives/{DRIVE_ID}/root/children"
        response = self._make_graph_request(endpoint)
        return response.get('value', []) if response else []
    
    def download_file(self, item_id, file_name):
        """Download a file from OneDrive/SharePoint."""
        try:
            # Create a temporary directory
            temp_dir = os.path.join(tempfile.gettempdir(), 'doc_summarizer')
            os.makedirs(temp_dir, exist_ok=True)
            
            # Download the file
            endpoint = f"/drives/{DRIVE_ID}/items/{item_id}/content"
            url = f"{GRAPH_ENDPOINT}{endpoint}"
            
            with requests.get(url, headers=self.headers, stream=True) as r:
                r.raise_for_status()
                file_path = os.path.join(temp_dir, file_name)
                with open(file_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
            
            return file_path
            
        except Exception as e:
            print(f"Error downloading file {file_name}: {str(e)}")
            return None
    
    def read_text_file(self, file_path: str) -> str:
        """Read text from a file with proper encoding handling."""
        try:
            encodings = ['utf-8', 'latin-1', 'cp1252']
            for encoding in encodings:
                try:
                    with open(file_path, 'r', encoding=encoding) as f:
                        return f.read()
                except UnicodeDecodeError:
                    continue
            raise Exception(f"Could not read file with any of the supported encodings: {encodings}")
        except Exception as e:
            print(f"Error reading file {file_path}: {str(e)}")
            return ""
    
    def chunk_text(self, text: str, max_length: int = 1024) -> list:
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
                
            try:
                summary = self.summarizer(
                    chunk,
                    max_length=150,
                    min_length=40,
                    do_sample=False,
                    truncation=True
                )
                summaries.append(summary[0]['summary_text'])
            except Exception as e:
                print(f"Error summarizing text chunk: {str(e)}")
                continue
        
        return ' '.join(summaries) if summaries else "[No summary generated]"
    
    def process_sharepoint_folder(self, folder_url, output_folder="summaries"):
        """Process all text files in the specified SharePoint folder."""
        # Create output directory if it doesn't exist
        os.makedirs(output_folder, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Extract drive and item ID from URL
        drive_id, item_id = self.extract_drive_and_item_id(folder_url)
        if not drive_id or not item_id:
            print("Could not extract drive and item ID from the provided URL.")
            return
        
        # Get folder contents
        endpoint = f"/drives/{drive_id}/items/{item_id}/children"
        response = self._make_graph_request(endpoint)
        
        if not response or 'value' not in response:
            print("No files found in the specified folder.")
            return
        
        processed_count = 0
        
        for item in response['value']:
            if 'file' in item and item['name'].lower().endswith(('.txt', '.pdf', '.docx')):
                print(f"\nProcessing: {item['name']}")
                
                # Download the file
                file_path = self.download_file(item['id'], item['name'])
                if not file_path:
                    print(f"Skipping {item['name']} - could not download")
                    continue
                
                # Read and summarize the file
                try:
                    text = self.read_text_file(file_path)
                    if not text.strip():
                        print(f"Skipping {item['name']} - empty or could not read content")
                        continue
                        
                    print(f"Summarizing {item['name']}...")
                    start_time = time.time()
                    
                    summary = self.summarize_text(text)
                    
                    # Save the summary
                    output_filename = f"{Path(item['name']).stem}_summary_{timestamp}.txt"
                    output_path = os.path.join(output_folder, output_filename)
                    
                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(f"Summary of: {item['name']}\n")
                        f.write(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                        f.write("-" * 50 + "\n\n")
                        f.write(summary)
                    
                    processed_count += 1
                    elapsed = time.time() - start_time
                    print(f"✓ Summary saved to: {output_path} (took {elapsed:.1f} seconds)")
                    
                except Exception as e:
                    print(f"✗ Error processing {item['name']}: {str(e)}")
                
                # Clean up the downloaded file
                try:
                    os.remove(file_path)
                except:
                    pass
        
        print(f"\nProcessing complete! {processed_count} files were summarized.")
        print(f"All summaries were saved in the '{output_folder}' directory.")

def main():
    print("=" * 60)
    print("SharePoint/OneDrive Document Summarizer".center(60))
    print("=" * 60)
    
    # Check if environment variables are set
    required_vars = ['CLIENT_ID', 'CLIENT_SECRET', 'TENANT_ID']
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        print("\nError: Missing required environment variables:")
        for var in missing_vars:
            print(f"- {var}")
        print("\nPlease update the .env file with your Azure AD app credentials.")
        return
    
    # Get the SharePoint folder URL from the user
    sharepoint_url = input("\nEnter the SharePoint/OneDrive folder URL: ").strip()
    
    if not sharepoint_url:
        print("No URL provided. Exiting...")
        return
    
    print("\nStarting to process files...")
    print("This may take some time depending on the number and size of files...\n")
    
    # Create and run the summarizer
    summarizer = SharePointSummarizer()
    summarizer.process_sharepoint_folder(sharepoint_url)

if __name__ == "__main__":
    main()
