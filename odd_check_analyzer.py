import os
import re
import spacy
from pathlib import Path
import docx
from typing import Dict, List, Set, Tuple, Optional, Any

try:
    nlp = spacy.load("en_core_web_sm")
except OSError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "spacy", "download", "en_core_web_sm"])
    nlp = spacy.load("en_core_web_sm")

# Common words/patterns to exclude from results
EXCLUDED_TERMS = {
    'llc', 'inc', 'ltd', 'co', 'corp', 'corporation', 'limited', 'usa', 'email', 'fax', 'telephone',
    'module', 'appendix', 'section', 'sequence', 'request', 'they', 'the', 'and', 'for', 'with',
    'cover letter', 'electronic submissions gateway', 'orange book', 'drug product', 'reference listed drug',
    'pre', 'strengths', 'new york', 'beltsville', 'june', 'march', 'january', 'february', 'march', 'april',
    'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'
}

def is_valid_entity(text: str) -> bool:
    """Check if the entity text is valid and not in excluded terms."""
    if not text or len(text) < 2:
        return False
    
    # Remove common prefixes/suffixes and normalize
    text = re.sub(r'^[^\w]+|[^\w]+$', '', text.lower())
    
    # Check against excluded terms
    if text in EXCLUDED_TERMS or any(term in text for term in EXCLUDED_TERMS):
        return False
    
    # Exclude dates and numbers
    if re.match(r'^\d+[/-]\d+[/-]\d+$', text):  # Matches dates like 06/10/2021
        return False
    if re.match(r'^\d+[A-Za-z]*$', text):  # Matches numbers with optional letters (e.g., 356h)
        return False
    
    # Exclude email addresses and phone numbers
    if '@' in text or re.match(r'^[\d\s-]+\(?:\)\s*)?$', text):
        return False
    
    return True

def extract_pharmaceutical_entities(doc) -> Dict[str, Set[str]]:
    """Extract pharmaceutical-related entities with improved filtering."""
    entities = {
        'manufacturers': set(),
        'products': set(),
        'substances': set(),
        'references': set()  # For regulatory references
    }
    
    # Common pharmaceutical company suffixes (case insensitive)
    company_suffixes = {
        'pharma', 'pharmaceuticals', 'laboratories', 'labs', 'healthcare',
        'biotech', 'therapeutics', 'sciences', 'pharm', 'pharma ltd',
        'pharmaceuticals inc', 'pharmaceuticals llc'
    }
    
    # Common substance patterns
    substance_patterns = {
        'ine$', 'ate$', 'ide$', 'one$', 'ol$', 'ene$', 'ium$', 'gen$', 'xide$', 'acid$',
        'amine$', 'zole$', 'caine$', 'sulfa', 'cycline$', 'mycin$', 'dipine$', 'pril$',
        'sartan$', 'lol$', 'prazole$', 'tidine$', 'zosin$', 'triptan$', 'oxetine$',
        'oxetine$', 'tadine$', 'dine$', 'vir$', 'mab$', 'ximab$', 'zumab$', 'mab$',
        'tide$', 'grastim$', 'kinra$', 'nib$', 'caine$', 'cillin$', 'barbital$'
    }
    
    # Process named entities
    for ent in doc.ents:
        text = ent.text.strip()
        if not is_valid_entity(text):
            continue
            
        # Check for manufacturers (ORG entities with company-like names)
        if ent.label_ == 'ORG':
            # Check for company name patterns
            if any(suffix in text.lower() for suffix in company_suffixes) or \
               any(word.istitle() for word in text.split() if len(word) > 3):
                entities['manufacturers'].add(text)
        
        # Check for substances (CHEM entities or chemical-looking names)
        elif ent.label_ == 'CHEM' or any(re.search(p, text.lower()) for p in substance_patterns):
            if len(text) > 3:  # Filter out very short potential false positives
                entities['substances'].add(text)
    
    # Process noun chunks for additional product names and substances
    for chunk in doc.noun_chunks:
        text = chunk.text.strip()
        if not is_valid_entity(text):
            continue
            
        # Check for product names (title case, multiple words, not too long)
        words = text.split()
        if (text.istitle() or (len(words) > 1 and any(w[0].isupper() for w in words))) and \
           len(text) < 50 and not any(w.isdigit() for w in words):
            # Check if it's likely a product name (not a common phrase)
            if len(text) > 5 and not any(p in text.lower() for p in ['the', 'and', 'for', 'with']):
                entities['products'].add(text)
    
    # Process the document text for regulatory references
    for sent in doc.sents:
        text = sent.text.strip()
        if any(term in text.lower() for term in ['fda', 'nda', 'anda', 'orange book', 'therapeutic equivalence']):
            # Clean up the reference
            ref = re.sub(r'\s+', ' ', text).strip()
            if 10 < len(ref) < 200:  # Reasonable length for a reference
                entities['references'].add(ref)
    
    return entities

def analyze_document(file_path: str) -> Dict[str, Any]:
    """Analyze a document and extract pharmaceutical entities."""
    try:
        # Extract text based on file type
        if str(file_path).lower().endswith('.docx'):
            doc = docx.Document(file_path)
            text = '\n'.join(para.text for para in doc.paragraphs if para.text.strip())
        else:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
        
        if not text.strip():
            return {"status": "error", "message": "Document is empty"}
        
        # Process with spaCy
        doc = nlp(text)
        
        # Extract entities
        entities = extract_pharmaceutical_entities(doc)
        
        # Convert sets to sorted lists
        result = {k: sorted(v) for k, v in entities.items() if v}
        
        return {
            "status": "success",
            "file": os.path.basename(file_path),
            "entities": result,
            "text_sample": text[:500] + ("..." if len(text) > 500 else "")
        }
        
    except Exception as e:
        return {
            "status": "error",
            "file": os.path.basename(file_path) if 'file_path' in locals() else 'unknown',
            "error": str(e)
        }

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Analyze pharmaceutical documents for key entities.')
    parser.add_argument('file_path', help='Path to the document to analyze')
    parser.add_argument('--debug', action='store_true', help='Enable debug output')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.file_path):
        print(f"Error: File not found: {args.file_path}")
        sys.exit(1)
    
    result = analyze_document(args.file_path)
    
    print("\n" + "="*50)
    print(" Pharmaceutical Document Analysis")
    print("="*50)
    
    if result["status"] == "success":
        print(f"\n File: {result['file']}")
        
        # Print manufacturers
        if 'manufacturers' in result['entities'] and result['entities']['manufacturers']:
            print("\n Manufacturers/Companies:")
            for i, mfg in enumerate(result['entities']['manufacturers'], 1):
                print(f"  {i}. {mfg}")
        
        # Print products
        if 'products' in result['entities'] and result['entities']['products']:
            print("\n Products/Devices:")
            for i, product in enumerate(result['entities']['products'], 1):
                print(f"  {i}. {product}")
        
        # Print substances
        if 'substances' in result['entities'] and result['entities']['substances']:
            print("\n Active Substances/Ingredients:")
            for i, substance in enumerate(result['entities']['substances'], 1):
                print(f"  {i}. {substance}")
        
        # Print regulatory references in debug mode
        if args.debug and 'references' in result['entities'] and result['entities']['references']:
            print("\n Regulatory References:")
            for i, ref in enumerate(result['entities']['references'], 1):
                print(f"  {i}. {ref}")
        
        if args.debug and 'text_sample' in result:
            print("\n Document Text Sample:")
            print("-" * 50)
            print(result['text_sample'])
            print("-" * 50)
    else:
        print(f"\n Error: {result.get('error', 'Unknown error occurred')}")
    
    print("\n" + "="*50)
    print("Analysis complete. ")
    print("="*50 + "\n")