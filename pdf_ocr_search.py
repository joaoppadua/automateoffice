import pytesseract
from pdf2image import convert_from_path
import re
import sys
import os

def ocr_pdf_and_find_keywords(pdf_path, keywords):
    """OCR a PDF and find paragraphs mentioning specific keywords."""
    
    # Set Tesseract path (common Windows installation path)
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    
    # Check if PDF file exists
    if not os.path.exists(pdf_path):
        return None, f"Error: PDF file not found at {pdf_path}"
    
    try:
        print(f"Converting PDF to images...")
        pages = convert_from_path(pdf_path, 300)  # 300 DPI for good quality
        
        all_text = ""
        for i, page in enumerate(pages):
            print(f"Processing page {i+1}/{len(pages)}...")
            page_text = pytesseract.image_to_string(page)  # Default English
            all_text += f"\n--- PAGE {i+1} ---\n" + page_text
        
        # Split text into paragraphs (double newlines or significant spacing)
        paragraphs = re.split(r'\n\s*\n', all_text)
        
        # Prepare search terms (case-insensitive)
        search_terms = [keyword.lower() for keyword in keywords]
        
        # Find paragraphs mentioning the keywords
        matching_paragraphs = []
        for i, paragraph in enumerate(paragraphs):
            paragraph_lower = paragraph.lower()
            for term in search_terms:
                if term in paragraph_lower:
                    matching_paragraphs.append({
                        'paragraph_number': i+1,
                        'content': paragraph.strip(),
                        'matched_term': term
                    })
                    break  # Don't add the same paragraph multiple times
        
        return all_text, matching_paragraphs
    
    except Exception as e:
        return None, f"Error: {str(e)}"

def main():
    """Main function to handle user input and execute OCR search."""
    
    if len(sys.argv) < 3:
        print("Usage: python pdf_ocr_search.py <pdf_path> <keyword1> [keyword2] [keyword3] ...")
        print("Example: python pdf_ocr_search.py 'document.pdf' 'John Smith' 'contract' 'payment'")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    keywords = sys.argv[2:]
    
    print(f"Starting OCR for: {pdf_path}")
    print(f"Looking for keywords: {', '.join(keywords)}")
    print("=" * 60)

    full_text, results = ocr_pdf_and_find_keywords(pdf_path, keywords)

    if isinstance(results, str):  # Error occurred
        print(results)
        sys.exit(1)
    else:
        print(f"\nFound {len(results)} paragraphs mentioning the keywords:\n")
        
        for match in results:
            print(f"PARAGRAPH {match['paragraph_number']} (matched: '{match['matched_term']}'):")
            print("-" * 40)
            print(match['content'])
            print("\n" + "=" * 60 + "\n")
        
        if not results:
            print(f"No paragraphs found mentioning any of the keywords: {', '.join(keywords)}")
            print("\nFirst 1000 characters of OCR'd text for reference:")
            print(full_text[:1000] if full_text else "No text extracted")

if __name__ == "__main__":
    main()