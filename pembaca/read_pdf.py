import argparse
import sys
from pypdf import PdfReader

def extract_text_from_pdf(pdf_path, output_path=None):
    """Extracts all text from the specified PDF file and prints it or writes to a file."""
    try:
        # Reconfigure standard output to prevent Windows console encoding errors
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
        if hasattr(sys.stderr, 'reconfigure'):
            sys.stderr.reconfigure(encoding='utf-8')
            
        reader = PdfReader(pdf_path)
        full_text = []
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                full_text.append(f"--- Page {i+1} ---\n{text}")
            else:
                full_text.append(f"--- Page {i+1} ---\n[No text content detected on this page]")
        
        content = "\n\n".join(full_text)
        if output_path:
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(content)
            print(f"Successfully extracted text to: {output_path}")
        else:
            print(content)
    except Exception as e:
        print(f"Error reading PDF: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract text from a PDF file.")
    parser.add_argument("pdf_path", help="Path to the PDF file to read.")
    parser.add_argument("-o", "--output", help="Optional path to save the extracted text.")
    args = parser.parse_args()
    extract_text_from_pdf(args.pdf_path, args.output)
