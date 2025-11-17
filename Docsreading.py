import sys
try:
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    print("Error: 'python-docx' library not found.")
    print("Please install it by running: pip install python-docx")
    sys.exit(1)

def extract_docx_content(filepath):
    """
    Extracts and prints text and table content from a .docx file
    in sequential order as they appear in the document.
    """
    try:
        doc = Document(filepath)
        
        print(f"\n--- [START] Extracting content from: {filepath} ---")
        
        # We iterate through the document's body's inner content.
        # This (private) list holds Paragraph and Table objects in 
        # the order they appear in the document. This is the key
        # to maintaining the correct sequence.
        for block in doc._body._inner_content:
            
            if isinstance(block, Paragraph):
                # We check if the paragraph is empty (e.g., just whitespace)
                # before printing, to keep the output clean.
                if block.text.strip():
                    print(f"\n[Paragraph]:")
                    print(block.text)
            
            elif isinstance(block, Table):
                print(f"\n[Table (Rows: {len(block.rows)}, Cols: {len(block.columns)})]:")
                
                # Iterate through all rows in the table
                for i, row in enumerate(block.rows):
                    row_data = []
                    for cell in row.cells:
                        # Get cell text. Replace newlines within a cell 
                        # with a space for cleaner, single-line-per-row output.
                        cell_text = cell.text.replace("\n", " ").strip()
                        row_data.append(cell_text)
                    
                    # Join all cell data in the row with ' | ' for a clear view
                    print(f"  Row {i+1}: | {' | '.join(row_data)} |")
        
        print(f"\n--- [END] Extraction complete for: {filepath} ---")

    except IOError:
        print(f"Error: File not found at '{filepath}'", file=sys.stderr)
    except Exception as e:
        # Catch other potential errors (e.g., file is not a docx, is corrupt)
        print(f"An error occurred: {e}", file=sys.stderr)
        print("Please ensure the file is a valid .docx file and you have permissions to read it.", file=sys.stderr)

if __name__ == "__main__":
    # Check if a file path was provided as a command-line argument
    if len(sys.argv) < 2:
        print("Usage: python extract_docx_sequentially.py <path_to_your_file.docx>")
        sys.exit(1)
        
    file_path = sys.argv[1]
    
    # Check if the file is a .docx file
    if not file_path.endswith('.docx'):
        print(f"Error: The file '{file_path}' is not a .docx file.", file=sys.stderr)
        sys.exit(1)
        
    extract_docx_content(file_path)

