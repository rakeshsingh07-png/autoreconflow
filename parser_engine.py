import io
import re
import polars as pl
import pandas as pd
import pdfplumber

def parse_to_polars(uploaded_file) -> pl.DataFrame:
    """Reads CSV, Excel, or PDF and returns a clean, multi-column Polars DataFrame."""
    filename = uploaded_file.name.lower()
    
    if filename.endswith(".csv"):
        uploaded_file.seek(0)
        content = uploaded_file.read()
        # Handle string or bytes decoding
        if isinstance(content, bytes):
            # Try utf-8 first, fallback to latin1
            try:
                text_data = content.decode("utf-8")
            except UnicodeDecodeError:
                text_data = content.decode("latin1")
        else:
            text_data = str(content)
            
        # Wrap in io.StringIO to ensure Polars accurately detects the comma delimiter
        return pl.read_csv(io.StringIO(text_data), separator=",", ignore_errors=True)
        
    elif filename.endswith((".xlsx", ".xls")):
        uploaded_file.seek(0)
        # Read Excel via Pandas then convert to Polars
        pdf = pd.read_excel(uploaded_file)
        return pl.from_pandas(pdf)
        
    elif filename.endswith(".pdf"):
        uploaded_file.seek(0)
        return extract_pdf_table_to_polars(uploaded_file)
        
    else:
        raise ValueError(f"Unsupported file format: {uploaded_file.name}")

def extract_pdf_table_to_polars(pdf_file) -> pl.DataFrame:
    """Extracts tabular records from digital PDF bank statements."""
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    cleaned_row = [str(c).strip() if c is not None else "" for c in row]
                    if any(cleaned_row):
                        rows.append(cleaned_row)
                        
    if not rows:
        return pl.DataFrame()
        
    # Find table header row
    header_idx = 0
    keywords = ["date", "particulars", "description", "amount", "balance", "debit", "credit", "ref"]
    for idx, row in enumerate(rows[:10]):
        row_str = " ".join(row).lower()
        if any(kw in row_str for kw in keywords):
            header_idx = idx
            break
            
    headers = [re.sub(r'[^a-zA-Z0-9_]', '', h.replace(" ", "_").lower()) or f"col_{i}" 
               for i, h in enumerate(rows[header_idx])]
    data = rows[header_idx + 1:]
    valid_data = [r for r in data if len(r) == len(headers)]
    
    return pl.DataFrame(valid_data, schema=headers, orient="row")
