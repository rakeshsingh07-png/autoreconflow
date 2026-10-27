import io
import re
import polars as pl
import pandas as pd
import pdfplumber

def parse_to_polars(uploaded_file) -> pl.DataFrame:
    """Reads CSV, Excel (XLSX/XLS), or digital PDF statements into a clean Polars DataFrame."""
    filename = uploaded_file.name.lower()
    uploaded_file.seek(0)
    
    # 1. CSV File Parsing
    if filename.endswith(".csv"):
        content_bytes = uploaded_file.read()
        
        # Decode bytes with fallback
        try:
            text = content_bytes.decode("utf-8-sig")
        except UnicodeDecodeError:
            try:
                text = content_bytes.decode("utf-8")
            except UnicodeDecodeError:
                text = content_bytes.decode("latin1")
                
        # Parse using StringIO with python sniffer
        try:
            pdf = pd.read_csv(io.StringIO(text), sep=None, engine="python")
        except Exception:
            pdf = pd.read_csv(io.StringIO(text), sep=",")
            
        # Clean column headers
        pdf.columns = [
            str(col).strip().strip('"').strip("'").replace("\ufeff", "") 
            for col in pdf.columns
        ]
        
        # Ensure column names are strictly unique strings
        seen = {}
        unique_cols = []
        for c in pdf.columns:
            if c in seen:
                seen[c] += 1
                unique_cols.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                unique_cols.append(c)
        pdf.columns = unique_cols
        pdf = pdf.reset_index(drop=True)
        
        return pl.from_pandas(pdf)
        
    # 2. Excel Parsing
    elif filename.endswith((".xlsx", ".xls")):
        pdf = pd.read_excel(uploaded_file)
        pdf.columns = [str(col).strip() for col in pdf.columns]
        pdf = pdf.reset_index(drop=True)
        return pl.from_pandas(pdf)
        
    # 3. Digital PDF Statement Parsing
    elif filename.endswith(".pdf"):
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
        
    header_idx = 0
    keywords = ["date", "particulars", "description", "amount", "balance", "debit", "credit", "ref", "cheque", "chq"]
    for idx, row in enumerate(rows[:10]):
        row_str = " ".join(row).lower()
        if any(kw in row_str for kw in keywords):
            header_idx = idx
            break
            
    headers = [
        re.sub(r'[^a-zA-Z0-9_]', '', h.replace(" ", "_").lower()) or f"col_{i}" 
        for i, h in enumerate(rows[header_idx])
    ]
    data = rows[header_idx + 1:]
    valid_data = [r for r in data if len(r) == len(headers)]
    
    pdf = pd.DataFrame(valid_data, columns=headers).reset_index(drop=True)
    return pl.from_pandas(pdf)
