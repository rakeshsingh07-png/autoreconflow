import io
import re
import polars as pl
import pandas as pd
import pdfplumber

def parse_to_polars(uploaded_file) -> pl.DataFrame:
    """
    Reads CSV, Excel (XLSX/XLS), or digital PDF statements and converts
    them into a clean, multi-column Polars DataFrame.
    """
    filename = uploaded_file.name.lower()
    uploaded_file.seek(0)
    
    # 1. CSV File Parsing with delimiter sniffing
    if filename.endswith(".csv"):
        try:
            # Let pandas engine auto-detect delimiters (comma, tab, semicolon)
            pdf = pd.read_csv(uploaded_file, sep=None, engine="python")
        except Exception:
            uploaded_file.seek(0)
            try:
                pdf = pd.read_csv(uploaded_file, encoding="utf-8-sig")
            except Exception:
                uploaded_file.seek(0)
                pdf = pd.read_csv(uploaded_file, encoding="latin1")
                
        # Clean column headers: strip whitespace, surrounding quotes, and BOM markers
        pdf.columns = [
            str(col).strip().strip('"').strip("'").replace("\ufeff", "") 
            for col in pdf.columns
        ]
        return pl.from_pandas(pdf)
        
    # 2. Excel Parsing
    elif filename.endswith((".xlsx", ".xls")):
        pdf = pd.read_excel(uploaded_file)
        pdf.columns = [str(col).strip() for col in pdf.columns]
        return pl.from_pandas(pdf)
        
    # 3. Digital PDF Statement Parsing
    elif filename.endswith(".pdf"):
        return extract_pdf_table_to_polars(uploaded_file)
        
    else:
        raise ValueError(f"Unsupported file format: {uploaded_file.name}")

def extract_pdf_table_to_polars(pdf_file) -> pl.DataFrame:
    """
    Extracts tabular records from digital PDF bank statements.
    """
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
        
    # Locate table header row based on common financial transaction keywords
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
    
    pdf = pd.DataFrame(valid_data, columns=headers)
    return pl.from_pandas(pdf)
