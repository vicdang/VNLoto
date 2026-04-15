#!/usr/bin/env python3
"""Split generated tickets into separate files by round."""

import json
import os
from docx import Document
from docx.shared import Pt, Cm, Mm, Inches, RGBColor, Emu
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# Import the save_docx function from loto_generator
import sys
sys.path.insert(0, os.path.dirname(__file__))
from loto_generator import save_docx, init_column_ranges, convert_docx_to_pdf

def split_rounds(num_rounds=6, total_tables=1800, config_path="config.json"):
    """Split generated tables into separate files by round."""
    
    # Load config
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)
    
    # Initialize column ranges from config
    num_cols = config.get("columns", 9)
    init_column_ranges(num_cols)
    print(f"Initialized {num_cols} columns")
    
    # Get export format from config
    export_format = config.get("export_format", "docx")
    if export_format not in ['docx', 'pdf']:
        export_format = "docx"
    
    # Load all generated tables
    json_path = "loto_tables.json"
    if not os.path.exists(json_path):
        print(f"Error: {json_path} not found. Run loto_generator.py first.")
        return
    
    with open(json_path, "r", encoding="utf-8") as f:
        all_tables_data = json.load(f)
    
    # Extract just the grids from the table data
    all_tables = [table_obj['grid'] for table_obj in all_tables_data]
    
    print(f"Loaded {len(all_tables)} tables from {json_path}")
    
    tables_per_round = total_tables // num_rounds
    
    # Create a separate DOCX file for each round
    for round_num in range(1, num_rounds + 1):
        start_idx = (round_num - 1) * tables_per_round
        end_idx = round_num * tables_per_round
        
        round_tables = all_tables[start_idx:end_idx]
        
        # Update config for this round
        round_config = config.copy()
        new_title = f"{config.get('title', 'LÔ TÔ')} - Round {round_num}"
        round_config["title"] = new_title
        
        # Generate DOCX for this round
        output_path = f"loto_7cols_round_{round_num}.docx"
        save_docx(round_tables, output_path, round_config, tables_per_page=10, round_num=round_num)
        print(f"[OK] Round {round_num}: {len(round_tables)} tables >> {output_path}")
        
        # Convert to PDF if requested
        if export_format == "pdf":
            pdf_path = convert_docx_to_pdf(output_path)
            if pdf_path:
                print(f"     PDF version  >> {pdf_path}")

if __name__ == "__main__":
    split_rounds(num_rounds=6, total_tables=1800, config_path="config.json")
    print("\nAll rounds split successfully!")
