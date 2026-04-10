# VNLoto - Vietnamese Lô Tô (90-ball Bingo) Card Generator

## Overview

VNLoto is a Python-based generator for authentic Vietnamese Lô Tô (90-ball Bingo) cards with strict validation rules. It generates unique tables in batch, outputs to multiple formats (DOCX, JSON, text), and supports configurable styling and round assignment.

## Features

- **Authentic Lô Tô Rules**: Configurable columns (default 9), 3 rows per card, exactly 5 numbers per row, 15 total numbers per card, strict column range enforcement, ascending order per column
- **Batch Generation**: Generate large batches of unique tables with no duplicates (for example 1200 tables)
- **Multiprocessing**: Uses available CPU cores for fast parallel generation
- **Round System**: Distribute tables across configurable rounds (for example 6 rounds for 1200 tables)
- **DOCX Output**: Professional Word document with:
  - Configurable grid layout (4 tables × 2 rows per page)
  - Round and table numbering
  - Exact cell dimensions (0.3" × 0.3")
  - Title row with config text (ALL CAPS, grey background)
  - Footer row split into 2 cells across 9 columns:
    - Message cell spans 6 columns and is center-aligned
    - Table ID cell spans 3 columns and is center-aligned with 6-digit format (`000001` ... `001200`)
  - Thick outer borders on each table, thin interior borders
  - Special cells (★ character) with optional replacement
- **JSON Export**: Machine-readable format for data processing
- **Text Output**: ASCII visual representation for quick verification
- **Config-Driven**: All styling parameters in `config.json` (fonts, sizes, colors, messages, borders)

## Requirements

- Python 3.10+
- python-docx

## Installation

```bash
pip install python-docx
```

## Usage

### Generate with defaults from `config.json` (for example 1200 tables, 6 rounds, 8 tables per page)

```bash
python loto_generator.py
```

### Generate with custom table count

```bash
python loto_generator.py -n 100
```

### Generate with custom tables per page

```bash
python loto_generator.py -n 560 -p 8
```

### Generate with custom config file

```bash
python loto_generator.py -c /path/to/config.json
```

### Options

- `-n, --tables`: Number of tables to generate (default from config.json)
- `-p, --per-page`: Tables per page in DOCX (default 8)
- `-c, --config`: Path to config file (default ./config.json)

## Configuration

Edit `config.json` to customize output. See config.json for all available options.

## Output Files

- `loto_tables_YYYYMMDD_HHMMSS.docx` - Word document with formatted tables
- `loto_tables.json` - JSON array of all generated tables
- `loto_tables.txt` - Text representation of all tables

## Validation

All generated tables pass strict Lô Tô rules validation (default 9 columns):
- Exactly 3 rows, configurable columns per card (default 9)
- Exactly 5 numbers per row
- Exactly 15 unique numbers per card
- Numbers strictly in column ranges
- Numbers ascending within each column
- No duplicate tables across all generations

## Configuration Reference

Key configuration options in `config.json`:
- `columns`: Number of columns per ticket (default: 9)
- `rounds`: Number of rounds to divide tables into (default: 6)
- `total_tables`: Total number of unique tables to generate (default: 1200)
- `title`: Document title (default: "DC34 Trip 2026")
- `font`: Font name for table content (default: "Calibri")
- `header_font_size`: Title row font size (default: 18)
- `special_cells`: Max special cells per row (default: 3)
- `footer_messages`: Array of random messages for table footers
- `table_border`: Border style/size/color for outer table edges

See `config.json` for complete documentation of all parameters.