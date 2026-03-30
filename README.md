# VNLoto - Vietnamese Lô Tô (90-ball Bingo) Card Generator

## Overview

VNLoto is a Python-based generator for authentic Vietnamese Lô Tô (90-ball Bingo) cards with strict validation rules. It generates unique tables in batch, outputs to multiple formats (DOCX, JSON, text), and supports configurable styling and round assignment.

## Features

- **Authentic Lô Tô Rules**: 3 rows × 9 columns, 5 numbers per row, 15 total numbers per card, strict column range enforcement (1-9, 10-19, ..., 80-90), ascending order per column
- **Batch Generation**: Generate up to 560+ unique tables with no duplicates
- **Multiprocessing**: 8 worker processes for fast parallel generation
- **Round System**: Distribute tables across multiple rounds (default 3 rounds of ~187 tables each)
- **DOCX Output**: Professional Word document with:
  - Configurable grid layout (4 tables × 2 rows per page)
  - Round and table numbering
  - Exact cell dimensions (0.3" × 0.3")
  - Title row with config text (ALL CAPS, grey background)
  - Footer row with random motivational messages
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

### Generate with defaults (560 tables, 3 rounds, 8 tables per page)

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

- `loto_tables_YYYYMMDD_HHMMSS.docx` - Word document with formatted tables (70 pages for 560 tables)
- `loto_tables.json` - JSON array of all generated tables
- `loto_tables.txt` - Text representation of all tables

## Validation

All generated tables pass strict Lô Tô rules validation:
- Exactly 3 rows, 9 columns per card
- Exactly 5 numbers per row
- Exactly 15 unique numbers per card
- Numbers strictly in column ranges (1-9, 10-19, ..., 80-90)
- Numbers ascending within each column
- No duplicate tables across all generations