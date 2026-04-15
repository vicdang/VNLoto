# VNLoto - Vietnamese Lô Tô (90-ball Bingo) Ticket Generator

A powerful Python-based generator for authentic Vietnamese Lô Tô (90-ball Bingo) tickets with strict validation rules, batch processing, and multiple output formats.

## Features

### Core Functionality
- **Authentic Lô Tô Rules**
  - Configurable columns: 7-column (numbers 1-69) or 9-column (numbers 1-89)
  - Dynamic marked cells per row (configurable, default 4 for 7-col, 5 for 9-col)
  - Strict column range enforcement
  - Ascending order within each column
  - Multiple format support: DOCX, JSON, Text

- **Batch Generation**
  - Generate large batches of unique tickets (tested up to 1800+)
  - Duplicate detection via fingerprinting
  - Parallel processing with multiprocessing
  - Configurable table count and page layouts

- **Professional Document Output**
  - DOCX format with configurable styling
  - Grid layout (configurable tickets per row/page)
  - Round numbering
  - Title with winning condition info
  - Customizable footer messages
  - Border styling and special cell markers
  - PDF export (requires LibreOffice)

### Advanced Features
- **Preview Mode**: Generate 1-page preview for quick testing
- **Split by Rounds**: Automatically split output into separate files by rounds
- **Export Format Control**: DOCX (default) or PDF with fallback
- **Configuration-Driven**: All parameters in `config.json`
- **JSON Export**: Machine-readable format
- **Text Output**: ASCII visual representation for verification

## System Requirements

- Python 3.10+
- `python-docx` library
- LibreOffice (optional, for PDF export)

## Installation

### 1. Clone the Repository
```bash
git clone https://github.com/yourusername/VNLoto.git
cd VNLoto
```

### 2. Install Dependencies
```bash
pip install python-docx
```

### 3. (Optional) Install LibreOffice for PDF Support
**Windows:**
- Download from https://www.libreoffice.org/download/
- Or: `choco install libreoffice-fresh` (if using Chocolatey)

**macOS:**
```bash
brew install libreoffice
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install libreoffice
```

## Quick Start

### Generate preview (1 page):
```bash
python loto_generator.py
```

### Generate full 1800 tickets with rounds:
```bash
python loto_generator.py --preview-only false
```

### Generate as PDF (requires LibreOffice):
```bash
python loto_generator.py --preview-only false --format pdf
```

## Configuration

All parameters are configurable via `config.json`:

```json
{
    "title": "Your Title",
    "columns": 7,                    // 7 or 9
    "total_tables": 1800,           // Total tickets to generate
    "winning_cells": 4,              // Cells to win per row (7-col: 4, 9-col: 5)
    "rounds": 6,                     // Number of rounds to split into
    "page_layout": {
        "tickets_per_row": 5,       // Tickets per row
        "tickets_per_page": 10      // Tickets per page
    },
    "split_by_rounds": true,         // Split into round files
    "generate_preview": true,        // Generate preview (1 page)
    "export_format": "docx",         // "docx" or "pdf"
    "font": "Calibri",
    "font_size": 16,
    "header_font_size": 18,
    "footer_messages": ["Message 1", "Message 2", ...]
}
```

## Command-Line Options

### Basic Parameters

- **`-n, --tables TABLES`**
  - Number of tickets to generate
  - Default: from config.json
  - Example: `-n 300`

- **`-p, --per-page PER_PAGE`**
  - Tickets per page in output
  - Default: from config.json
  - Example: `-p 10`

- **`-c, --config CONFIG`**
  - Config file path
  - Default: `config.json`
  - Example: `-c config_7cols.json`

### Feature Flags

- **`--preview-only true|false`**
  - Generate preview only (1 page)
  - Default: from config.json
  - Example: `--preview-only false` (full generation)

- **`--split-rounds true|false`**
  - Split output by rounds
  - Default: from config.json
  - Example: `--split-rounds false` (single file)

- **`--format docx|pdf`**
  - Export format
  - Default: from config.json
  - Example: `--format pdf`

## Usage Examples

### Example 1: Quick Preview
```bash
python loto_generator.py
# Generates 1 preview page (1 ticket)
```

### Example 2: Full Generation (Default)
```bash
python loto_generator.py --preview-only false
# Generates 1800 tickets split into 6 round files
# Output: loto_tables_latest.docx + loto_7cols_round_1-6.docx
```

### Example 3: Full Generation Without Rounds
```bash
python loto_generator.py --preview-only false --split-rounds false
# Generates single file with all 1800 tickets
```

### Example 4: Custom Table Count
```bash
python loto_generator.py -n 300 --preview-only false --split-rounds false
# Generates 300 tickets in single file
```

### Example 5: Generate as PDF
```bash
python loto_generator.py --preview-only false --format pdf
# Requires LibreOffice installed
```

### Example 6: Use 9-Column Config
```bash
python loto_generator.py -c config_9cols.json --preview-only false
# Uses 9-column format (numbers 1-89)
```

## Output Files

### Main Output
- **`loto_tables_latest.docx`** - Latest generated file with all tickets
- **`loto_7cols_round_1.docx` through `loto_7cols_round_6.docx`** - Round-based files (if split_by_rounds=true)

### Supporting Files
- **`loto_tables.json`** - Machine-readable ticket data
- **`loto_tables.txt`** - ASCII preview of all tickets

### Generated During Execution (Temporary)
- Generated files are listed with timestamps during generation
- Only latest files are kept when pushing to repository

## Configuration Files

### Main Configuration
- **`config.json`** - Primary configuration (7-column format)
  - Used for DC34 Trip 2026 event
  - 1800 tickets, 6 rounds
  - 4 winning cells per row

### Alternative Configurations
- **`config_7cols.json`** - 7-column template
- **`config_9cols.json`** - 9-column template

## Project Structure

```
VNLoto/
├── loto_generator.py          # Main generator script
├── split_by_rounds.py         # Round file splitter
├── config.json                # Primary configuration
├── config_7cols.json          # 7-column config template
├── config_9cols.json          # 9-column config template
├── USAGE.md                   # Detailed usage guide
├── README.md                  # This file
├── .gitignore                 # Git ignore rules
├── LICENSE                    # License
└── loto_tables_latest.docx    # Latest generated output
    loto_7cols_round_*.docx    # Round-based outputs
```

## How It Works

### Ticket Generation Algorithm

1. **Column Definition**: Numbers are strictly assigned to columns
   - 7-col: [1-9, 10-19, 20-29, 30-39, 40-49, 50-59, 60-69]
   - 9-col: [1-9, 10-19, 20-29, 30-39, 40-49, 50-59, 60-69, 70-79, 80-89]

2. **Card Generation**: For each card (3 rows):
   - Randomly decide how many numbers per column (1, 2, or 3)
   - Total marked cells = `winning_cells × 3` (e.g., 4 × 3 = 12 for 7-col)
   - Distribute numbers to rows ensuring each row has exactly `winning_cells` marked
   - Use greedy row-assignment algorithm for efficiency

3. **Validation**: Each card is validated:
   - Correct number of cells per row
   - Correct total number of cells
   - No duplicates
   - Numbers in correct column ranges
   - Ascending order within columns

4. **Uniqueness**: Tables fingerprinted to prevent duplicates

5. **Output**: Generated to DOCX, JSON, and TXT formats

## Performance

- **Generation Speed**: ~300-500 tickets per second (varies by hardware)
- **1800 Tickets**: ~5-10 seconds with multiprocessing
- **Memory Usage**: ~50-100 MB for full generation
- **DOCX File Size**: ~1.7 MB for 1800 tickets

## Troubleshooting

### PDF Conversion Fails
**Issue**: "LibreOffice not found" warning
**Solution**: Install LibreOffice (see Installation section)

### Generator Hangs
**Issue**: Process seems stuck
**Solution**: Check that `config.json` exists and is valid JSON

### Validation Failures
**Issue**: Some tickets fail validation
**Solution**: Check that `winning_cells` matches expected value

## Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues, questions, or suggestions, please open an issue on GitHub.

## Changelog

### v2.0 (April 15, 2026)
- Added preview mode (`--preview-only`)
- Added split by rounds feature (`--split-rounds`)
- Added PDF export support (`--format`)
- Dynamic winning cells configuration
- Parameter precedence system
- Comprehensive documentation

### v1.0 (Initial)
- Basic ticket generation
- DOCX output
- Round distribution
- Configuration support

## Credits

Developed for DC34 Trip 2026 event. Vietnamese Lô Tô is a traditional 90-ball bingo game.
