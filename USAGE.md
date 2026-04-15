# Vietnamese Lô Tô Ticket Generator - Usage Guide

## Command Line Usage

### Basic Commands

**Generate preview (1-page preview with default config):**
```bash
python loto_generator.py
```

**Generate full 1800 tickets with split by rounds (default):**
```bash
python loto_generator.py --preview-only false
```

**Generate full without splitting by rounds:**
```bash
python loto_generator.py --preview-only false --split-rounds false
```

**Generate preview only (explicit):**
```bash
python loto_generator.py --preview-only true
```

### Parameters

#### `-n, --tables TABLES`
Number of unique tables to generate
- Default: 1800 (from config) or 500
- Example: `-n 300` generates 300 tickets
- **Note:** When `--preview-only true`, this is automatically set to 1 regardless of `-n` value

#### `-p, --per-page PER_PAGE`
Number of tickets per page in DOCX output
- Default: from `config.json` (tickets_per_page)
- Example: `-p 5` shows 5 tickets per page

#### `-c, --config CONFIG`
Path to configuration JSON file
- Default: `config.json` in current directory
- Example: `-c config_7cols.json` uses alternate config

#### `--preview-only true|false`
Control preview vs full mode
- `true`: Generate 1-page preview file only (overrides `-n` parameter)
- `false`: Generate full file based on `-n` parameter
- Default: from `config.json` (`generate_preview` parameter)
- Example: `--preview-only false`

#### `--split-rounds true|false`
Control whether to split output by rounds
- `true`: Generate main DOCX + 6 round-based DOCX files (300 tickets each)
- `false`: Generate main DOCX file only
- Default: from `config.json` (`split_by_rounds` parameter)
- Example: `--split-rounds false`

#### `--format docx|pdf`
Export file format
- `docx`: Generate Microsoft Word DOCX format (default)
- `pdf`: Generate PDF format (requires LibreOffice to be installed)
- Default: from `config.json` (`export_format` parameter)
- Example: `--format pdf`
- **Note:** PDF conversion requires LibreOffice to be installed on the system

## Configuration File (config.json)

### New Parameters

```json
{
    "split_by_rounds": true,      // Enable/disable round splitting
    "generate_preview": true,     // Enable/disable preview mode
    "export_format": "docx"       // Export format: "docx" or "pdf"
}
```

- `split_by_rounds` (boolean, default: true)
  - When true: Generates 6 round files after main file generation
  - When false: Generates only main DOCX file

- `generate_preview` (boolean, default: true)
  - When true: Generates 1-page preview file (overrides total_tables)
  - When false: Generates full file with total_tables count

- `export_format` (string, default: "docx")
  - When "docx": Generates Microsoft Word format files
  - When "pdf": Generates PDF files (requires LibreOffice installed)
  - **Note:** PDF conversion requires LibreOffice to be installed

## Output Files

When generating full 1800 tickets with split_by_rounds=true:

1. **loto_tables.txt** - Visual text representation of all tickets
2. **loto_tables.json** - Structured data of all tickets in JSON format
3. **loto_tables_YYYYMMDD_HHMMSS.docx** - Main DOCX with all 1800 tickets
4. **loto_7cols_round_1.docx** through **loto_7cols_round_6.docx** - Round-based files (300 tickets each)

When using preview mode:
- Only 1 table is generated
- Same output files but with 1 ticket

## Examples

### Example 1: Quick preview of current config
```bash
python loto_generator.py
# or explicitly
python loto_generator.py --preview-only true
```
Output: 1 preview DOCX + 1 page preview

### Example 2: Full generation with rounds (default behavior)
```bash
python loto_generator.py --preview-only false
```
Output: 1800-ticket main DOCX + 6 round files (300 each)

### Example 3: Full generation without rounds
```bash
python loto_generator.py --preview-only false --split-rounds false
```
Output: 1800-ticket main DOCX only (no round files)

### Example 4: Generate 300 tickets with preview mode disabled
```bash
python loto_generator.py -n 300 --preview-only false --split-rounds false
```
Output: 300-ticket main DOCX only

### Example 5: Use alternate config (9 columns)
```bash
python loto_generator.py -c config_9cols.json --preview-only false
```
Output: Full generation using 9-column config

### Example 6: Generate PDF format (requires LibreOffice)
```bash
python loto_generator.py --format pdf
```
Output: PDF preview file + 6 round PDF files (if LibreOffice is installed)

### Example 7: Generate full DOCX (explicit)
```bash
python loto_generator.py --preview-only false --format docx
```
Output: Full 1800-ticket DOCX + 6 round DOCX files

## PDF Export

PDF export requires LibreOffice to be installed on your system:

**Windows:**
```bash
# Install from https://www.libreoffice.org/download/
# Or use package manager
choco install libreoffice-fresh  # If using Chocolatey
```

**macOS:**
```bash
brew install libreoffice
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install libreoffice
```

Once installed, you can use `--format pdf` to generate PDF files instead of DOCX.

## Configuration Precedence

For each parameter, the precedence is:
1. **Command-line argument** (highest priority)
2. **Config file value**
3. **Built-in default** (lowest priority)

Example:
- If config.json has `export_format: pdf` but you run `--format docx`, DOCX format is used
- If you don't specify `--format`, it uses the value from config.json
- If config.json doesn't have `export_format`, it defaults to "docx"
