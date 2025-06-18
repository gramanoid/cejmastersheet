# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

The CEJ Master Spec Sheet Transformer is a Python application that processes marketing campaign specification sheets for Haleon's Creative Effectiveness Journey (CEJ) tracking system. It transforms complex multi-platform marketing specifications from Excel format into standardized row-based output.

## Key Architecture

### Core Pipeline
1. **Platform Detection**: Searches for "Funnel Stage" markers to identify platform tables (YouTube, META, TikTok, etc.)
2. **Header Parsing**: Extracts main headers (Funnel Stage, Format, Duration) and sub-headers (aspect ratios, languages)
   - Dynamically identifies aspect ratio columns under merged "Aspect Ratio" header
   - Supports variable number of AR columns (now includes 4th column for YouTube)
3. **Data Transformation**: Converts matrix-style data to row-based format using:
   - Dual Language: (Sum of AR ticks) × (Count of selected languages) = Total combinations
   - Single Language: Sum of AR ticks = Total combinations
4. **Validation**: Ensures TOTAL column matches calculated combinations
5. **Output**: Generates timestamped Excel files with transformed data

### Key Files
- `excel_transformer.py` - Core transformation logic with platform detection and data processing
- `streamlit_app.py` - Web UI with file upload, real-time status, and download options
- `config.py` - Configuration constants (platform names, headers, row offsets)
- `validation_script.py` - Standalone validation tool for data integrity checks

## Common Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run command-line version
python excel_transformer.py

# Run web UI
streamlit run streamlit_app.py

# Run validation
python validation_script.py "input.xlsx" "output.xlsx"

# Build Windows executable
pyinstaller ExcelTransformer.spec
```

## Configuration

Key settings in `config.py`:
- `PLATFORM_NAMES`: Dictionary mapping platform identifiers to full names
- `START_ROW_SEARCH`: Row to begin searching for platform tables (default: 8)
- `OUTPUT_DATE_FORMAT`: Timestamp format for output files
- Platform table offsets: Title (0), Main headers (+2), Sub-headers (+3), Data (+4)

## Testing & Validation

- Test files: `Haleon CEJ Master Spec Sheet 3.1.xlsx` (240 combinations expected)
- Check `transformer.log` for detailed processing information
- Use `validation_script.py` to verify input/output totals match
- Web UI displays platform-specific breakdowns for verification

## Recent Updates (v2.3.0)

### New Features
- **Dynamic Aspect Ratio Columns**: Now supports 4 aspect ratio columns
  - YouTube: "IMG", "16x9", "9x16", "83 X 28 (1660 X 550)"
  - Other platforms: Show "n.a." in 4th column
- **Improved Error Handling**: Streamlit app handles missing header.png gracefully
- **Fixed Dependencies**: Added Pillow to requirements.txt for image handling

### Technical Details
- Platform detection uses "Funnel Stage" as anchor (2 rows below platform name in column B)
- Aspect ratio columns dynamically detected between "Aspect Ratio" and "Languages"/"TOTAL"
- Empty main header cells indicate merged "Aspect Ratio" header spans multiple columns
- Case-insensitive platform name matching in column B
- Logging overwrites `transformer.log` on each run

## Development Notes

- Current version: 2.3.0
- Branches: master (main), v2 (development)
- Tested with: V2_Haleon CEJ Master Spec Sheet 3.2_Master Template 1.xlsx