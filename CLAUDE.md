# CEJ Master Spec Sheet

## Project Overview
Excel transformation tool for CEJ (Content & Editorial Journey) spec sheets. Converts master spec sheet data for various platforms and regions.

## Tech Stack
- Python (see pyproject.toml)
- Streamlit for UI
- openpyxl for Excel processing

## Key Files
- `streamlit_app.py` - Main application entry point
- `excel_transformer.py` - Core transformation logic
- `config.py` - Configuration settings
- `cej_transformer/` - Transformation module package

## Workflows
- Use `/start` to initialize daily session
- Track work in `docs/{date}/{date}.md`
- Update `PROJECT_STATUS.md` when priorities shift

## Environment
- Copy `.env.example` to `.env` and configure
- Install dependencies: `pip install -r requirements.txt`
