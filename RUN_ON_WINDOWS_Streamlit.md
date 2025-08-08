# Run the Streamlit App on Windows (Latest Code)

This guide runs the latest Streamlit UI for the Excel Transformer on Windows using a Python virtual environment.

## 0) Prerequisites
- Windows 10/11
- Python 3.9+ installed and on PATH
  - Check: `python --version` or `py --version`
  - If Python isn’t found, install from https://www.python.org/downloads/ and re-open PowerShell.

## 1) Open PowerShell in the project folder
- Navigate to the project root (where `streamlit_app.py`, `requirements.txt` are located).

Example:
- Right-click the folder in Explorer → “Open in Terminal” (PowerShell).
- Or: Start → type “PowerShell” → Run as Administrator (optional) → `cd "C:\path\to\CEJ Master Spec Sheet"`

## 2) Create and Activate a Virtual Environment
PowerShell:
```
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

If you get an execution policy error:
```
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
.\.venv\Scripts\Activate.ps1
```

CMD (Command Prompt) alternative:
```
python -m venv .venv
.\.venv\Scripts\activate
```

## 3) Upgrade pip and install dependencies
PowerShell or CMD (with venv activated):
```
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Notes:
- `requirements.txt` includes: pandas, openpyxl, streamlit, Pillow.
- If `pip` is not found, use `python -m pip ...`.

## 4) Run the Streamlit app
Preferred invocation (avoids PATH issues):
```
python -m streamlit run streamlit_app.py
```

If you prefer and `streamlit` is on PATH:
```
streamlit run streamlit_app.py
```

What you should see:
- Streamlit will print a local URL (typically `http://localhost:8501`).
- A browser window may open automatically. If not, copy the URL into your browser.

## 5) Use the App
- Upload the input Excel file: `SSD_F&D_EG B2_V1_Haleon CEJ Master Spec Sheet 3.2_Master Template_25Jun.xlsx`.
- Click “Transform Excel Data”.
- Review logs and platform breakdowns in the UI.
- Download combined output per provided buttons (if data was generated).

## 6) Behavior Notes (current configuration)
- Funnel Stage “ALL” expansion is enabled by default:
  - Rows with Funnel Stage “ALL” are emitted as Awareness, Consideration, Purchase AFTER the row’s TOTAL validation passes.
- Presence reporting is QA-only by default (no placeholder rows mixed into transformed data yet). QA report generation is planned in the next phase.

## 7) Common Issues & Fixes
- “streamlit: command not found”
  - Ensure the venv is activated (`.\.venv\Scripts\Activate.ps1`) and use `python -m streamlit run streamlit_app.py`.
- “python not recognized”
  - Try `py -3 -m venv .venv` then `.\.venv\Scripts\Activate.ps1` and use `py -3 -m pip install -r requirements.txt` and `py -3 -m streamlit run streamlit_app.py`.
- Firewall prompt
  - Allow access so Streamlit can open the local server.
- Excel engine errors
  - `openpyxl` is included in requirements; re-run `pip install -r requirements.txt`.
- Permissions
  - If activation fails in PowerShell, use the `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` workaround above.

## 8) Stop and Deactivate
- Stop Streamlit: press `Ctrl+C` in the terminal where it’s running.
- Deactivate venv: `deactivate`

## 9) WSL Note (if running inside WSL)
- Use `python3` and `python3 -m streamlit run streamlit_app.py`.
- The app URL might be `http://localhost:8501` accessible from Windows browser if WSL networking is configured (default on WSL2).
