# IT Security Report Builder

Interactive script that asks you questions and generates either:
- a Word IT security report template (`.docx`), or
- an Excel report worksheet (`.xlsx`).

## Install
```bash
python -m pip install -r requirements.txt
```

## Run
```bash
python document_builder.py
```

The script prompts for report title, site/environment info, status, recommendation rows, and next steps.
Generated files are written to `output/`.
