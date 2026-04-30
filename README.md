# IT Security Report Template Generator

Generates a `.docx` template that follows a strict internal IT security report specification with reusable Word styles, fixed color system, required section structure, and placeholder content.

## Install
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run
```bash
python document_builder.py
```

Output file:
- `output/it_security_report_template.docx`

## Variables supported in code (`ReportData`)
- `report_title`
- `site_name`
- `environment_name`
- `date`
- `classification`
- `site_labels`
- `section_content`
- `recommendation_rows`
- `status_value`
- `next_steps`
- `prepared_by`
