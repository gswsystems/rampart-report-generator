# Rampart Report Generator

Generate professional, branded Word documents (.docx) from [Rampart](https://www.gswsystems.com/products/rampart) firewall security audit exports. Design a template once with your company's branding, then produce consistent reports for every engagement.

```
Rampart JSON Export  +  Word Template  →  Branded Audit Report
```

## Features

- **Template-driven** — Use any `.docx` file as a template with `{{ placeholder }}` syntax
- **Preserves formatting** — Fonts, colours, styles, logos, headers, and footers are all retained
- **Dynamic tables** — Repeating data (findings, compliance results, etc.) automatically expands into table rows
- **70+ variables** — Risk scores, finding counts, compliance rates, segmentation data, and more
- **16 table datasets** — Findings, shadowed rules, compliance, lateral movement, duplicate objects, and more
- **CLI metadata** — Set client name, auditor, date, and confidentiality from the command line
- **Variable discovery** — List all available data for any JSON export with `--list-variables`

## Quick Start

### Install

```bash
pip install -r requirements.txt
```

### Generate a Starter Template

```bash
python create_template.py template.docx
```

This creates a ready-made template with all sections, placeholders, and table markers pre-configured. Open it in Word to customise the branding.

### Generate a Report

```bash
python rampart_report.py audit.json template.docx report.docx \
    --client "Acme Corporation" \
    --auditor "Jane Smith" \
    --company "SecureAudit Pty Ltd"
```

### List Available Variables

```bash
python rampart_report.py --list-variables audit.json
```

## How It Works

1. **Export** your firewall analysis results from Rampart as JSON.
2. **Design** a Word template with your branding and insert `{{ placeholders }}` where data should appear.
3. **Run** the generator — it replaces every placeholder with actual values and expands table markers into data rows.
4. **Review** the output and update the Table of Contents in Word.

### Placeholder Syntax

Use double curly braces anywhere in the document:

```
Client: {{ client_name }}
Risk Score: {{ risk_score }} (Grade: {{ risk_grade }})
Total Findings: {{ total_findings }}
```

### Table Markers

For repeating data, use markers in a Word table row:

| Rule | Severity | Description | Remediation |
|------|----------|-------------|-------------|
| `{{#findings}}` `{{ rule_name }}` | `{{ severity }}` | `{{ description }}` | `{{ remediation }}` `{{/findings}}` |

The marker row is replaced by one row per data item.

## Command-Line Options

```
python rampart_report.py <json_file> <template.docx> <output.docx> [options]
```

| Option | Description | Default |
|--------|-------------|---------|
| `--client NAME` | Client name | _(empty)_ |
| `--client-contact NAME` | Client contact person | _(empty)_ |
| `--auditor NAME` | Auditor name | _(empty)_ |
| `--company NAME` | Auditing company | _(empty)_ |
| `--title TEXT` | Report title | Firewall Security Audit Report |
| `--date YYYY-MM-DD` | Report date | Today |
| `--confidentiality LEVEL` | Confidentiality marking | CONFIDENTIAL |
| `--list-variables` | List all variables and tables for a JSON file | |

## Documentation

- **[Tutorial](TUTORIAL.md)** — Step-by-step guide walking through the entire workflow with detailed explanations
- **[User Guide](USER_GUIDE.md)** — Complete reference for all placeholders, tables, template design, and troubleshooting

## Requirements

- Python 3.8+
- [python-docx](https://python-docx.readthedocs.io/) >= 1.1.0
- [Jinja2](https://jinja.palletsprojects.com/) >= 3.1.0

## License

This project is licensed under the GNU General Public License v3.0 — see [LICENSE](LICENSE) for details.
