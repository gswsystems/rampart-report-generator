# Rampart Report Generator — User Guide

## Overview

The Rampart Report Generator creates Word documents (.docx) from Rampart JSON exports using custom templates. This allows security consultants to produce branded audit deliverables that match their company's document standards.

**Workflow:** Rampart JSON Export + Word Template = Branded Report

## Requirements

- Python 3.8 or later
- pip (Python package manager)

## Installation

```bash
cd rampart-report-generator
pip install -r requirements.txt
```

This installs two dependencies:
- **python-docx** — reads and writes Word documents
- **Jinja2** — template placeholder processing

## Quick Start

### 1. Export JSON from Rampart

In Rampart, load a firewall configuration and run the analysis. Then click **Export Results** in the sidebar (or press Ctrl+E) and choose **JSON** format. Save the file (e.g. `audit-results.json`).

### 2. Create a Word Template

Create a `.docx` file in Word (or any compatible editor) using your company's branding — logos, headers, footers, fonts, colours, styles. Insert placeholders where you want Rampart data to appear:

```
Firewall Security Audit Report
Client: {{ client_name }}
Date: {{ report_date }}

Total rules analysed: {{ total_rules }}
Rules with issues: {{ rules_with_issues }}
Risk Score: {{ risk_score }} / 100 (Grade: {{ risk_grade }})

Critical: {{ critical_count }}  |  High: {{ high_count }}  |  Medium: {{ medium_count }}  |  Low: {{ low_count }}
```

### 3. Generate the Report

```bash
python rampart_report.py audit-results.json template.docx output-report.docx \
    --client "Acme Corporation" \
    --auditor "Jane Smith" \
    --company "SecureAudit Pty Ltd"
```

The output file will be a copy of your template with all placeholders replaced by the actual analysis data.

## Command-Line Options

```
python rampart_report.py <json_file> <template.docx> <output.docx> [options]
```

| Option | Description |
|--------|-------------|
| `--client NAME` | Client name (populates `{{ client_name }}`) |
| `--client-contact NAME` | Client contact person |
| `--auditor NAME` | Auditor name |
| `--company NAME` | Auditing company name |
| `--title TEXT` | Report title (default: "Firewall Security Audit Report") |
| `--date YYYY-MM-DD` | Report date (default: today) |
| `--confidentiality LEVEL` | Confidentiality marking (default: CONFIDENTIAL) |
| `--list-variables` | List all available variables and table data, then exit |

### Listing Available Variables

To see every variable and table available for your specific JSON export:

```bash
python rampart_report.py --list-variables audit-results.json
```

This prints all variable names, their current values, and all available table datasets with column names.

## Template Placeholders

Placeholders use double curly braces: `{{ variable_name }}`. They can appear anywhere in the document — body text, headers, footers, and table cells.

### Report Metadata

These are set via command-line options:

| Placeholder | Description |
|-------------|-------------|
| `{{ report_date }}` | Report date |
| `{{ report_title }}` | Report title |
| `{{ client_name }}` | Client name |
| `{{ client_contact }}` | Client contact person |
| `{{ auditor_name }}` | Auditor name |
| `{{ auditor_company }}` | Auditing company name |
| `{{ confidentiality }}` | Confidentiality level |

### Analysis Summary

| Placeholder | Description | Example |
|-------------|-------------|---------|
| `{{ total_rules }}` | Total rules analysed | 247 |
| `{{ rules_with_issues }}` | Rules with at least one finding | 38 |
| `{{ compliance_rate }}` | Overall compliance rate (%) | 84.2 |
| `{{ config_type }}` | Configuration type | panorama |
| `{{ analysis_timestamp }}` | When analysis was run | 2026-03-14T10:30:00 |
| `{{ device_group_count }}` | Number of device groups | 3 |
| `{{ device_groups }}` | Comma-separated device group names | DG-Perth, DG-Sydney |

### Severity Counts

| Placeholder | Description |
|-------------|-------------|
| `{{ critical_count }}` | Number of critical findings |
| `{{ high_count }}` | Number of high-severity findings |
| `{{ medium_count }}` | Number of medium-severity findings |
| `{{ low_count }}` | Number of low-severity findings |
| `{{ total_findings }}` | Total findings across all severities |

### Risk Rating

| Placeholder | Description | Example |
|-------------|-------------|---------|
| `{{ risk_score }}` | Composite risk score (0–100) | 72 |
| `{{ risk_grade }}` | Letter grade (A–F) | C |
| `{{ best_practices_score }}` | Best practices component score | 65 |
| `{{ segmentation_score }}` | Segmentation component score | 78 |
| `{{ critical_issues }}` | Critical issue count (for risk calc) | 2 |
| `{{ high_risk_rules }}` | High-risk rule count | 8 |
| `{{ shadowed_rule_count }}` | Shadowed rules found | 3 |
| `{{ lateral_movement_paths }}` | Lateral movement paths found | 5 |

### Duplicate Objects

| Placeholder | Description |
|-------------|-------------|
| `{{ duplicate_address_count }}` | Duplicate address objects |
| `{{ duplicate_service_count }}` | Duplicate service objects |

### Segmentation

| Placeholder | Description |
|-------------|-------------|
| `{{ seg_score }}` | Segmentation score (0–100) |
| `{{ seg_grade }}` | Segmentation grade |
| `{{ seg_zone_count }}` | Number of zones |
| `{{ seg_allowed_pairs }}` | Allowed zone pairs |
| `{{ seg_blocked_pairs }}` | Blocked zone pairs |

### Best Practices

| Placeholder | Description |
|-------------|-------------|
| `{{ best_practices_overall_score }}` | Overall best practices score |
| `{{ best_practices_grade }}` | Best practices grade |

### Analyzer Counts

| Placeholder | Description |
|-------------|-------------|
| `{{ rule_expiry_count }}` | Expired + temporary rules |
| `{{ cleartext_rule_count }}` | Rules with cleartext protocols |
| `{{ geo_ip_unrestricted_count }}` | Unrestricted geo-IP rules |
| `{{ lateral_movement_count }}` | Lateral movement findings |
| `{{ stale_rule_count }}` | Stale rules detected |
| `{{ egress_risk_count }}` | Egress risk findings |
| `{{ decryption_gap_count }}` | Decryption policy gaps |

## Table Data

For repeating data (findings, shadowed rules, etc.), use table markers in Word tables.

### How Table Markers Work

1. Create a table in your Word template with the column headers you want
2. In the first data row, add `{{#table_name}}` at the start and `{{/table_name}}` at the end
3. Use `{{ column_name }}` placeholders in between for each column
4. The marker row is replaced by one row per data item

### Example: Findings Table

Create this table in your Word template:

| Rule | Severity | Type | Description | Remediation |
|------|----------|------|-------------|-------------|
| {{#findings}} {{ rule_name }} | {{ severity }} | {{ type }} | {{ description }} | {{ remediation }} {{/findings}} |

The second row is the template row. It will be repeated for every finding, then the marker row is removed.

### Available Tables and Columns

#### `findings` — All findings

Columns: `rule_name`, `device_group`, `severity`, `type`, `description`, `remediation`, `risk_score`

#### `critical_findings` — Critical findings only

Same columns as `findings`, filtered to Critical severity.

#### `high_findings` — High findings only

Same columns as `findings`, filtered to High severity.

#### `shadowed_rules` — Shadowed rules

Columns: `rule_name`, `shadowed_by`, `device_group`, `severity`, `description`, `remediation`

#### `duplicate_addresses` — Duplicate address objects

Columns: `type`, `value`, `count`, `objects`, `remediation`

#### `compliance` — Compliance framework results

Columns: `framework`, `percentage`, `status`, `passed`, `failed`, `total`

#### `lateral_movement` — Lateral movement findings

Columns: `rule_name`, `severity`, `source_zones`, `dest_zones`, `risk_factors`

#### `weak_segments` — Weak segmentation pairs

Columns: `source_zone`, `dest_zone`, `openness`, `remediation`

#### `egress_findings` — Egress filtering findings

Columns: `rule_name`, `severity`, `risk_factors`, `remediation`

#### `cleartext_rules` — Cleartext protocol exposure

Columns: `rule_name`, `protocol`, `severity`, `secure_alternative`

#### `stale_rules` — Stale rules

Columns: `rule_name`, `severity`, `indicators`, `disabled`

#### `decryption_gaps` — Decryption policy gaps

Columns: `rule_name`, `severity`, `reason`, `remediation`

#### `geo_ip_findings` — Geo-IP exposure

Columns: `rule_name`, `severity`, `type`, `remediation`

#### `rule_expiry` — Expired and temporary rules

Columns: `rule_name`, `type`, `detail`

## Template Design Tips

### Formatting is preserved

All formatting in your template is preserved — fonts, colours, styles, spacing, logos, headers, footers, page layout. Placeholders are replaced in-place, inheriting the formatting of the surrounding text.

### Using styles for severity

If you want severity values to stand out, you can format the `{{ severity }}` cell in your template with conditional formatting or simply rely on the text values (Critical, High, Medium, Low).

### Headers and footers

Placeholders work in document headers and footers. Common uses:

- `{{ client_name }}` in the header
- `{{ confidentiality }}` in the footer
- `{{ report_date }}` in the footer

### Cover page

Design your cover page normally in the template. Use placeholders for dynamic content:

```
{{ report_title }}

Prepared for: {{ client_name }}
Prepared by: {{ auditor_name }}, {{ auditor_company }}
Date: {{ report_date }}

{{ confidentiality }}
```

### Conditional sections

The generator does not remove sections when data is empty — it simply leaves table data sections with zero rows. If you want to omit sections entirely when no data exists, check the corresponding count variable first and design your template accordingly, or remove those sections manually after generation.

### Multiple tables

You can include as many tables as needed. Each table operates independently. For example, you might have:

- An executive summary with `{{ risk_score }}` and severity counts
- A critical findings table using `{{#critical_findings}}`
- A compliance summary table using `{{#compliance}}`
- A full findings appendix using `{{#findings}}`

## Batch Generation

To generate reports for multiple JSON files:

```bash
for f in exports/*.json; do
    name=$(basename "$f" .json)
    python rampart_report.py "$f" template.docx "reports/${name}-report.docx" \
        --client "Client Name"
done
```

## Troubleshooting

### Placeholders not replaced

- Ensure the placeholder is typed as continuous text. If you type `{{`, then format part of the variable name differently (e.g. bold a few characters), Word splits it into multiple XML runs and the placeholder won't be detected. To fix: delete the placeholder and retype it in one go without changing formatting mid-word.
- Use `--list-variables` to confirm the variable name matches exactly.

### Table rows not appearing

- Ensure `{{#table_name}}` and `{{/table_name}}` are in the same row.
- Check that the table name matches one of the available tables listed above.
- Use `--list-variables` to confirm the table has data (check the row count).

### Empty output values

- Some variables depend on which analysers were enabled in Rampart. If an analyser was disabled, its variables will be 0 or empty.

### Large reports

- For configurations with hundreds of findings, the generated document may be large. Consider using filtered tables (`critical_findings`, `high_findings`) instead of the full `findings` table for the main body, and put the complete list in an appendix.
