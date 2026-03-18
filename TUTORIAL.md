# Rampart Report Generator — Tutorial

This tutorial walks you through every step of using the Rampart Report Generator, from installation to producing your first branded audit report. By the end, you will know how to export data from Rampart, build a custom Word template, generate a report, and tailor the output to your needs.

---

## Table of Contents

1. [What This Tool Does](#1-what-this-tool-does)
2. [Prerequisites](#2-prerequisites)
3. [Installation](#3-installation)
4. [Step 1: Export Your Data from Rampart](#4-step-1-export-your-data-from-rampart)
5. [Step 2: Understand the JSON Export](#5-step-2-understand-the-json-export)
6. [Step 3: Create a Word Template](#6-step-3-create-a-word-template)
7. [Step 4: Add Placeholders to Your Template](#7-step-4-add-placeholders-to-your-template)
8. [Step 5: Add Repeating Tables](#8-step-5-add-repeating-tables)
9. [Step 6: Generate the Report](#9-step-6-generate-the-report)
10. [Step 7: Review and Finalise](#10-step-7-review-and-finalise)
11. [Using the Starter Template](#11-using-the-starter-template)
12. [Discovering Available Variables](#12-discovering-available-variables)
13. [Working with Headers and Footers](#13-working-with-headers-and-footers)
14. [Batch Report Generation](#14-batch-report-generation)
15. [Troubleshooting Common Issues](#15-troubleshooting-common-issues)
16. [Tips and Best Practices](#16-tips-and-best-practices)

---

## 1. What This Tool Does

The Rampart Report Generator takes two inputs and produces one output:

```
Rampart JSON Export  +  Word Template (.docx)  →  Branded Audit Report (.docx)
```

**The JSON export** contains all the data from a Rampart firewall security analysis — findings, risk scores, compliance results, shadowed rules, and more.

**The Word template** is a `.docx` file you design with your company's branding (logos, headers, footers, fonts, colours). You place special placeholders like `{{ client_name }}` and `{{ risk_score }}` where you want data to appear.

**The generated report** is a copy of your template with every placeholder replaced by the actual values from the analysis. All formatting, styles, and branding are preserved exactly as you designed them.

This means you design the report once, and then generate it repeatedly for different audits — saving hours of manual copy-paste work while maintaining a consistent, professional output.

---

## 2. Prerequisites

Before you begin, make sure you have:

- **Python 3.8 or later** — Check your version by running `python --version` or `python3 --version` in your terminal.
- **pip** — Python's package manager (bundled with most Python installations).
- **A Rampart JSON export** — The data file produced by Rampart's export feature.
- **Microsoft Word** (or a compatible editor like LibreOffice Writer) — For designing your template and reviewing the output.

---

## 3. Installation

Open a terminal and navigate to the project directory, then install the required Python packages:

```bash
cd rampart-report-generator
pip install -r requirements.txt
```

This installs two libraries:

- **python-docx** — Reads and writes Word `.docx` files programmatically.
- **Jinja2** — Provides the placeholder syntax used in templates.

To verify the installation worked, run:

```bash
python rampart_report.py --help
```

You should see the usage information and a list of all command-line options.

---

## 4. Step 1: Export Your Data from Rampart

To generate a report, you first need data from Rampart:

1. Open **Rampart** and load your firewall configuration.
2. Run the **analysis** — this evaluates your firewall rules against Rampart's security checks.
3. Once the analysis is complete, click **Export Results** in the sidebar (or press **Ctrl+E**).
4. Choose **JSON** as the export format.
5. Save the file somewhere convenient, for example `acme-audit.json`.

This JSON file contains everything the report generator needs: findings, risk scores, compliance data, shadowed rules, duplicate objects, segmentation analysis, and more. The specific data available depends on which Rampart analysers you enabled.

---

## 5. Step 2: Understand the JSON Export

Before building a template, it helps to understand what data is available. The report generator extracts two types of data from the JSON:

### Scalar Variables

These are single values — a number, a string, or a date. They are used for individual metrics like:

- `{{ total_rules }}` — the total number of firewall rules (e.g., "247")
- `{{ risk_grade }}` — the overall risk grade (e.g., "C")
- `{{ critical_count }}` — the number of critical findings (e.g., "5")

### Table Data

These are lists of items with multiple columns — used for repeating rows in tables. For example, the `findings` table contains one entry per finding, each with columns like `rule_name`, `severity`, `description`, and `remediation`.

To see exactly which variables and tables are available for a specific JSON export, use the `--list-variables` command (covered in detail in [Section 12](#12-discovering-available-variables)).

---

## 6. Step 3: Create a Word Template

The template is a standard `.docx` file. Open Microsoft Word (or your preferred editor) and design the document the way you want the final report to look:

1. **Set up page layout** — Choose your paper size (A4, Letter), margins, and orientation.
2. **Add your branding** — Insert your company logo, set header/footer content, choose your fonts and colour scheme.
3. **Create the document structure** — Add headings for each section (Executive Summary, Risk Assessment, Findings, etc.).
4. **Write the static content** — Add any boilerplate text that stays the same in every report (methodology descriptions, disclaimers, terms of service).

At this point, you have a nicely formatted but empty report shell. The next step is adding placeholders.

### Important: Use Heading Styles

Use Word's built-in Heading styles (Heading 1, Heading 2, Heading 3) for your section titles rather than just making text bold and large. This ensures:

- The Table of Contents generates correctly.
- The document structure is accessible and navigable.
- Consistent formatting throughout.

---

## 7. Step 4: Add Placeholders to Your Template

Placeholders tell the generator where to insert data. They use double curly braces:

```
{{ variable_name }}
```

Simply type a placeholder wherever you want a value to appear. For example, on your cover page you might write:

```
{{ report_title }}

Prepared for: {{ client_name }}
Prepared by: {{ auditor_name }}, {{ auditor_company }}
Date: {{ report_date }}

Classification: {{ confidentiality }}
```

In your executive summary section, you might write a paragraph like:

```
This report presents the results of a firewall security audit conducted for
{{ client_name }} on {{ report_date }}. The audit analysed {{ total_rules }}
firewall rules across {{ device_group_count }} device group(s). Of these,
{{ rules_with_issues }} rule(s) had at least one finding, resulting in a
compliance rate of {{ compliance_rate }}%.
```

When the report is generated, every `{{ ... }}` placeholder is replaced by its corresponding value from the JSON data or command-line options.

### Key Points About Placeholders

- **Formatting is inherited** — The replacement text takes on the font, size, colour, and style of the placeholder text. If you make `{{ risk_grade }}` bold and red, the actual grade value will be bold and red.
- **Placeholders work everywhere** — Body text, headers, footers, table cells, and even text boxes.
- **Spacing is flexible** — `{{risk_grade}}`, `{{ risk_grade }}`, and `{{  risk_grade  }}` all work.
- **Type them in one go** — Do not change formatting partway through a placeholder (e.g., making `risk` bold but `grade` italic). Word stores text in "runs" split by formatting changes, and this will break the placeholder. If this happens, delete the placeholder and retype it without changing formatting mid-way.

### Common Placeholder Groups

**Report metadata** (set via command-line options):
- `{{ report_title }}`, `{{ report_date }}`, `{{ client_name }}`, `{{ client_contact }}`
- `{{ auditor_name }}`, `{{ auditor_company }}`, `{{ confidentiality }}`

**Summary statistics:**
- `{{ total_rules }}`, `{{ rules_with_issues }}`, `{{ compliance_rate }}`
- `{{ critical_count }}`, `{{ high_count }}`, `{{ medium_count }}`, `{{ low_count }}`, `{{ total_findings }}`

**Risk assessment:**
- `{{ risk_score }}`, `{{ risk_grade }}`, `{{ best_practices_score }}`, `{{ segmentation_score }}`

See the [User Guide](USER_GUIDE.md) for a complete list of all available placeholders.

---

## 8. Step 5: Add Repeating Tables

For data that has multiple rows — such as a list of findings or compliance frameworks — you use table markers. This is how you turn a static table into a dynamic one that grows based on the data.

### How It Works

1. **Create a normal Word table** with your desired column headers.
2. **Add a marker row** below the headers that tells the generator which dataset to use and which columns to fill.

The marker row uses a special syntax:

- The **first cell** of the row contains `{{#table_name}}` — this is the opening marker.
- The **remaining cells** contain `{{ column_name }}` placeholders for each column.
- The **last cell** also contains `{{/table_name}}` — the closing marker.

### Example: Creating a Findings Table

In Word, create a table that looks like this:

| Rule | Device Group | Severity | Type | Description | Remediation |
|------|-------------|----------|------|-------------|-------------|
| `{{#findings}}` `{{ rule_name }}` | `{{ device_group }}` | `{{ severity }}` | `{{ type }}` | `{{ description }}` | `{{ remediation }}` `{{/findings}}` |

When the report is generated:

1. The generator finds the `{{#findings}}` marker and identifies this as a repeating row.
2. It looks up the `findings` dataset from the JSON data.
3. For each finding, it creates a copy of the marker row with the placeholders replaced by that finding's values.
4. The original marker row is removed.

If there are 15 findings, the table will have 15 data rows (plus the header row). If there are no findings, the table will have just the header row.

### Available Table Datasets

Here are the most commonly used tables:

| Table Name | Description | Columns |
|-----------|-------------|---------|
| `findings` | All findings | `rule_name`, `device_group`, `severity`, `type`, `description`, `remediation`, `risk_score` |
| `critical_findings` | Critical findings only | Same as `findings` |
| `high_findings` | High findings only | Same as `findings` |
| `shadowed_rules` | Shadowed rules | `rule_name`, `shadowed_by`, `device_group`, `severity`, `description`, `remediation` |
| `compliance` | Compliance results | `framework`, `percentage`, `status`, `passed`, `failed`, `total` |
| `duplicate_addresses` | Duplicate objects | `type`, `value`, `count`, `objects`, `remediation` |
| `lateral_movement` | Lateral movement | `rule_name`, `severity`, `source_zones`, `dest_zones`, `risk_factors` |
| `weak_segments` | Weak segments | `source_zone`, `dest_zone`, `openness`, `remediation` |
| `cleartext_rules` | Cleartext protocols | `rule_name`, `protocol`, `severity`, `secure_alternative` |
| `stale_rules` | Stale rules | `rule_name`, `severity`, `indicators`, `disabled` |
| `egress_findings` | Egress risks | `rule_name`, `severity`, `risk_factors`, `remediation` |
| `decryption_gaps` | Decryption gaps | `rule_name`, `severity`, `reason`, `remediation` |
| `geo_ip_findings` | Geo-IP exposure | `rule_name`, `severity`, `type`, `remediation` |
| `rule_expiry` | Rule expiry | `rule_name`, `type`, `detail` |

### Tips for Tables

- You can include as many different tables as you want in a single template.
- You can use the same dataset in multiple tables (e.g., `findings` in the main body and again in an appendix).
- You do not need to include all columns — pick just the ones you want to show.
- The column order in your table can differ from the order listed above.
- You can include static text alongside placeholders in a cell (e.g., "Rule: {{ rule_name }}").

---

## 9. Step 6: Generate the Report

With your JSON export and template ready, run the generator:

```bash
python rampart_report.py acme-audit.json my-template.docx acme-report.docx \
    --client "Acme Corporation" \
    --client-contact "John Davies" \
    --auditor "Jane Smith" \
    --company "SecureAudit Pty Ltd" \
    --date "2026-03-19" \
    --confidentiality "CONFIDENTIAL"
```

Let's break down each argument:

| Argument | Purpose |
|----------|---------|
| `acme-audit.json` | The Rampart JSON export file (input data) |
| `my-template.docx` | Your Word template with placeholders (input template) |
| `acme-report.docx` | The output file to create (generated report) |
| `--client` | Sets the `{{ client_name }}` placeholder |
| `--client-contact` | Sets the `{{ client_contact }}` placeholder |
| `--auditor` | Sets the `{{ auditor_name }}` placeholder |
| `--company` | Sets the `{{ auditor_company }}` placeholder |
| `--date` | Sets the `{{ report_date }}` placeholder (defaults to today if omitted) |
| `--confidentiality` | Sets the `{{ confidentiality }}` placeholder (defaults to "CONFIDENTIAL") |

### Default Values

If you omit optional arguments:

- `--date` defaults to today's date.
- `--title` defaults to "Firewall Security Audit Report".
- `--confidentiality` defaults to "CONFIDENTIAL".
- `--client`, `--client-contact`, `--auditor`, and `--company` default to empty strings — any placeholders referencing them will be replaced with blank text.

---

## 10. Step 7: Review and Finalise

After generation, open the output `.docx` in Word and review it:

1. **Check the cover page** — Verify client name, date, and auditor details populated correctly.
2. **Update the Table of Contents** — If your template includes a TOC field, right-click it and select **Update Field** > **Update entire table**. This refreshes page numbers and headings.
3. **Review the tables** — Confirm that findings, compliance data, and other repeating sections have the expected number of rows.
4. **Spot-check values** — Compare a few key metrics (risk score, finding counts) against the Rampart UI to ensure accuracy.
5. **Look for unreplaced placeholders** — If you see any `{{ ... }}` text still in the document, the variable name may be misspelled. Use `--list-variables` to check the correct name.
6. **Final polish** — Add any manual commentary, executive observations, or recommendations specific to this engagement that go beyond what the automated data provides.

---

## 11. Using the Starter Template

This project includes a script to generate a ready-made starter template with all sections and placeholders pre-configured:

```bash
python create_template.py template.docx
```

This creates `template.docx` with:

- A cover page with report metadata placeholders
- A Table of Contents field (update it in Word after generation)
- Executive Summary with key metrics and severity breakdown
- Risk Assessment section
- Compliance Summary with the `{{#compliance}}` table
- Detailed Findings sections (critical, high, and all findings tables)
- Shadowed Rules table
- Duplicate Objects table
- Network Segmentation with weak segments and lateral movement tables
- Additional Analysis sections (cleartext, stale rules, egress, decryption gaps, geo-IP, rule expiry)
- Recommendations section with dynamic bullet points
- Appendices for methodology and best practices
- Headers and footers with company name, confidentiality marking, and report metadata

You can use this template as-is for a quick start, or open it in Word and customise the branding, colours, fonts, logos, and layout to match your organisation's standards. The placeholders and table markers will continue to work regardless of the visual styling you apply.

---

## 12. Discovering Available Variables

Every Rampart JSON export may contain different data depending on which analysers were enabled. To see exactly what is available for your specific export:

```bash
python rampart_report.py --list-variables acme-audit.json
```

This prints two sections:

### Scalar Variables

A table showing every variable name and its current value:

```
Template Variables:
  report_date            = 2026-03-19
  total_rules            = 247
  rules_with_issues      = 38
  compliance_rate        = 84.2
  risk_score             = 72
  risk_grade             = C
  critical_count         = 5
  ...
```

Use these names in `{{ ... }}` placeholders in your template.

### Table Datasets

A list of every available table, its column names, and row count:

```
Table Data:
  findings (38 rows): rule_name, device_group, severity, type, description, remediation, risk_score
  critical_findings (5 rows): rule_name, device_group, severity, type, ...
  compliance (3 rows): framework, percentage, status, passed, failed, total
  ...
```

Use these names in `{{#table_name}}` / `{{/table_name}}` markers. The column names go in `{{ column_name }}` placeholders within the marker row.

This command is your definitive reference — if a variable shows a value, you can use it. If a table shows 0 rows, that section will be empty in the report.

---

## 13. Working with Headers and Footers

Placeholders in headers and footers work exactly like those in the document body. The generator processes all sections of the document, including:

- Default headers and footers
- First-page headers and footers
- Section-specific headers and footers

### Common Header/Footer Patterns

**Header (right-aligned):**
```
{{ auditor_company }}  |  {{ confidentiality }}
```

**Footer (centred):**
```
{{ report_title }}  |  {{ client_name }}  |  {{ report_date }}
```

**First-page footer:**
```
This document is classified {{ confidentiality }}.
```

To set these up, double-click the header or footer area in Word, type the placeholder, apply your desired formatting, then click back into the document body.

---

## 14. Batch Report Generation

If you have multiple JSON exports to process (e.g., from different firewalls or different clients), you can automate report generation with a shell loop:

### Process All JSON Files in a Directory

```bash
for f in exports/*.json; do
    name=$(basename "$f" .json)
    python rampart_report.py "$f" template.docx "reports/${name}-report.docx" \
        --client "Acme Corporation" \
        --auditor "Jane Smith" \
        --company "SecureAudit Pty Ltd"
done
```

This generates one report per JSON file, named after the input file.

### Different Clients from a CSV

For more complex scenarios, you could read client details from a file:

```bash
while IFS=',' read -r json_file client_name contact; do
    output="reports/${client_name// /-}-report.docx"
    python rampart_report.py "$json_file" template.docx "$output" \
        --client "$client_name" \
        --client-contact "$contact" \
        --auditor "Jane Smith" \
        --company "SecureAudit Pty Ltd"
done < clients.csv
```

---

## 15. Troubleshooting Common Issues

### Placeholder Not Replaced (Still Shows `{{ ... }}` in Output)

**Cause:** Word may have split the placeholder across multiple text "runs" due to formatting changes, spell-check marks, or editing history.

**Fix:** In your template, delete the entire placeholder and retype it in a single action without pausing or changing formatting. Do not type `{{`, then select part of the text to bold it, then continue typing — type the whole thing at once.

**Verification:** Use `--list-variables` to confirm you are using the exact variable name.

### Table Rows Not Appearing

**Cause:** The table marker names do not match, or the markers are not in the same row.

**Fix:**
- Ensure `{{#table_name}}` appears in the **first cell** of the marker row.
- Ensure `{{/table_name}}` appears in the **last cell** of the same row.
- Check that the table name matches one of the available datasets (use `--list-variables` to verify).
- Confirm the dataset has data — if it shows 0 rows, no data rows will be generated.

### Values Are Empty or Zero

**Cause:** The corresponding Rampart analyser was not enabled when the analysis was run.

**Fix:** Re-run the analysis in Rampart with all required analysers enabled, then re-export the JSON.

### Report Looks Different from Template

**Cause:** This is usually not a generator issue — the generator preserves all formatting. If the output looks different:

- Check that your template file is not corrupted.
- Ensure you are opening the output in the same application used to create the template (Word vs. LibreOffice may render some features differently).

### Error: "File Not Found"

**Fix:** Check that all three file paths (JSON, template, output) are correct. Use absolute paths if relative paths are causing confusion:

```bash
python rampart_report.py /path/to/audit.json /path/to/template.docx /path/to/output.docx
```

### Error During Generation

If the script fails with a Python error:

1. Verify the JSON file is valid — open it in a text editor and check for obvious issues.
2. Verify the template is a valid `.docx` file — try opening it in Word first.
3. Check the Python version (`python --version` should show 3.8+).
4. Reinstall dependencies: `pip install -r requirements.txt --force-reinstall`.

---

## 16. Tips and Best Practices

### Template Design

- **Start from the starter template** — Run `python create_template.py` and customise from there rather than building from scratch. This ensures all sections and placeholders are correctly set up.
- **Use styles consistently** — Apply Word's built-in Heading 1, Heading 2, and Heading 3 styles so the Table of Contents works.
- **Design for the worst case** — Consider how tables will look with many rows. Wide tables may benefit from landscape sections or smaller font sizes.
- **Keep a master template** — Maintain one "golden" template file and make copies for specific clients or engagements.

### Report Structure

- **Lead with critical findings** — Use the `critical_findings` and `high_findings` tables in the executive summary for immediate impact.
- **Use the full `findings` table sparingly** — For large environments, the complete findings list can run to many pages. Consider placing it in an appendix.
- **Add manual commentary** — The generator fills in the data, but a good audit report includes expert interpretation. Leave space for your analysis and recommendations.

### Workflow

- **Preview with `--list-variables` first** — Before generating, check what data is available. This prevents surprises in the output.
- **Version your templates** — Keep templates in version control alongside the generator so you can track changes over time.
- **Automate where possible** — Use batch generation for recurring audits or multi-firewall environments.
- **Review every report** — Automated generation saves time, but always review the output before delivering to a client. Check for empty sections, unusual values, or formatting issues.
