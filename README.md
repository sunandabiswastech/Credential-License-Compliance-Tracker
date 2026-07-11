# Credential-License-Compliance-Tracker

A Python tool that audits employee license and certification records against
healthcare credentialing standards, flags discrepancies for follow-up, and
generates an audit-ready Excel report. Built to reflect how healthcare HR
and compliance teams actually monitor credentials, not just a generic
expiration-date checker.

## Why this exists

Healthcare organizations are required to verify that staff hold current,
valid licenses and certifications before providing care, and to keep that
verification current on an ongoing basis, not just at hire. This project
models that require as code: it checks each credential record for the
kinds of gaps that show up in real audits (missing identifiers, lapsed
verification, non-primary-source documentation, expired credentials) and
produces a report that a compliance team could actually use in survey prep.

## What it checks, and why

| Check | Regulatory basis |
|---|---|
| Missing credential/license number | Without an identifier, primary source verification (PSV) cannot be completed or documented. |
| Expired credential | The Joint Commission's HR chapter (HR.01.01.01 / HR.01.02.05) expects staff to hold current licensure/certification while providing services. |
| Approaching expiration | Modeled on risk-based monitoring: high-risk clinical credentials (RN, NP, MD/DO, DEA registration) get a longer warning window and more frequent review than lower-risk administrative credentials. |
| No verification on file | HR.01.02.05 requires documented verification: who verified it, when, and by what method. |
| Verification overdue for risk tier | Reflects guidance that high-risk clinical credentials (RN, NP/PA licenses, DEA registrations) warrant monthly-cadence monitoring, with quarterly and annual cadences for lower-risk credentials. |
| Non-primary-source verification method | The Joint Commission's PSV definition requires confirmation from the original issuing source (board portal, CVO report, DEA database), not self-reporting or a copy of the license alone. |
| Retain record - within NJ retention window | New Jersey requires personnel files to be retained for at least 2 years after separation of employment; this flag prevents premature record purges for recently separated staff. |

**Sources referenced:**
- The Joint Commission, Human Resources chapter, Standards HR.01.01.01 and HR.01.02.05 (primary source verification requirements)
- CMS Conditions of Participation, 42 CFR 482.12 (license/certification verification before patient care)
- New Jersey personnel record retention guidance (minimum 2 years post-separation; 6 years for wage/hour records under N.J.S.A. 34:11-4.1 et seq.)
- EEOC recordkeeping requirements (minimum 1 year retention from a qualifying personnel action)

## How it works

1. `data/sample_credentials.csv` holds a synthetic dataset of employee credential
   records (name, department, credential type, issuing authority, dates,
   verification method, risk tier). No real employee data is used anywhere
   in this project.
2. `src/credential_tracker.py` loads the records, runs each one through the
   checks above, and writes `reports/compliance_report.xlsx`, a two-tab
   workbook with a summary of flag counts and a color-coded detail sheet.
3. Run it with:
   ```bash
   pip install -r requirements.txt
   python src/credential_tracker.py
   ```

## Example output

The generated report includes:
- **Summary tab**: total records reviewed, discrepancy counts by severity, and a breakdown by issue type
- **Discrepancy Detail tab**: every flagged record with employee, credential type, issue, severity, and a plain-language explanation of why it was flagged, color-coded by severity for quick triage

## Scope and limitations

This is a portfolio and learning project, not a production compliance
system. The verification frequency and retention thresholds used here are
simplified defaults based on publicly available guidance, not a substitute
for an organization's actual policies, legal counsel, or the specific
requirements of a given state licensing board or accrediting body.
Real-world implementation would also need to handle secure PII storage,
integration with primary-source verification services, and role-based
access controls, none of which are in scope here.

## Tech stack

Python, pandas, openpyxl
