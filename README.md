# memic-payroll

Generates MEMIC Comp-As-You-Go workers comp submission files from Square payroll CSV exports.

## Setup

```bash
pip3 install openpyxl
```

## Running

```bash
cd ~/Dev/memic-payroll
python3 memic_payroll.py
```

The script will prompt for the Square payroll CSV path, then output a filled `.xlsx` and the Check Date needed for manual upload to [memic.payrollpl.us](https://memic.payrollpl.us/).

Each Square payroll run produces one file. If a pay period has two runs (e.g. a separate PTO-only run), run the script twice and upload both.

## Inputs

**Square payroll CSV** — exported from Square Payroll dashboard. Contains one row per employee with hours, earnings, and gross pay.

**MEMIC template** (`Express_Payroll_Template_*.xlsx`) — downloaded from [memic.payrollpl.us](https://memic.payrollpl.us/). Contains the full employee roster with MEMIC employee IDs and class codes. Only needed when the script encounters an employee it can't match against its stored roster (see below).

## Employee mapping (`employees.json`)

The script maintains `employees.json` to map Square names to MEMIC names and IDs. It is built automatically the first time you provide a template and persists across runs.

**When the script prompts for a template:**
- An employee appeared in the payroll CSV who isn't in the stored roster
- If they're a new hire: **add them to MEMIC first**, then download a fresh template before providing it here — otherwise they won't be in the file
- If they're an existing employee with an unexpected name format: provide the current template and confirm the fuzzy match

The script will confirm any new name matches (e.g. "Valerie Levanos" → "Levanos, Val") before saving them. Once confirmed, that mapping is remembered for all future runs.

## Output

Files are saved to `output/MEMIC_Payroll_YYYY-MM-DD.xlsx` named by pay date. All MEMIC roster employees are included; those not in the current run have zeros.

Upload each file to [memic.payrollpl.us](https://memic.payrollpl.us/) and enter the Check Date printed by the script.

## What's not automated

The actual upload to the MEMIC portal is manual — log in, upload the `.xlsx`, confirm the payment.
