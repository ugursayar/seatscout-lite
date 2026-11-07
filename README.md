# SeatScout — Microsoft 365 License Savings Audit

Find the Microsoft 365 licenses you're paying for and nobody's using — then recover the spend.
SeatScout runs **read-only inside your own tenant**. It has no server and collects nothing; the report is written to your machine.

---

## This repository: the free Lite build

This is **SeatScout Lite**, free and open source under the MIT License. A run prints a console summary and writes `inactive-users.csv`, `never-signed-in.csv` and `disabled-but-licensed.csv`; it shows no dollar totals and builds no HTML report. Anything below that mentions the HTML report, dollar totals, Solo or Pro belongs to the paid editions at [seatscout.dev](https://seatscout.dev); that code is not in this script. The `-Lite` switch changes nothing in this build: it always runs as Lite.

## What you get

| File | Purpose |
|------|---------|
| `SeatScout-Audit.ps1` | The audit engine (PowerShell + Microsoft Graph) |
| `config/PriceTable.csv` | Editable license price list — set your real prices here |
| `test/mock-tenant.json` | Demo data to try the tool with no tenant access |

Licensed MIT — Solo and Pro are under the EULA.

Outputs per run: a polished **HTML savings report** + **CSV exports** for each finding.

---

## Requirements

- **PowerShell 7+** (Windows, macOS, or Linux) — or Windows PowerShell 5.1
- **Microsoft Graph PowerShell SDK**:
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  ```
- An account that can consent to **read-only** Graph scopes (Global Reader is enough)
- For sign-in activity: **Microsoft Entra ID P1** (without it, SeatScout still finds disabled, never-signed-in and unassigned seats)

---

## Quick start (try it with demo data — no tenant needed)

```powershell
.\SeatScout-Audit.ps1 -MockDataPath .\test\mock-tenant.json
```
Open the generated `SeatScout-Report-<timestamp>\SeatScout-Report.html`.

## Run against your tenant

```powershell
# 1) (first time) install the SDK
Install-Module Microsoft.Graph -Scope CurrentUser

# 2) run the audit (you'll be prompted to sign in & consent to read-only scopes)
.\SeatScout-Audit.ps1 -InactiveDays 30
```

You'll be asked to consent to: `Organization.Read.All`, `User.Read.All`, `Directory.Read.All`, `AuditLog.Read.All`. All read-only.

---

## Options

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-InactiveDays` | `30` | A licensed, enabled user with no sign-in in this many days is flagged inactive |
| `-OutputPath` | auto | Folder for the report + CSVs |
| `-PriceTablePath` | `.\config\PriceTable.csv` | Your price list |
| `-CompanyName` | — | **(Pro)** White-label the report header |
| `-CurrencySymbol` / `-CurrencyLabel` | `$` / `USD` | Display currency |
| `-Lite` | off | Free mode: CSV + summary only (no $ totals, no HTML) |
| `-MockDataPath` | — | Use demo JSON instead of connecting to Graph |

**Example (consultant, white-labelled, 60-day threshold):**
```powershell
.\SeatScout-Audit.ps1 -InactiveDays 60 -CompanyName "Contoso IT Partners"
```

---

## Set your prices (important)

License prices vary by agreement, region and currency. Open `config/PriceTable.csv` and set `MonthlyPriceUSD` for the SKUs you own. The four most common Microsoft 365 SKUs ship with 2026 public list prices; others are blank and **excluded from totals until you set them**. The report tells you which purchased SKUs had no price.

> Tip: `Get-MgSubscribedSku | Select SkuPartNumber, SkuId` shows the exact part numbers in your tenant.

---

## What it checks

1. **Unassigned purchased seats** — bought but assigned to no one (cleanest win)
2. **Disabled but licensed** — blocked accounts still billed
3. **Never signed in** — licensed, never logged in (older than your threshold)
4. **Inactive > N days** — no recent sign-in
5. **Overlapping base plans** *(advisory)* — e.g. E3 + E5 on one user
6. **Downgrade candidates** *(advisory)* — E5 → E3 review list

Headline savings = items 1–4 (conservative). Advisory items are listed but not added to the total.

---

## Security & privacy

- **Read-only.** SeatScout never writes or changes anything in your tenant.
- **No exfiltration.** No network calls except to Microsoft Graph. The report is generated locally.
- Review the script before running — it's plain PowerShell, intentionally readable.

---

## Troubleshooting

- *"Sign-in data unavailable"* → tenant has no Entra ID P1; results still valid, real waste likely higher.
- *Module not found* → `Install-Module Microsoft.Graph -Scope CurrentUser`
- *Throttling on large tenants* → re-run; Graph paging is handled, transient 429s resolve on retry.
- *A SKU shows "not set"* → add its price in `config/PriceTable.csv`.

---

SeatScout · seatscout.dev · Independent tool, not affiliated with or endorsed by Microsoft.

## Author

Built and maintained by Uğur Sayar, a SharePoint and Microsoft 365 architect in Istanbul ([LinkedIn](https://www.linkedin.com/in/ugursayar)). Bug reports and pull requests are welcome through GitHub issues.
