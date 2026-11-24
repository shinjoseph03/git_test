# Desk reservation bot

GitHub Actions + Python bot that books (and clears) a preferred desk in the shared SharePoint workbook **Desk Reservation Master.xlsx** on the Las Vegas Office Homepage team site.

Workbook sheet **Master Bookings** columns: `Date | Desk | Area | AM Employee | PM Employee | AM Status | PM Status`.

On office weekdays the bot finds today’s row for the configured desk and, if AM and PM status are both `Available`, writes the configured display name into AM/PM Employee. Owner-blocked / locked / already-booked rows are skipped and logged (not treated as hard failures).

> **Note:** This code currently lives on a branch of [`shinjoseph03/git_test`](https://github.com/shinjoseph03/git_test) because creating a new private repo `desk-reservation-bot` was blocked for the cloud agent. Prefer moving it to a dedicated private repo when you can.

## SharePoint file

| Field | Value |
| --- | --- |
| File | `Desk Reservation Master.xlsx` |
| Site | `https://hdrinc.sharepoint.com/teams/LasVegasOfficeHomepage` |
| Drive item GUID (`sourcedoc`) | `4CC41F6E-391F-4BBC-B75B-3D0BABCA2BE5` |
| Browser URL | [Open workbook](https://hdrinc.sharepoint.com/:x:/r/teams/LasVegasOfficeHomepage/_layouts/15/doc2.aspx?sourcedoc=%7B4CC41F6E-391F-4BBC-B75B-3D0BABCA2BE5%7D&file=Desk%20Reservation%20Master.xlsx&action=default&mobileredirect=true) |

## Setup (Joseph)

### 1. Azure AD app registration

1. Azure Portal → **App registrations** → New registration (name e.g. `desk-reservation-bot`).
2. **Certificates & secrets** → New client secret → copy the value once.
3. **API permissions** → Microsoft Graph **Application** permissions:
   - `Sites.Selected` (preferred) **or** `Files.ReadWrite.All` / `Sites.ReadWrite.All` for broader access.
4. Grant admin consent.
5. If using `Sites.Selected`, grant this app write access to the LasVegasOfficeHomepage site (Graph `sites/{site-id}/permissions` or SharePoint admin “Manage app access”).

The app identity must be able to **edit** the workbook. If cells are locked by the sheet owner, the bot will skip and log.

### 2. GitHub repository secrets

Settings → Secrets and variables → Actions → New repository secret:

| Secret | Required | Notes |
| --- | --- | --- |
| `AZURE_TENANT_ID` | yes | Directory (tenant) ID |
| `AZURE_CLIENT_ID` | yes | Application (client) ID |
| `AZURE_CLIENT_SECRET` | yes | Client secret value |
| `SHAREPOINT_DRIVE_ITEM_ID` | recommended | `4CC41F6E-391F-4BBC-B75B-3D0BABCA2BE5` (overrides config) |
| `SHAREPOINT_SITE_HOSTNAME` | optional | Default `hdrinc.sharepoint.com` |
| `SHAREPOINT_SITE_PATH` | optional | Default `/teams/LasVegasOfficeHomepage` |

Without Azure secrets, the **test** job still passes; the live SharePoint job skips.

### 3. Config

```bash
cp config.example.yaml config.yaml
# edit display_name, preferred_desk, office_weekdays, timezone
```

`config.yaml` is gitignored. For Actions, either:

- commit a non-secret `config.yaml` with your name/desk/weekdays, or
- keep using `config.example.yaml` defaults after editing the example (or add a step that writes config from a secret).

Set at least:

- `display_name` — exact string written into AM/PM Employee  
- `preferred_desk` — e.g. `Desk 01`  
- `office_weekdays` — e.g. `Mon`, `Tue`, `Thu`, `Fri`  
- `timezone` — e.g. `America/Los_Angeles`  
- `sharepoint.worksheet_name` — tab name (`Master Bookings` by default)

### 4. Enable the workflow

- **Schedule:** weekdays `30 14 * * 1-5` UTC (~6:30am PT).  
- **Manual:** Actions → **Desk reservation bot** → Run workflow  
  - `mode=book` — book for today (or override `date`)  
  - `mode=clear` — clear your name on `clear_dates` (comma-separated)  
  - `dry_run=true` — compute only, no upload  

Clear mode only removes the **configured** display name; other people’s bookings are never cleared.

## Local / dry-run

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp config.example.yaml config.yaml

# Unit tests (no secrets)
pytest -q

# Offline against a downloaded xlsx
python -m desk_bot --mode book --fixture ./Desk\ Reservation\ Master.xlsx --date 2026-11-24 --dry-run
python -m desk_bot --mode clear --fixture ./master.xlsx --clear-date 2026-11-24 --dry-run
```

## Behavior summary

| Situation | Result |
| --- | --- |
| Today not an office weekday | skip + log |
| No row for preferred desk/date | skip + log |
| AM or PM status ≠ `Available` | skip + log |
| Slot held by someone else | skip + log |
| Already booked as you | no-op success |
| Both Available | write name, upload |
| Upload / cell locked | skip + warn (exit 0) |

## Layout

```
desk_bot/          # Python package
tests/             # Fixture-based unit tests
config.example.yaml
.github/workflows/desk-bot.yml
```
