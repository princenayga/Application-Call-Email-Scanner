# Science Corps Fellowship Scanner — Claude Code Context

## What this project is

A fully-automated Python CLI tool that:
1. Fetches replies to a Science Corps fellowship email blast from Gmail (via OAuth2)
2. Auto-classifies obvious bounces (mailer-daemon / postmaster senders)
3. Sends all other emails to the Claude API (claude-opus-4-6) in batches for classification
4. Exports a colour-coded Excel report with a Summary sheet and a Replies sheet

There is no manual copy-paste step — the pipeline runs end-to-end with a single command.

## Project owner

Prince Nayga — Philippine Manager, Science Corps (`pnayga@science-corps.org`)

## Files

| File | Purpose |
|---|---|
| `fellowship_scanner.py` | Main script — all logic lives here |
| `requirements.txt` | Pinned Python dependencies |
| `credentials.json` | Google OAuth2 client secret — **never commit this** |
| `token.json` | Saved Gmail session token — **never commit this** |
| `output/emails_to_classify.txt` | (Legacy) formatted prompt file — no longer used |
| `output/fellowship_replies.xlsx` | The generated Excel report |
| `venv/` | Python virtual environment |

## Security — files that must NOT be committed

Add these to `.gitignore` before any `git push`:

```
credentials.json
token.json
.env
output/
venv/
__pycache__/
```

The `ANTHROPIC_API_KEY` must be supplied as an environment variable — it must never be hardcoded in `fellowship_scanner.py`. Setting it as a string in the config block is only acceptable on a machine that will never push to a shared repo.

## How to run (on any machine)

```bash
# 1. Clone / copy the repo
# 2. Create and activate a virtual environment
python -m venv venv
venv\Scripts\activate          # Windows
# source venv/bin/activate     # Mac / Linux

# 3. Install dependencies
pip install -r requirements.txt

# 4. Place credentials.json (Google OAuth2 Desktop App) in the project root

# 5. Set your Anthropic API key
export ANTHROPIC_API_KEY="sk-ant-..."   # Mac / Linux / Git Bash
# $env:ANTHROPIC_API_KEY="sk-ant-..."  # Windows PowerShell

# 6. Run
python fellowship_scanner.py
```

On first run a browser window opens for Gmail login. The session is saved to `token.json`; subsequent runs skip the browser.

## CONFIG block (top of fellowship_scanner.py)

All user-tunable settings are in the `# ─── CONFIG ───` block at the top of the script. Key variables:

| Variable | Default | Meaning |
|---|---|---|
| `SEARCH_KEYWORD` | `"Paid Teaching Fellowship Abroad for Recent STEM PhDs"` | Keyword used in both Gmail queries |
| `SEARCH_AFTER_DATE` | `"2025/07/01"` | Ignore emails older than this (YYYY/MM/DD) |
| `MY_EMAIL` | `pnayga@science-corps.org` | Outgoing emails from this address are skipped |
| `CLAUDE_MODEL` | `claude-opus-4-6` | Anthropic model used for classification |
| `CLASSIFICATION_BATCH_SIZE` | `20` | Emails per Claude API call; lower if you hit token limits |
| `API_CALL_DELAY_SECONDS` | `1.0` | Pause between batches (rate-limit courtesy) |

## Gmail search queries

The script runs **two** queries and deduplicates by message ID:

1. **Main**: `(category:primary OR category:updates) "KEYWORD" after:DATE`
   — Catches human replies, auto-replies, and list moderation notices in Primary + Updates only.

2. **Bounce**: `(from:mailer-daemon OR from:postmaster) "KEYWORD" after:DATE`
   — Catches NDR / undeliverable messages that may land in any category.

## Auto-classification rules

Before calling the Claude API, the script checks each sender:
- If the local part of the sender email is `mailer-daemon` or `postmaster` → instantly tagged **Bounce / Delivery Failure**, no API call needed.
- All other emails → sent to Claude in batches of `CLASSIFICATION_BATCH_SIZE`.

## Classification categories

| Category | Excel colour |
|---|---|
| Positive / Interested | Green |
| Bounce / Delivery Failure | Light red |
| No Longer Works There | Orange-yellow |
| Declined / Not Interested | Dark red (white font) |
| Has a Question | Yellow |
| Application Submitted | Blue |

## Excel output structure

- **Sheet 1 — Summary**: total rows, count per classification category, count per Gmail category, generation date.
- **Sheet 2 — Replies**: one row per email with columns: `#`, `Sender Name`, `Email Address`, `Institution`, `Date Received`, `Gmail Category`, `Category`, `Summary`, `Action Needed`.

Re-running the script appends new results without duplicating rows already present (matched by sender email + date received).

## Dependencies

All pinned in `requirements.txt`. Key packages:
- `anthropic` — Claude API client
- `google-api-python-client`, `google-auth-oauthlib`, `google-auth-httplib2` — Gmail API
- `openpyxl` — Excel export

## Known behaviour notes

- Mailing list rejection emails (from `*-owner@lists.*`) are NOT auto-tagged as bounces — Claude classifies them (usually as Bounce / Delivery Failure or Declined).
- Out-of-office auto-replies are classified by Claude (usually as No Longer Works There or Declined / Not Interested depending on content).
- The `institution` column is extracted from the sender's email domain; free domains (gmail.com, yahoo.com, etc.) trigger a regex scan of the email body instead.
- `emails_to_classify.txt` in the `output/` folder is no longer generated — that was part of the old manual workflow.
