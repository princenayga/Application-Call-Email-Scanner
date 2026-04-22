# Science Corps Fellowship Scanner — Claude Code Context

## What this project is

A fully-automated Python CLI tool that:
1. Fetches replies to a Science Corps fellowship email blast from Gmail (via OAuth2)
2. Cross-references replies against a listserv CSV to track non-replies
3. Auto-classifies obvious bounces (mailer-daemon / postmaster / -owner senders)
4. Sends all other emails to the Claude API (claude-opus-4-6) in batches for classification
5. Extracts failed recipient addresses from NDR bodies to reclassify No Reply → Bounce
6. Exports a colour-coded Excel report with Action Plan, Summary, and Replies sheets
7. Generates a cleaned listserv (dead bounces removed) and a positive-reply outreach report

There is no manual copy-paste step — the pipeline runs end-to-end with a single command.

## Project owner

Prince Nayga — Philippine Manager, Science Corps (`pnayga@science-corps.org`)

## Files

| File | Purpose |
|---|---|
| `fellowship_scanner.py` | Main script — all logic lives here |
| `clean_listserv.py` | Generates cleaned listserv Excel (bounces stripped, review flagged) |
| `positive_replies_report.py` | Generates HTML outreach report with draft emails for positive replies |
| `generate_guide.py` | Generates printable HTML implementation guide |
| `requirements.txt` | Pinned Python dependencies |
| `credentials.json` | Google OAuth2 client secret — **never commit this** |
| `token.json` | Saved Gmail session token — **never commit this** |
| `*.csv` | Listserv CSV files — gitignored (contain contact emails) |
| `output/fellowship_replies.xlsx` | Main colour-coded Excel report (Action Plan + Summary + Replies) |
| `output/cleaned_listserv.xlsx` | Filtered listserv for the next blast cycle |
| `output/positive_replies_outreach.html` | Outreach drafts for positive replies — open in Chrome, print to PDF |
| `output/implementation_guide.html` | Step-by-step workflow guide — open in Chrome, print to PDF |
| `read_this_files/` | Snapshot of latest output files committed to repo |
| `venv/` | Python virtual environment |

## Security — files that must NOT be committed

Already in `.gitignore`:

```
credentials.json
token.json
*.csv
output/
venv/
__pycache__/
```

The `ANTHROPIC_API_KEY` must be supplied as an environment variable — never hardcoded in `fellowship_scanner.py`.

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

# 6. Run the main scanner
python fellowship_scanner.py

# 7. (Optional) Generate the positive replies outreach report
python positive_replies_report.py

# 8. (Optional) Generate the cleaned listserv for the next cycle
python clean_listserv.py

# 9. (Optional) Regenerate the implementation guide
python generate_guide.py
```

On first run a browser window opens for Gmail login. The session is saved to `token.json`; subsequent runs skip the browser.

## CONFIG block (top of fellowship_scanner.py)

All user-tunable settings are in the `# ─── CONFIG ───` block at the top of the script. Key variables:

| Variable | Default | Meaning |
|---|---|---|
| `SEARCH_KEYWORD` | `"Paid Teaching Fellowship Abroad for Recent STEM PhDs"` | Keyword used in Gmail queries |
| `SEARCH_AFTER_DATE` | `"2025/12/13"` | Ignore emails older than this (YYYY/MM/DD) — update each cycle |
| `MY_EMAIL` | `pnayga@science-corps.org` | Outgoing emails from this address are skipped |
| `CC_EMAILS` | `{ccorry@..., cjellareroma@...}` | Addresses CC'd on the blast — excluded from NDR extraction |
| `LISTSERV_CSV` | `Filtered Listserv for Feb 2026 Applications - Sheet1.csv` | Listserv file to cross-reference — update each cycle |
| `CLAUDE_MODEL` | `claude-opus-4-6` | Anthropic model used for classification |
| `CLASSIFICATION_BATCH_SIZE` | `20` | Emails per Claude API call; lower if you hit token limits |
| `API_CALL_DELAY_SECONDS` | `1.0` | Pause between batches (rate-limit courtesy) |
| `FROM_SEARCH_BATCH_SIZE` | `20` | Listserv addresses per from: query batch |

## Gmail search queries

The script runs **four** searches and deduplicates by message ID:

1. **Main** (inbox + spam): `"KEYWORD" after:DATE -from:MY_EMAIL`
2. **Bounce sender** (inbox + spam): `(from:mailer-daemon OR from:postmaster) after:DATE`
3. **Bounce subject** (inbox + spam): subject-line NDR patterns after:DATE
4. **Listserv from-search** (inbox + spam): batched `(from:addr1 OR from:addr2...) after:DATE` — 41 batches of 20

All four searches include `includeSpamTrash=True` so replies mis-classified by Gmail as spam are not missed.

## Auto-classification rules

Before calling the Claude API, the script checks each sender:
- `mailer-daemon` or `postmaster` local part → **Bounce / Delivery Failure** (no API call)
- local part ending in `-owner` → **Bounce / Delivery Failure** (no API call)
- All other emails → sent to Claude in batches

NDR bodies are also scanned with a regex to extract failed recipient addresses. Any listserv address found in an NDR body is reclassified from No Reply → Bounce (excluding `MY_EMAIL` and `CC_EMAILS`).

## Classification categories

| Category | Excel colour | Meaning |
|---|---|---|
| Positive / Interested | Green | Human reply showing interest |
| Has a Question | Yellow | Human reply with a question |
| Application Submitted | Blue | Applicant confirmed submission |
| No Longer Works There | Orange-yellow | Person has left that institution |
| Declined / Not Interested | Dark red (white font) | Explicit decline |
| Auto-Reply | Light purple | Delivered OK; automated OOO response |
| Bounce / Delivery Failure | Light red | Address failed or NDR received |
| No Reply | Light grey | Listserv address with no reply found |
| Unverified | Very light grey | Not found in sent records (rare) |

## Excel output structure

- **Sheet 1 — Action Plan**: rebuilt fresh every run; 4 prioritised steps with full detail rows for Steps 1–2
- **Sheet 2 — Summary**: total counts by category and Gmail category, generation date
- **Sheet 3 — Replies**: one row per email/address with columns: `#`, `Sender Name`, `Email Address`, `Institution`, `Date Received`, `Gmail Category`, `Category`, `Summary`, `Action Needed`

Re-running appends new results without duplicating rows already present (matched by sender email + date received).

## Action Plan steps

| Step | Categories | What to do |
|---|---|---|
| 1 — Follow Up Now | Positive, Has a Question, Application Submitted | Reply with next steps / answer questions / add to tracker |
| 2 — Special Action Needed | No Longer Works There, Declined | Read Claude summary — may need to find new contact, post to board, or resend directly |
| 3 — Next Blast | No Reply, Auto-Reply | No action now — include in cleaned listserv |
| 4 — Remove from List | Bounce / Delivery Failure | Run clean_listserv.py |

## clean_listserv.py output

Produces `output/cleaned_listserv.xlsx` with three sheets:
- **Cleaned Listserv**: addresses with confirmed bounces removed — use for next blast
- **Needs Review**: No Longer Works There + Declined — manual decision required per row
- **Removed (Bounces)**: confirmed dead addresses, for records

## positive_replies_report.py output

Produces `output/positive_replies_outreach.html`. Each card contains:
- Sender name, email, institution, date
- Claude's summary of their reply
- Draft thank-you email asking if they know other **institutional/department email addresses** to add to the listserv

Open in Chrome → Ctrl+P → Save as PDF.

## Dependencies

All pinned in `requirements.txt`. Key packages:
- `anthropic` — Claude API client
- `google-api-python-client`, `google-auth-oauthlib`, `google-auth-httplib2` — Gmail API
- `openpyxl` — Excel export

## Known behaviour notes

- `*-owner@lists.*` senders are now auto-tagged as Bounce / Delivery Failure (not sent to Claude).
- Out-of-office auto-replies are classified by Claude as **Auto-Reply** (separate from hard bounces).
- ~700 No Reply rows out of 815 listserv addresses is normal — ~85% non-response is expected for cold academic blasts.
- `get_sent_recipients()` checks the Sent folder for the blast, but Gmail does not expose BCC in sent copies — the function always returns an empty set; the Unverified category is effectively unused.
- The `institution` column is extracted from the sender's email domain; free domains (gmail.com, yahoo.com, etc.) trigger a regex scan of the email body instead.
