# tov-learn-feedback

Tool for improving the course **Tov-Learn** (an AI course teaching how/when to use AI — Claude — for job and business needs, delivered conversationally via Claude Code/Desktop) using participants' feedback.

This file is auto-loaded by Claude Code at the start of any session in this folder. Keep it updated so a session on either computer has full context without needing to be told manually.

## What this project does

`feedback_tool.py` collects, validates, and analyzes participant feedback on Tov-Learn using 5 fixed Hebrew questions, via Gemini (`gemini-3.5-flash-lite`).

**Pipeline:**
1. Participants fill a Google Form → Sheet 1 ("responses") of a shared Google Sheet.
2. `python feedback_tool.py sync` — reads new/changed responses (deduped by email), plus the local seed feedback in `feedbacks.csv` (index 0 = the project author's own feedback). For each answer: **validate** (is it specific enough? if not, generate a Hebrew follow-up question) then **analyze** (summarize into one-bullet-per-claim). Writes results to Sheet 2 ("analysis"). Skips unchanged rows unless `--force-index N` is passed (works for both Google-form rows and the local index-0 row).
3. `python feedback_tool.py aggregate` — merges all analyzed rows from Sheet 2 into a "סיכום" (summary) tab: per-question and per-category (positive / negative-improvements / new-ideas). Must include every distinct claim (even ones only one participant raised, just without a participant count), ordered by descending popularity, without ever attributing one entity's trait/limitation to a different entity.
4. `validate`/`analyze` commands (no `--all` needed) work directly against `feedbacks.csv` for local-only inspection.

**Files:**
- `feedback_tool.py` — the whole tool (prompts, Sheets I/O, Gemini calls, CLI).
- `init_feedbacks.py` — one-time script that created `feedbacks.csv` with the author's own seed feedback (row index 0). Already run; don't rerun.
- `feedbacks.csv` — local feedback data (currently just index 0). Contains raw feedback text — treat as sensitive/PII, gitignored.
- `sync_state.json` — tracks last-synced answers per participant (by email) and per local row (`_local_0`), to skip reprocessing unchanged answers. Contains participant emails + feedback — gitignored.
- `.env` — `GEMINI_API_KEY` only. Gitignored, never commit.
- `service_account.json` — Google service account credentials (Sheets scope). Gitignored, never commit.
- `.gitignore` — excludes `.env`, `service_account.json`, `sync_state.json`, `feedbacks.csv`.
- `output_notes.txt`, `outputs.txt` — old scratch/output files, not part of the current pipeline.

**External systems (not secret, safe to keep here):**
- Google Sheet ID `14mdsbX_aHo29xJhcsGU52lcw5tOvhb7i2anwA5dXdEY` — Tab 1 = responses (gid `2076837780`), Tab 2 = "analysis" (gid `2047537741`), Tab 3 = "סיכום" (created on demand).
- Service account client email: `tov-learn-feedback@valid-song-508806-a2.iam.gserviceaccount.com` (GCP project `valid-song-508806-a2`).
- **Actual secrets** (the private key in `service_account.json`, the key in `.env`) are NOT in this file and should never be added here — this file may end up copied more widely (USB, Drive) than those two files.

## Working across two computers

This project is moved between machines via USB drive (whole folder, including `.env` and `service_account.json` — keep those off Google Drive/cloud sync; physical transfer is the safer path for actual credentials). This `CLAUDE.md` travels with the folder either way, so a session on either machine has full context automatically.

## Progress log

- **2026-09-20**: Created local `.env` (GEMINI_API_KEY only) and `.gitignore` (excludes secrets + PII files); pointed `feedback_tool.py`'s `load_dotenv` at the local `.env` instead of the shared `AI-exercises/.env`.
- **2026-09-24**: Added `read_local_feedback()` + wired local `feedbacks.csv` rows into `sync_command()` (with `--force-index` support) so the author's own feedback (index 0) flows into the Google Sheet analysis tab and therefore into `aggregate`, not just Google Form participants. Strengthened `MERGER_SYSTEM`/`CATEGORY_MERGER_SYSTEM` prompts to never drop singleton claims and to order all claims (counted and uncounted) by descending popularity in one list.
- **2026-09-24 (later)**: Fixed three real issues found in the generated summary: (1) a stale "nothing to report" entry (participant answered "לא היו קשים") that predated the nothing-to-report filter — force-reprocessed so it's now correctly excluded; (2) `ANALYZER_SYSTEM` was misattributing a participant's personal trait (communication difficulties) as the reason the AI couldn't answer — added an explicit rule + worked example keeping personal traits and AI's structural context-limitation as separate claims; (3) `ANALYZER_SYSTEM` was bundling multiple unrelated technical complaints (screen self-study, script→pptx translation, reversed Hebrew in terminal, tool preference) into one bullet — added a rule + example for splitting unrelated bundled issues. Also had to add a rule against the merge step silently dropping detail when folding a unique claim into a more general one, and against leaking participant-number labels (e.g. "ציין משתתף 2") on singleton claims, which broke the existing no-attribution convention.
- **2026-10-02**: Moved to the second computer (Windows user `ronma`). Installed deps (`pip install pandas gspread google-auth google-genai python-dotenv`). Replaced hardcoded `C:\Users\noama\...` paths for `.env`, `service_account.json`, `sync_state.json` with paths relative to the script (`BASE_DIR`), so the folder works on either machine. Verified with read-only tests: `validate` (Gemini OK) and `sync --list-columns` (Sheets connection OK, no writes).
- **2026-10-02 (later)**: Replaced the validation criterion. Before, the per-answer prompt asked only "does it contain a concrete example?", which overrode the system prompt. Now `VALIDATOR_SYSTEM` treats an answer as valid if it has an identifiable subject (general is fine: "the quizzes", "the pace"), a real claim (the question itself can supply it, e.g. naming a subject under "what worked well?"), and is understandable. Examples are a plus, not a requirement. Only answers with no identifiable subject, no claim, or no intelligible meaning get a follow-up. Both call sites now share `validate_answer()`, which also returns a `reason`. Tested on 15 sample answers (all matched). The 6 cells the old rule flagged in Sheet 2 (participant indices 1, 2, 7) all pass under the new rule, but Sheet 2 still shows them flagged until `sync --force-index` is run for 1, 2, 7.
- **2026-10-03**: Ran `sync --force-index` for 1, 2, 7. All 6 cells passed validation. This exposed an older bug: `_is_nothing_to_report()` matched keyword fragments ("לא היה", "לא היו"…) anywhere in short answers, so negatively phrased real complaints were filtered out of the aggregate as "nothing to report". Participant 7's "ההסברים לא היו עמוקים מספיק… דילג על תירגולים" had been silently missing, and participant 2's "ה-JSON לא היה ממש ברור" was newly misfiled. Removed `_NOTHING_MARKERS`/`_is_nothing_to_report`. The validator now returns `nothing_to_report` and decides by meaning. Re-synced 2 and 7, and both complaints are now analyzed claims. Participant 4's real "לא היו קשים" is still correctly marked "—". `aggregate` has not been re-run yet, so the סיכום tab is stale. Sheet 2 backup taken before the re-sync is in the session scratchpad only, not in the repo.
