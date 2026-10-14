# Case study: automating a dealership's daily sales reporting

*Role: MIS / sales operations analyst, building alone, end to end: Python, data model, Power BI, operations.*
*The business is anonymised. Every number and screenshot in this repository comes from fictional demo data.*

## The situation

A multi-branch car dealership sells passenger and electric vehicles. Its only source of sales data was the manufacturer's CRM, which e-mails six Excel reports every morning:

- bookings for the month
- the open pipeline
- customer-satisfaction surveys

Each report covers PV and EV separately.

Several problems followed from this:

- **No history.** Bookings drop out of the "current month" file when the month ends. Open records vanish when they are delivered or lost, so nobody could say how long customers had waited, or how many leads had been lost.
- **Manual consolidation.** The reports were copied into spreadsheets by hand every day. Numbers differed depending on who prepared them.
- **Unclear credit.** Some CRM logins are shared between consultants, so bookings and survey scores landed on the wrong person.
- **Invisible data problems.** Duplicate bookings, missing dates, unassigned leads and typed-twice customer names went straight into the totals.

## What I built

1. **A mail bot.** It reads the reports from the mailbox (read-only), waits until every required file has arrived, and loads each file through its own handler. It then sends a GREEN or RED health e-mail with file-by-file status, data-quality counts and the SLA (data ready by 13:00).
2. **A history-keeping database in five Excel workbooks.**
   - Bookings are upserted by opportunity ID, with stage history.
   - A daily pipeline snapshot records trend and exits.
   - Surveys are kept question by question.
   - A team map runs CA → TL → SM, with history.
   - An operations log records every run.
3. **A 12-page Power BI report** with one entry page per role: directors, sales managers, team leaders, customer experience, and sales operations. It has 150+ DAX measures and drill-through call lists.
4. **A synthetic replica.** A generator produces fictional reports in the exact CRM layouts, and the unmodified bot replays them at simulated times. This made it possible to test, demo and publish the system without exposing any real data.

## Problems worth telling

| Problem | What I did |
|---|---|
| The CRM export has **several columns called `Division`**, in varying positions | The bot picks the column by its *content* (pattern `NNNNNNN-Sales-<Branch>-<Dealer>`), not by position or name |
| One morning, the **survey penalty points for EV were halved** after a file loaded without any error. The manufacturer had placed two questionnaires side by side, with score and answer columns interleaved | I found it by reconciling the points against the previous day. Columns are now paired by **question text**, a one-letter typo is merged, and a *layout check* warning goes into the health mail if a score column has no answer column. I reloaded the day and the points matched again |
| A "No" to "Is anything pending?" was counted as a **failure** | Information-only questions are now typed separately, so "No" is the good answer there |
| **Shared CRM logins** credited the wrong consultant | A credit chain: the sales register by chassis → the login-ID list → the CRM owner. Every survey says which rule was used (`CA_Source`), so the result can be audited |
| Fuzzy name matching **merged two different people** with similar surnames | Strict matching, word by word, with a whole-name threshold. One-word names are never matched by spelling. Every match is listed for review |
| The team list said EV team = EV sales, which **would have hidden 20% of bookings** made across teams | I separated the CA's team (`Team_Type`) from what was sold (`Vehicle_Type`). Cross-sales became a metric instead of missing data |
| The laptop **slept through 08:30**, woke up without Wi-Fi, and the bot gave up | The bot now waits for the network, retries every 10 minutes until 20:00, keeps Windows awake while it runs, warns about a late start, and sends a RED crash mail on unexpected errors. I also documented the Task Scheduler and power settings |
| The health mail said "2 of 2 files received" when the **mailbox could not be reached** | That case now says "not checked (mailbox could not be reached)". It never pretends |
| SLA % looked worse than reality (3 days of history), and reruns **moved the ready time** | Ready time = the first GREEN mail of the day. A rerun never moves it, and a day recovered on a later date counts as a miss. The history was rebuilt from the saved mails and the raw folder |
| **Replaying the demo** showed that a day which was RED at the 13:00 cut-off was "pending", not "missed" | Fixed, with a regression test. The simulated replay paid for itself |
| The laptop woke up late in the evening, Task Scheduler started its missed run, and a manual run was started in the same minute. **Both copies wrote the master workbook through the same temp file**, and it was damaged (`BadZipFile`) | Reproduced on a copy of the data by starting two copies a fraction of a second apart. Fixed in three layers: (1) a lock, so a second copy stops at once; (2) a temp file per process that is read back before it replaces the workbook; (3) dated daily backups of every workbook, plus a `check_workbooks.py --restore` tool that puts back the newest good copy and prints the days to reload. The restore took one command, and nothing was lost |
| A survey whose seller is in no list got the team key `nan|PV`, so Power BI showed a "(Blank)" team | The key is now `SELLER NOT FOUND|PV`, and the team list gives it its own row ("Seller not found (SSI)"). A test checks that every survey has a team row |
| Power BI: a card failed, colours were all wrong, a list showed one month only | Explicit column types (Power Query guesses from 200 rows); Number instead of Percent in colour rules; a **hidden slicer** found in the Selection pane |

## Results

- The daily load runs unattended. The team gets a GREEN or RED mail before the cut-off instead of a spreadsheet late in the day.
- One version of the numbers for every role. Each page answers one question: how the month is going, who to call, which team needs help, which customers to call back, and whether the data can be trusted.
- Data quality became visible and actionable. Every issue says *where to fix it* (CRM, team list, ID list), and the counts are tracked daily.
- Examples of what the dashboard surfaced:
  - most long waits for delivery came from one EV model
  - long EV handovers cost the most survey points
  - many open enquiries had no lead class and no test drive

## What I would do next

- Move the storage from Excel workbooks to SQLite or PostgreSQL once the volume grows. The handlers already isolate the reading and writing.
- Credit bookings, not only surveys, to the actual seller.
- Add invoicing and retail data, and monthly targets per team.
- Use scheduled refresh in the Power BI service, if the business licenses it.

## Skills shown

Python (pandas, openpyxl, IMAP/SMTP), data cleaning and validation, idempotent incremental loads, schema-drift handling, entity matching, operational monitoring and alerting, test design with synthetic data and a simulated clock, CI on Linux and Windows, Power Query, DAX, star-schema modelling, dashboard design for different roles, and documentation for non-technical users.
