# Power BI guide

Power BI Desktop is free. This guide covers three things:

- pointing a report at the demo workbooks
- what each page contains
- how to share the work without a paid licence

> **A `.pbix` file stores a full copy of its data.** Never share or commit a report that was last refreshed on real data. The `.gitignore` blocks every `.pbix` except `powerbi/*_DEMO.pbix`. Commit a file under that name only after it has been refreshed on the demo data and checked against the [privacy checklist](#privacy-checklist-before-you-share-a-pbix).

## 1. Create the demo data

Run the demo in its own folder, never in the folder of a live installation:

```bash
python demo/run_demo.py --end-date 2026-10-12
```

The workbooks are written to `data\database\` inside the repository.

For the screenshots, use `python demo/run_demo.py --full` instead. Its history starts on 1 April of the previous financial year. If your Date table uses fixed dates (`CALENDAR(DATE(2026,4,1), …)`), move the start back to that date, or use `CALENDARAUTO ( 3 )`, otherwise last year's bookings and surveys are cut off.

## 2. Point the report at the demo data

### If you already have the report (the author's path)

1. Open the production report. Choose **File → Save as** and name the copy `powerbi\Sales_Dashboard_DEMO.pbix`. Work only in the copy from now on.
2. Choose **Home → Transform data → Manage parameters → New parameter**.
   - Name: `DataFolder`
   - Type: **Text**
   - Current value: the full path of `data\database\`, ending with a backslash. For example: `C:\Projects\dealership-sales-pipeline\data\database\`
3. Open every query in **Advanced Editor**. Replace the fixed path with the parameter:

   ```
   File.Contents("E:\...\data\database\master_bookings.xlsx")      // before
   File.Contents(DataFolder & "master_bookings.xlsx")              // after
   ```

   **Home → Data source settings → Change source** works too, one file at a time. The parameter is better, because it switches all queries at once.
4. Choose **Close & Apply**, then **Refresh**. Every page now shows Demo Motors.
5. Look for text that was typed by hand and still refers to the real business, then work through the [privacy checklist](#privacy-checklist-before-you-share-a-pbix).

### If you are building the model from scratch

Create one query per table with this pattern. Use *Home → New source → Blank query → Advanced editor*:

```
let
    Source   = Excel.Workbook(File.Contents(DataFolder & "run_log.xlsx"), null, true),
    Sheet    = Source{[Item="Daily_Runs", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Typed    = Table.TransformColumnTypes(Promoted, {{"Run_Date", type date}, {"Files_OK", Int64.Type}, {"Ready_Time", type text}})
in
    Typed
```

| Query | Workbook | Sheet |
|---|---|---|
| Bookings, Stage_History, Load_Log, DQ_Issues | master_bookings.xlsx | same name |
| Open_Now, Open_Trend, Pipeline_Exits, Open_Issues | open_pipeline.xlsx | same name |
| SSI_Surveys, SSI_Answers | ssi.xlsx | same name |
| Team_Map | team_map.xlsx | Team_Map |
| Daily_Runs, File_Loads, Data_Issues, Data_Freshness | run_log.xlsx | same name |

Then build the model:

- **Set every column type yourself.** Power Query guesses types from the first 200 rows only.
  - `SSI_Answers[Rating]` is blank for the first questions, so it is guessed as text, and `AVERAGE` then fails. Set it to **Decimal number**.
  - Mixed cells such as `Ready_Time` (`"10:00"` or empty) must stay **Text**.
- **Add `CA_Key`** to Bookings and Open_Now: *Add column → Custom column* with `[Employee Full Name] & "|" & [Vehicle_Type]`.
- **Add `Box_Order`** to SSI_Surveys as a *Conditional column* in Power Query: Top Box 1, Middle Box 2, Bottom Box 3, Bottom Box Critical 4, otherwise 9. Then sort `Box Category` by it. A DAX column here creates a circular dependency.
- **Create a Date table**, for example `Date = CALENDARAUTO ( 3 )` (the financial year ends in March), with a Month column for the slicer.
- **Create relationships** as described in [data-model.md](data-model.md#power-bi-model). Delete any relationship that Power BI detected on its own.
- **Set the sort-by columns:**
  - `File_Loads[File_Label]` by `File_Order`
  - `Data_Issues[Severity]` by `Severity_Order`
  - `Data_Freshness[Source]` by `Source_Order`
  - `Open_Now[Age_Band]` by `Age_Band_Order`
  - `Follow_Up` by `Follow_Up_Order`
- **Add the measures** from [dax-measures.md](dax-measures.md).

## 3. The pages

The report has 12 pages: 6 in the top menu and 6 detail pages that open from a card or tile. Every page has the same navy header, page buttons and a "Data up to" box. Colours mean one thing everywhere:

- orange: SSI below 930
- green: SSI of 980 or more
- red: late or missing

| # | Page | Opened from | What is on it |
|---|---|---|---|
| 1 | Home | menu | "Choose your view": one button per role, plus the data date |
| 2 | Overview | menu | Five cards: bookings this month vs the same days last month, projected month end, sales closed, waiting for delivery, SSI latest month with year to date. Also bookings by model (PV/EV), a branch table with a weekly sparkline, cumulative bookings this month vs last month, and four "needs attention" tiles |
| 3 | Detail · Waiting customers | Overview tile | Typical and longest wait, the call list sorted by days waiting, and a model → variant → colour matrix by 31–60 / 61–90 / 90+ days |
| 4 | Detail · CRM data issues | Overview tiles | Issue tiles, bookings to correct in the CRM, and backward stage moves |
| 5 | Sales Pipeline | menu | Enquiries (new in 7 days, over 30 days), hot / warm leads, test-drive %, booked but not delivered, enquiry age bands, the pipeline by branch and advisor, and four follow-up tiles |
| 6 | Pipeline Follow-up | Sales Pipeline tiles | List picker (Follow_Up), SM / TL / branch / PV-EV slicers, and the list sorted by days open |
| 7 | Team Performance | menu | SM and TL slicers; active CAs, bookings per CA, CAs with no booking, cross-sales; bookings by SM (PV/EV); a scorecard SM → TL → CA |
| 8 | Customer Experience | menu | SSI score, surveys, % at target, customers under 930, Top Box %, average rating, penalty per survey; the trend for PV vs EV with 980 / 930 lines; box mix; SSI by team (worst first); where points are lost |
| 9 | SSI Call-back List | Customer Experience card | Customers under 930, lowest first (red under 800), with the weakest question and the "No" answers |
| 10 | Data Operations | menu | Last run status and ready time, SLA %, files OK today, issues to fix; the file-arrival grid for the last 14 days; data freshness; daily runs; the data-issues matrix |
| 11 | Data Issue Details | Data Operations card, or right-click a check → Drill through | Every issue with *where to fix it* |
| 12 | Detail · SLA history | Data Operations SLA card | SLA met % by month (`Ops SLA Met %` by `Daily_Runs[Run_Month]`): every month since the bot started, six months in the `--full` demo |

## 4. Things that cost time (and the fix)

| Symptom | Cause | Fix |
|---|---|---|
| A card shows an error on `AVERAGE` | The column was loaded as text (blank first 200 rows) | Set the type in Power Query |
| Conditional colours turn everything orange or red | The rule boxes were set to *Percent* | Use **Number** |
| A colour rule behaves oddly | The summarization was *Standard deviation* | Use **Average** or **Sum** |
| A card shows "53K" instead of a count | It summed a score column | Write a `COUNTROWS` measure |
| A list shows only one month though no visible slicer is set | A **hidden** slicer on that page is filtering it | Open **View → Selection pane**, show all objects and check every slicer, including in **Sync slicers** |
| Drill-through says "No available actions" | The newer on-object layout sets it elsewhere | Go to **Format page → Page information → Page type = Drillthrough, Drill through from = the field**. Then right-click a row of that field |
| Sorting by a DAX column gives a circular dependency | The sort column depends on the column it sorts | Create the order column in Power Query |
| "Latest month" and "year to date" differ | Both are right, but the label did not say which is which | Put the period in the card label ("Oct 2026 (14 surveys) · YTD …") |
| The Month slicer starts in April of this year, although older data is loaded | The Date table has a fixed start date (`CALENDAR ( DATE ( 2026, 4, 1 ), … )`) or is built from a table that starts this year | Move the start back (`DATE ( 2025, 4, 1 )`) or use `CALENDARAUTO ( 3 )`, then check the slicer's own filters |
| "YTD" includes last year once two financial years are loaded | The measure took every row, which was the same as year to date while only this year was loaded | Filter from 1 April of the latest month's financial year (see `SSI Score YTD` in [dax-measures.md](dax-measures.md)) |

## 5. Sharing without a paid licence

| Option | Cost | What viewers get |
|---|---|---|
| Screenshots and a short screen recording in the README / LinkedIn | free | The fastest way to show the work. Use the demo data only |
| **File → Export → PDF** from Power BI Desktop | free | Every page as a PDF, which you can attach to the repository or a post |
| Commit `powerbi/Sales_Dashboard_DEMO.pbix` | free | Anyone can open it in free Power BI Desktop and click through it |
| Publish to web (public link) | Needs a Power BI account under a work or school e-mail, and the tenant setting must allow it | A live link that anyone can open. **Never** use your employer's tenant for a personal showcase |
| Portfolio sites that host Power BI | Usually need a Publish-to-web link | Same as above |

## Privacy checklist before you share a pbix

- [ ] Every query uses `DataFolder`, and the parameter points at the **demo** folder. You have refreshed, and there were no errors.
- [ ] In **Data view**, open every table: there are no real names, phone numbers, IDs, chassis numbers or branches.
- [ ] Tables typed in by hand (**Enter data**: targets, a branch list, a lookup) hold no real values.
- [ ] Text boxes, page titles, tooltips, bookmarks, buttons and images (logos) show no company name.
- [ ] Visual-level filters that name real people or branches are removed (they would show blank on demo data anyway).
- [ ] The theme name and the file properties (**File → Info**) do not name the company.
- [ ] The parameter value does not reveal a private folder path. Set it to a neutral path before saving.
